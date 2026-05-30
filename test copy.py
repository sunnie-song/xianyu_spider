from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import os
import json
import time

def safe_get(data, *keys, default="暂无"):
    for key in keys:
        try:
            data = data[key]
        except (KeyError, TypeError, IndexError):
            return default
    return data

# cpvLabels 中 propertyName → 输出字段名 的映射
CPV_FIELD_MAP = {
    "品牌": "品牌",
    "型号": "型号",
    "成色": "成色",
    "速别": "速别",
    "车架材质": "车架材质",
    "车把类型": "车把类型",
}

def extract_item_id(link):
    """从商品链接中提取 itemId"""
    import re
    match = re.search(r'[?&]id=(\d+)', link)
    return match.group(1) if match else None

def fetch_detail_attributes(page, item_id):
    """
    通过详情页 API 获取商品结构化属性
    直接导航到详情页触发原生 API 请求
    返回 dict，key 为 CPV_FIELD_MAP 中的输出字段名
    """
    result = {v: "暂无" for v in CPV_FIELD_MAP.values()}
    if not item_id:
        return result

    detail_page = None
    try:
        detail_page = page.context.new_page()
        detail_data = {}

        def on_response(response):
            url = response.url
            if 'idle.pc.detail' in url and not detail_data:
                try:
                    ct = response.headers.get('content-type', '')
                    if 'json' in ct:
                        body = response.json()
                        if 'itemDO' in body.get('data', {}):
                            detail_data['body'] = body
                except:
                    pass

        detail_page.on("response", on_response)

        # 导航前先拦截弹窗
        detail_page.route("**/*", lambda route: route.continue_())
        
        detail_page.goto(f"https://www.goofish.com/item?id={item_id}", timeout=20000, wait_until="domcontentloaded")
        detail_page.wait_for_timeout(1500)

        # 立即隐藏所有弹窗
        detail_page.evaluate("""
            const hideAll = () => {
                document.querySelectorAll('.ant-modal-wrap, .ant-modal-mask, [class*="login"], [class*="dialog"]')
                    .forEach(el => { el.style.display = 'none'; el.style.pointerEvents = 'none'; });
            };
            hideAll();
            setInterval(hideAll, 500);
        """)

        # 等待详情 API
        for _ in range(20):
            if detail_data:
                break
            time.sleep(0.5)

        if detail_data:
            cpv_labels = safe_get(detail_data['body'], 'data', 'itemDO', 'cpvLabels', default=[])
            print(f"    cpvLabels: {len(cpv_labels)} found")
            for label in cpv_labels:
                prop_name = safe_get(label, 'propertyName', default='')
                value_name = safe_get(label, 'valueName', default='')
                if prop_name in CPV_FIELD_MAP:
                    result[CPV_FIELD_MAP[prop_name]] = value_name
                    print(f"      {prop_name}: {value_name}")
        else:
            print(f"    (详情API未触发)")

    except Exception as e:
        print(f"    获取详情失败 (itemId={item_id}): {e}")
    finally:
        if detail_page:
            detail_page.close()

    return result

def save_to_excel(data_list, filename="商品数据.xlsx"):
    if not data_list:
        print("没有需要保存的数据")
        return

    # 创建DataFrame并去重
    df = pd.DataFrame(data_list).drop_duplicates(subset=["商品链接"], keep="first")

    # 设置输出路径（桌面）
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filepath = os.path.join(desktop, filename)

    # 使用xlsxwriter引擎
    writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='商品列表')

    # 获取工作表对象
    workbook = writer.book
    worksheet = writer.sheets['商品列表']

    # 设置标题格式
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#4F81BD',
        'font_color': 'white',
        'border': 1
    })

    # 设置列宽
    col_widths = {
        '商品标题': 50,
        '当前售价': 15,
        '品牌': 15,
        '型号': 12,
        '成色': 12,
        '速别': 10,
        '车架材质': 12,
        '车把类型': 12,
        '发货地区': 15,
        '卖家昵称': 20,
        '商品链接': 60,
        '商品图片链接': 60,
        '发布时间': 20,
    }

    for col_num, (col_name, width) in enumerate(col_widths.items()):
        if col_name in df.columns:
            idx = list(df.columns).index(col_name)
            worksheet.set_column(idx, idx, width)

    # 设置标题行格式
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    writer.close()
    print(f"数据已保存到：{filepath}")

def hide_login_modal(page):
    """隐藏登录弹窗"""
    page.evaluate("""
        const selectors = [
            '.ant-modal-wrap', '.ant-modal-mask',
            '[class*="login-modal-wrap"]', '[class*="login-modal"]',
            '#alibaba-login-box',
        ];
        selectors.forEach(sel => {
            document.querySelectorAll(sel).forEach(el => {
                el.style.display = 'none';
                el.style.visibility = 'hidden';
                el.style.pointerEvents = 'none';
                el.style.zIndex = '-9999';
            });
        });
    """)

def scrape_xianyu(keyword, max_pages=1):
    data_list = []

    def on_response(response):
        if "h5api.m.goofish.com/h5/mtop.taobao.idlemtopsearch.pc.search" in response.url:
            try:
                result_json = response.json()
                items = result_json.get("data", {}).get("resultList", [])

                for item in items:
                    main_data = safe_get(item, "data", "item", "main", "exContent", default={})
                    click_params = safe_get(item, "data", "item", "main", "clickParam", "args", default={})

                    title = safe_get(main_data, "title", default="未知标题")

                    # 价格处理
                    price_parts = safe_get(main_data, "price", default=[])
                    if isinstance(price_parts, list):
                        price = "".join([str(p.get("text", "")) for p in price_parts if isinstance(p, dict)])
                    else:
                        price = "价格异常"
                    if "当前价" in price:
                        price = price.replace("当前价", "")
                    if "万" in price:
                        price = "¥" + str(int(float(price.replace("¥", "").replace("万", "")) * 10000))

                    area = safe_get(main_data, "area", default="地区未知")
                    seller = safe_get(main_data, "userNickName", default="匿名卖家")

                    raw_link = safe_get(item, "data", "item", "main", "targetUrl", default="")
                    clean_link = raw_link.replace("fleamarket://", "https://www.goofish.com/")

                    image_url = safe_get(main_data, "picUrl", default="")
                    if image_url and not image_url.startswith("http"):
                        image_url = f"https:{image_url}"

                    publish_time = safe_get(click_params, "publishTime", default="")
                    publish_date = "未知时间"
                    if publish_time.isdigit():
                        try:
                            dt = datetime.fromtimestamp(int(publish_time) / 1000)
                            publish_date = dt.strftime("%Y-%m-%d %H:%M")
                        except:
                            pass

                    data_list.append({
                        "商品标题": title,
                        "当前售价": price,
                        "发货地区": area,
                        "卖家昵称": seller,
                        "商品链接": clean_link,
                        "商品图片链接": image_url,
                        "发布时间": publish_date,
                        "品牌": "暂无",
                        "型号": "暂无",
                        "成色": "暂无",
                        "速别": "暂无",
                        "车架材质": "暂无",
                        "车把类型": "暂无",
                    })

            except Exception as e:
                print(f"数据处理异常: {str(e)}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path='/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            headless=False
        )
        page = browser.new_page()

        try:
            # 访问闲鱼首页
            page.goto("https://www.goofish.com")
            page.wait_for_timeout(3000)

            # 关闭广告弹窗
            try:
                page.click("[class*='closeIconBg']", timeout=2000)
            except:
                pass

            # 关闭登录弹窗
            for close_selector in [
                "[class*='login-modal'] [class*='close']",
                ".ant-modal-close",
                "[class*='modal-close']",
                "[aria-label='Close']",
            ]:
                try:
                    page.click(close_selector, timeout=1000)
                    print("已关闭登录弹窗")
                    break
                except:
                    pass

            hide_login_modal(page)
            page.wait_for_timeout(500)

            # 执行搜索
            print("正在搜索商品，请稍等...")
            page.fill('input[class*="search-input"]', keyword)
            page.click('button[type="submit"]', force=True)
            page.wait_for_timeout(3000)

            hide_login_modal(page)

            # 设置最新排序
            page.click('text=新发布', force=True)
            page.wait_for_timeout(1000)
            page.click('text=最新', force=True)

            # 注册响应监听
            page.on("response", on_response)

            # 分页爬取
            current_page = 1
            while current_page <= max_pages:
                print(f"正在处理第 {current_page}/{max_pages} 页...")

                try:
                    page.wait_for_selector("[class*='search-pagination-container']", timeout=1000)
                except:
                    print("未找到分页器，可能只有单页结果")
                    break

                try:
                    next_btn = page.query_selector("[class*='search-pagination-arrow-right']")
                    if (next_btn is None) or ("disabled" in next_btn.get_attribute("class")):
                        print(f"无法找到第 {current_page + 1} 页，终止爬取")
                        break
                    page.click("[class*='search-pagination-arrow-right']", timeout=1000)
                except:
                    print(f"无法找到第 {current_page + 1} 页，终止爬取")
                    break

                current_page += 1

            # ========== 搜索完成，开始获取详情属性 ==========
            print(f"\n搜索完成，共获取 {len(data_list)} 条商品，开始获取详情属性...")

            for i, item in enumerate(data_list):
                if i >= 1:
                    # 测试模式：只获取第一条的详情属性
                    print(f"  [{i+1}/{len(data_list)}] 跳过（测试模式，仅取第一条详情）")
                    continue
                item_id = extract_item_id(item["商品链接"])
                if item_id:
                    print(f"  [{i+1}/{len(data_list)}] 获取详情: {item['商品标题'][:30]}...")
                    attrs = fetch_detail_attributes(page, item_id)
                    item.update(attrs)
                else:
                    print(f"  [{i+1}/{len(data_list)}] 跳过（无itemId）: {item['商品标题'][:30]}...")

            # 最终保存数据（测试模式：只保存第一条）
            test_data = data_list[:1] if data_list else []
            if test_data:
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                keyword_clean = ''.join(e for e in keyword if e.isalnum())
                filename = f"闲鱼爬取结果_{keyword_clean}_测试_{timestamp}.xlsx"
                save_to_excel(test_data, filename)
            else:
                print("没有采集到任何商品数据")

        except Exception as e:
            print(f"爬取中断: {str(e)}")
            if data_list:
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                save_to_excel(data_list, f"闲鱼_异常中断_{timestamp}.xlsx")
            page.screenshot(path='xianyu_error.png')
            print("错误截图已保存至目录")

        finally:
            browser.close()

if __name__ == "__main__":
    keyword = input("请输入要搜索的商品关键词：").strip()
    pages = int(input("请输入要爬取的最大页数：").strip() or 1)
    print(f"开始爬取 {keyword} 的前 {pages} 页数据...")
    scrape_xianyu(keyword, max_pages=pages)
    print("爬取任务结束")
