from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import os
import re
import time
import json

def safe_get(data, *keys, default="暂无"):
    for key in keys:
        try:
            data = data[key]
        except (KeyError, TypeError, IndexError):
            return default
    return data

CPV_FIELD_MAP = {
    "品牌": "品牌",
    "型号": "型号",
    "成色": "成色",
    "速别": "速别",
    "车架材质": "车架材质",
    "车把类型": "车把类型",
}

def extract_item_id(link):
    match = re.search(r'[?&]id=(\d+)', link)
    return match.group(1) if match else None

def hide_login_modal(page):
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

def fetch_detail_attributes(context, item_id):
    """
    在同一 context 下打开详情页，拦截 detail API 获取 cpvLabels
    """
    result = {v: "暂无" for v in CPV_FIELD_MAP.values()}
    if not item_id:
        return result

    page = context.new_page()
    try:
        detail_data = {}

        def on_detail_response(response):
            if 'mtop.taobao.idle.pc.detail' in response.url and not detail_data:
                try:
                    detail_data['body'] = response.json()
                except:
                    pass

        page.on("response", on_detail_response)
        page.goto(f"https://www.goofish.com/item?id={item_id}", timeout=20000)

        # 持续隐藏登录弹窗，等待 API 返回
        waited = 0
        while not detail_data and waited < 15:
            time.sleep(0.5)
            hide_login_modal(page)
            waited += 0.5

        if detail_data:
            cpv_labels = safe_get(detail_data['body'], 'data', 'itemDO', 'cpvLabels', default=[])
            for label in cpv_labels:
                prop_name = safe_get(label, 'propertyName', default='')
                value_name = safe_get(label, 'valueName', default='')
                if prop_name in CPV_FIELD_MAP:
                    result[CPV_FIELD_MAP[prop_name]] = value_name

    except Exception as e:
        print(f"    [详情] 获取失败: {e}")
    finally:
        page.close()

    return result

def save_to_excel(data_list, filename="商品数据.xlsx"):
    if not data_list:
        print("没有需要保存的数据")
        return

    df = pd.DataFrame(data_list).drop_duplicates(subset=["商品链接"], keep="first")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filepath = os.path.join(desktop, filename)

    writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='商品列表')
    workbook = writer.book
    worksheet = writer.sheets['商品列表']

    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top',
        'fg_color': '#4F81BD', 'font_color': 'white', 'border': 1
    })
    col_widths = {
        '商品标题': 50, '当前售价': 15,
        '品牌': 15, '型号': 12, '成色': 12,
        '速别': 10, '车架材质': 12, '车把类型': 12,
        '发货地区': 15, '卖家昵称': 20,
        '商品链接': 60, '商品图片链接': 60, '发布时间': 20,
    }
    for col_num, (col_name, width) in enumerate(col_widths.items()):
        if col_name in df.columns:
            worksheet.set_column(list(df.columns).index(col_name), list(df.columns).index(col_name), width)
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    writer.close()
    print(f"数据已保存到：{filepath}")

def scrape_xianyu(keyword, max_pages=3):
    data_list = []

    # ====== 搜索响应处理器 ======
    def on_search_response(response):
        if "idlemtopsearch.pc.search" not in response.url:
            return
        try:
            result_json = response.json()
            items = result_json.get("data", {}).get("resultList", [])
            for item in items:
                main_data = safe_get(item, "data", "item", "main", "exContent", default={})
                click_params = safe_get(item, "data", "item", "main", "clickParam", "args", default={})

                title = safe_get(main_data, "title", default="未知标题")
                price_parts = safe_get(main_data, "price", default=[])
                price = "".join([str(p.get("text", "")) for p in price_parts]) if isinstance(price_parts, list) else "价格异常"
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
                        publish_date = datetime.fromtimestamp(int(publish_time) / 1000).strftime("%Y-%m-%d %H:%M")
                    except:
                        pass

                data_list.append({
                    "商品标题": title, "当前售价": price,
                    "发货地区": area, "卖家昵称": seller,
                    "商品链接": clean_link, "商品图片链接": image_url,
                    "发布时间": publish_date,
                    "品牌": "暂无", "型号": "暂无", "成色": "暂无",
                    "速别": "暂无", "车架材质": "暂无", "车把类型": "暂无",
                })
        except Exception as e:
            print(f"数据处理异常: {e}")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir='/tmp/xianyu_chrome_profile',
            executable_path='/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            headless=False,
        )
        page = context.new_page()
        seen_pages = 0  # 已处理的页数

        try:
            print("打开闲鱼首页...")
            page.goto("https://www.goofish.com")
            page.wait_for_timeout(3000)

            # 关闭弹窗
            try:
                page.click("[class*='closeIconBg']", timeout=2000)
            except:
                pass
            hide_login_modal(page)

            # 执行搜索
            print(f"搜索关键词: {keyword}")
            page.fill('input[class*="search-input"]', keyword)
            page.click('button[type="submit"]', force=True)
            page.wait_for_timeout(3000)
            hide_login_modal(page)

            # ====== 注册监听 + 清空初始搜索数据 + 切排序 ======
            page.on("response", on_search_response)
            data_list.clear()

            print("切换排序: 最新发布")
            page.click('text=新发布', force=True)
            page.wait_for_timeout(1000)
            page.click('text=最新', force=True)
            page.wait_for_timeout(3000)  # 等排序后的搜索结果返回
            seen_pages = 1

            # ====== 翻页 ======
            while seen_pages < max_pages:
                try:
                    next_arrow = page.query_selector("[class*='search-pagination-arrow-right']")
                    if next_arrow is None or "disabled" in (next_arrow.get_attribute("class") or ""):
                        print(f"已到最后一页（共 {seen_pages} 页）")
                        break
                    page.click("[class*='search-pagination-arrow-right']", timeout=3000)
                    page.wait_for_timeout(2000)
                    seen_pages += 1
                    print(f"  第 {seen_pages} 页完成")
                except:
                    print(f"翻页失败，共获取 {seen_pages} 页")
                    break

            print(f"\n搜索完成: {seen_pages} 页, {len(data_list)} 条商品")

            # ====== 获取详情属性 ======
            print("开始获取详情属性（品牌/型号/成色/速别/车架材质/车把类型）...\n")

            for i, item in enumerate(data_list):
                item_id = extract_item_id(item["商品链接"])
                if not item_id:
                    print(f"  [{i+1}/{len(data_list)}] 跳过（无itemId）")
                    continue

                print(f"  [{i+1}/{len(data_list)}] {item['商品标题'][:35]}...", end=" ", flush=True)
                attrs = fetch_detail_attributes(context, item_id)
                item.update(attrs)

                # 打印结果摘要
                filled = {k: v for k, v in attrs.items() if v != "暂无"}
                if filled:
                    print(f"✓ {filled}")
                else:
                    print("✗ 未获取到属性")

                time.sleep(1.5)  # 控制频率，避免被限

            # ====== 保存 ======
            if data_list:
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                keyword_clean = ''.join(e for e in keyword if e.isalnum())
                filename = f"闲鱼爬取结果_{keyword_clean}_共{seen_pages}页_{timestamp}.xlsx"
                save_to_excel(data_list, filename)
            else:
                print("没有采集到任何商品数据")

        except Exception as e:
            print(f"\n爬取中断: {e}")
            if data_list:
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                save_to_excel(data_list, f"闲鱼_异常中断_{timestamp}.xlsx")
            page.screenshot(path='xianyu_error.png')

        finally:
            context.close()

if __name__ == "__main__":
    keyword = input("请输入要搜索的商品关键词：").strip() or "迪卡侬RC100"
    pages = int(input("请输入要爬取的最大页数：").strip() or 3)
    print(f"开始爬取 {keyword} 的前 {pages} 页数据...")
    scrape_xianyu(keyword, max_pages=pages)
    print("爬取任务结束")
