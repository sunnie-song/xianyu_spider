# MEMORY.md - 项目长期记忆

## 项目概况
- **项目**：闲鱼爬虫（xianyu_spider）
- **路径**：`/Users/sunnie/Desktop/github/xianyu_spider/`
- **主要脚本**：`test.py`

## 环境配置
- Python 3.14.5，playwright、pandas、xlsxwriter（清华镜像）
- 系统 Chrome：`/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`
- 启动方式：`p.chromium.launch(executable_path='...', headless=False)`

## 搜索流程（正确顺序）
1. 打开 goofish.com，隐藏登录弹窗
2. 搜索关键词
3. 筛选"个人闲置"
4. 注册 `on_response` 监听 → 清空 `data_list`
5. 切排序"最新" → 触发 API 请求

## 字段数据源

| 字段 | 来源 | 备注 |
|------|------|------|
| 商品标题/售价/地区/昵称/链接/图片/发布时间 | `exContent` | 直接取值 |
| 服务标签 | `clickParam.args.serviceUtParams` | **值是 JSON 字符串**，需 `json.loads()`；文本在 `args.content` |
| 原价 (oriPrice) | 搜索 API **不存在** | 详情 API `cpvLabels` 中有，此字段当前用 `detailParams.soldPrice` 占位 |
| 品牌/型号/成色/速别/车架材质/车把类型 | 详情 API `cpvLabels` | 需额外请求详情接口 |

## 关键踩坑
- `serviceUtParams` 在 API 返回中是 JSON 字符串而非解析后的 list，必须 `json.loads()`
- 运行前必须关闭系统 Chrome：`osascript -e 'tell application "Google Chrome" to quit'`
- 排序点击加 `force=True` 防登录弹窗遮挡
