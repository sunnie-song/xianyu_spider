# MEMORY.md - 项目长期记忆

## 项目概况
- **项目**：闲鱼爬虫（xianyu_spider）
- **路径**：`/Users/sunnie/Desktop/github/xianyu_spider/`
- **主要脚本**：`test.py`（同步版本，输出Excel）；`spider.py`（FastAPI异步版，存MySQL）

## 环境配置

### Python 依赖
- Python 3.14.5（`/Library/Frameworks/Python.framework/Versions/3.14/bin/python3`）
- 已安装：playwright、pandas、xlsxwriter、openpyxl（通过清华镜像 `https://pypi.tuna.tsinghua.edu.cn/simple`）
- playwright 内置 Chromium 下载困难，改用系统 Chrome

### 浏览器配置
- 使用系统 Google Chrome：`/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`
- 启动方式：`p.chromium.launch(executable_path='...', headless=False)`

## 已知问题与解决方案

### 1. 登录弹窗遮挡
**问题**：访问 goofish.com 会弹出登录框（iframe 嵌入），遮挡页面操作  
**解决**：
1. 通过 JS 强制隐藏弹窗容器（`.ant-modal-wrap`、`.ant-modal-mask`、`[class*="login-modal-wrap"]`）
2. 点击操作加 `force=True` 参数

### 2. 排序设置
**问题**：需要点击"新发布" → "最新"才能获取最新数据  
**状态**：已处理，但也需要 `force=True`

## 用户偏好
- 代码风格：简洁、可读性强，减少不必要的导入
- 输出格式：Excel（桌面）
- 抓取关注点：公路车相关品类（迪卡侬RC100等）
