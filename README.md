# 闲鱼商品搜索API

[![FastAPI](https://img.shields.io/badge/FastAPI-0.68.0-009688?logo=fastapi)](https://fastapi.tiangolo.com/)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)](https://www.python.org/)

基于 FastAPI 构建的闲鱼商品搜索接口，支持异步并发请求和自动化数据去重存储。

## 功能特性

- 🔍 关键词商品搜索（支持分页）
- ⚡ 异步高性能爬取（Playwright 无头浏览器）
- 🛡️ 智能数据去重（基于链接特征哈希值）
- 💾 数据持久化存储（关系数据库）
- 📊 返回新增记录统计信息

## 技术栈

| 组件           | 用途                     |
|----------------|--------------------------|
| FastAPI        | RESTful API框架          |
| Playwright     | 浏览器自动化爬取         |
| Tortoise ORM   | 异步数据库ORM            |
| SQL            | 数据持久化存储           |
| Uvicorn        | ASGI服务器               |

## 快速开始

cd /Users/sunnie/Desktop/github/xianyu_spider && python3 test.py

### 环境配置

1. 安装依赖
```bash
pip install -r requirements.txt
playwright install chromium
```

2. 创建 `.env` 文件（请修改为自己的信息）
```env
DATABASE_URL=mysql://user:password@localhost/xianyu
```

### 启动服务
```bash
python spider.py
```

## API 文档

访问 `http://localhost:8000/docs` 查看交互式文档

### 搜索接口
```
POST /search/
```

**请求参数示例**：
```json
{
  "keyword": "手机",
  "max_pages": 1
}
```

**响应示例**：
```json
{
  "status": "success",
  "keyword": "手机",
  "total_results": 30,
  "new_records": 5,
  "new_record_ids": [101,102,103,104,105]
}
```

## 使用示例
建议使用 Apifox 或者 Postman 进行测试

### cURL 请求
```bash
curl -X POST "http://localhost:8000/search/" \
-H "Content-Type: application/json" \
-d '{"keyword": "笔记本电脑", "max_pages": 2}'
```

### Python 客户端
```python
import requests

response = requests.post(
    "http://localhost:8000/search/",
    json={"keyword": "数码相机", "max_pages": 3}
)
print(response.json())
```

## 注意事项

1. **法律合规**  
使用前请确保遵守《网络安全法》和闲鱼平台 Robots 协议，本代码仅用于学习研究

2. **反爬机制**  
建议配置代理 IP 池和随机请求间隔，默认配置可能触发反爬限制

3. **性能调优**  
- 调整数据库连接池配置（`pool_recycle`等参数）
- 建议生产环境部署时增加 Redis 缓存层

## 版权声明

本项目采用 [MIT License](LICENSE)，请合理使用并注明出处。数据抓取结果不得用于商业用途。