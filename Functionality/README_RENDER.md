# Functionality API 部署到 Render

一个 Web 服务提供 4 个接口，对应前端 index12–index15：

| 路径 | 功能 | 前端页 |
|------|------|--------|
| `POST /api/purchasing` | 采购 Excel 转 JSON | index12 采购信息 |
| `POST /api/sku-pescar` | 库存导 Excel 为 JSON | index13 |
| `POST /api/export-excel-json/preview` | 新品 Nx 获取可导日期（表头、日期列、日期列表） | index14 第一步 |
| `POST /api/export-excel-json` | 新品 Nx 按所选日期导出 JSON（form: file + dates[]） | index14 第二步 |
| `POST /api/empalagem` | 包裹尺寸 | index15 |

## 部署步骤

1. 在 Render 新建 **Web Service**，连接本仓库。
2. **Root Directory** 填：`Functionality`（重要）。
3. **Build Command**：`pip install -r requirements.txt`
4. **Start Command**：`gunicorn -w 1 -b 0.0.0.0:$PORT --timeout 120 app:app`
5. 部署完成后得到根 URL，例如：`https://functionality-api-xxx.onrender.com`
6. 在 index12–index15 的页面里，把 API 基础地址改为该 URL（例如 `API_BASE = 'https://functionality-api-xxx.onrender.com'`）。

## 免费版说明

- 单服务、单进程即可；4 个功能共用同一进程。
- 上传文件建议控制在 50MB 以内（服务端已限制 `MAX_CONTENT_LENGTH`）。
