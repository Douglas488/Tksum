# Upseller 月营业额报表 API - Render 部署说明

本服务将 `Uptotal.py` 的汇总逻辑以 Web API 形式部署到 Render，供 index10（Upseller月营业额报表）页面调用。

## 功能说明

- **接口**: `POST /api/generate`
- **请求**: 表单字段 `file` 为 **ZIP 文件**，ZIP 内包含多个 `.xlsx` 文件（每个文件为一家店铺数据）。
- **Excel 要求**: 每个 xlsx 需包含列：`日期`、`总销售额`、`有效订单量`、`有效销售额`。
- **响应**: 返回汇总后的 Excel 文件 `Upseller月营业额报表.xlsx`（带店铺目录书签）。

## 部署步骤

1. 将包含 Upmensal 的仓库连接到 Render（或单独推送 Upmensal 仓库）。
2. 登录 [Render](https://render.com) → New → Web Service，连接你的仓库。
3. **重要**：若仓库根目录是项目根（例如 `EtiquetaFull-main`），必须在 Render 里设置 **Root Directory** 为 `Upmensal`，否则会报错 `Could not open requirements file: requirements.txt`。在 Dashboard → 该 Web Service → Settings → **Root Directory** 中填写 `Upmensal` 并保存。
4. 配置：
   - **Name**: 例如 `upmensal-uptotal-api`
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn -w 1 -b 0.0.0.0:$PORT --timeout 120 app:app`
5. 免费版会自动休眠，首次请求可能较慢（冷启动）。
6. 部署完成后得到 URL，例如：`https://upmensal-uptotal-api.onrender.com`。
7. 在 **index10.html** 中将 `API_BASE` 改为该 URL。

## 本地测试

```bash
cd Upmensal
pip install -r requirements.txt
# 准备一个 ZIP，内含多个 .xlsx
curl -X POST -F "file=@test.zip" http://127.0.0.1:5000/api/generate -o out.xlsx
```

启动本地服务：`python -m flask --app app run` 或 `gunicorn -b 127.0.0.1:5000 app:app`。
