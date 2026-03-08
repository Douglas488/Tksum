# 将 TkPy 部署到 Render

部署完成后，在 **index9.html** 中把 API 地址改为 Render 提供的地址（如 `https://tkpy-summary-api.onrender.com`），无需每月后台激活。

---

## 一、准备仓库

确保 TkPy 目录下包含：

- `app.py`
- `generate_summary.py`
- `requirements.txt`
- `render.yaml`（可选，用于 Blueprint）

若项目在 GitHub，将上述文件放在同一仓库（可放在仓库根目录或子目录如 `TkPy/`）。

---

## 二、在 Render 创建 Web Service

1. 登录 [Render](https://render.com)，进入 **Dashboard**。
2. 点击 **New +** → **Web Service**。
3. **Connect a repository**：选择你的 GitHub/GitLab 仓库；若未连接，先按提示授权并选择仓库。
4. 配置：
   - **Name**：例如 `tkpy-summary-api`（会得到 `https://tkpy-summary-api.onrender.com`）。
   - **Region**：选离用户较近的节点。
   - **Branch**：选要部署的分支（如 `main`）。
   - **Root Directory**：若 TkPy 在子目录（如 `TkPy`），填 `TkPy`；在根目录则留空。
   - **Runtime**：**Python 3**。
   - **Build Command**：  
     `pip install -r requirements.txt`  
     （若 Root Directory 为 `TkPy`，Render 会在该目录下执行，即安装 `TkPy/requirements.txt`。）
   - **Start Command**：  
     `gunicorn -w 1 -b 0.0.0.0:$PORT --timeout 120 app:app`  
     `$PORT` 由 Render 自动注入；`--timeout 120` 避免大文件处理被提前断开。
5. **Plan** 选 **Free**。
6. 点击 **Create Web Service**，等待首次构建与部署完成。

---

## 三、验证

- 浏览器打开：`https://你的服务名.onrender.com/`  
  应看到 JSON，且包含 `"generate_summary": "ok"`。
- 用 index9 或 curl 测试上传 Excel：
  ```bash
  curl -X POST -F "file=@你的文件.xlsx" https://你的服务名.onrender.com/api/generate -o out.xlsx
  ```

---

## 四、修改前端 API 地址

1. 打开项目中的 **index9.html**。
2. 将脚本里的 `API_BASE` 从 PythonAnywhere 改为 Render 地址，例如：
   ```javascript
   var API_BASE = 'https://tkpy-summary-api.onrender.com';
   ```
   （不要加末尾斜杠，且必须是 HTTPS。）
3. 若页面底部有「服务地址」或说明文案，可一并改为 Render 地址；若仍提示「免费服务器须每月激活」，可改为「本功能托管于 Render，休眠约 15 分钟无访问后会冷启动，首次打开可能需等待约 30 秒」。

---

## 五、免费版注意

- **休眠**：约 15 分钟无请求后服务会休眠，下次访问需冷启动（约 30 秒～1 分钟），属正常现象。
- **时长**：每月 750 小时免费，休眠期间一般不扣，够日常使用。
- **内存**：512 MB，单次处理一个 Excel 足够；若表格特别大可考虑缩小单次上传或升级套餐。

---

## 六、使用 render.yaml（可选）

若仓库根目录或 TkPy 目录下有 `render.yaml`，可在 Render Dashboard 选择 **New +** → **Blueprint**，连接同一仓库，Render 会按 `render.yaml` 创建服务；之后只需在 Dashboard 中确认 **Root Directory** 和 **Start Command** 与上面一致即可。
