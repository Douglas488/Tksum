# 将 TkPy 托管到 PythonAnywhere 并供 index9 调用

## 一、在 PythonAnywhere 部署 TkPy

### 1. 注册与创建 Web 应用

1. 打开 [PythonAnywhere](https://www.pythonanywhere.com/) 并注册/登录。
2. 进入 **Web** 标签，点击 **Add a new web app**，选择 **Flask**，Python 版本选 3.10 或 3.11。
3. 记下你的站点地址，例如：`https://你的用户名.pythonanywhere.com`。

### 2. 上传代码

**方式 A：上传文件**

- 在 **Files** 标签下进入你的项目目录（如 `/home/你的用户名/TkPy`，需先建好该目录）。
- 上传本地的 `app.py`、`generate_summary.py`、`requirements.txt`（TkPy 目录下这三个文件务必都在）。

**方式 B：从 GitHub 克隆（若项目在 GitHub）**

- 在 **Consoles** 里打开 Bash，执行：
  ```bash
  cd ~
  git clone https://github.com/你的用户名/你的仓库.git
  cd 你的仓库/TkPy
  ```
  这样代码在 `~/你的仓库/TkPy/` 下。

### 3. 虚拟环境与依赖

在 **Web** 标签里：

1. 在 **Virtualenv** 一栏点击 **Enter path to a virtualenv**，输入例如：`/home/你的用户名/.virtualenvs/tkpy`（或你打算用的路径）。
2. 若该虚拟环境尚未创建，先到 **Consoles** 开一个 Bash，执行：
   ```bash
   mkvirtualenv --python=/usr/bin/python3.10 tkpy
   pip install flask pandas openpyxl numpy
   ```
   若你已上传了 `requirements.txt`，可在 TkPy 目录下执行：
   ```bash
   cd ~/TkPy   # 或 cd ~/你的仓库/TkPy
   pip install -r requirements.txt
   ```
3. 确认 **Web** 里 Virtualenv 指向的路径与上面一致。

### 4. 配置 WSGI

- 若已上传本仓库里的 **`wsgi.py`**：在 **Web** 的 **Code** 里，将 **WSGI configuration file** 设为 TkPy 目录下的 `wsgi.py`，例如：  
  `/home/Douglas488/TkPy/wsgi.py`  
  （`Project` 的 **Working directory** 也设为同一目录，如 `/home/Douglas488/TkPy`。）
- 若未使用 `wsgi.py`，可手动编辑默认的 WSGI 文件，在顶部加入路径后再 `from app import app as application`：
  ```python
  import sys
  sys.path.insert(0, '/home/你的用户名/TkPy')   # 改成你的 TkPy 实际路径
  from app import app as application
  ```

保存后回到 **Web** 页面，点击 **Reload** 重载应用。

### 5. 验证接口

- 浏览器访问：`https://你的用户名.pythonanywhere.com/`  
  应看到一段 JSON，说明服务正常。
- 用 Postman 或 curl 测试上传：
  ```bash
  curl -X POST -F "file=@你的测试.xlsx" https://你的用户名.pythonanywhere.com/api/generate -o out.xlsx
  ```
  若返回 Excel 且无报错，说明部署成功。

---

## 二、在 index9 中调用该服务

1. 打开项目里的 **index9.html**。
2. 在页面中的「API 地址」输入框里填写你在 PythonAnywhere 的站点根地址，例如：  
   `https://你的用户名.pythonanywhere.com`  
   （不要加末尾斜杠，不要带 `/api/generate`。）
3. 选择本地 Excel 文件，点击「生成总结表」。
4. 页面会向 `你的API地址/api/generate` 发送 POST 请求并下载返回的「Tk月总结表.xlsx」。

若部署在子路径或使用自定义域名，只需保证在 index9 里填写的「API 地址」与浏览器能访问的根地址一致即可；index9 会自动在其后拼接 `/api/generate`。

---

## 三、报错「没有名为 flask 的模块」与「IndentationError」的修复步骤（Douglas488 示例）

按下面顺序做，不要跳过。

### 步骤 1：创建虚拟环境并安装依赖

1. 登录 PythonAnywhere，打开 **Consoles**，点击 **Bash** 打开一个终端。
2. 在终端里依次执行（用户名换成你的，例如 Douglas488）：

```bash
cd ~
mkvirtualenv --python=/usr/bin/python3.10 tkpy
pip install flask pandas openpyxl numpy
```

若提示 `mkvirtualenv` 找不到，先执行：

```bash
pip install --user virtualenvwrapper
```

然后关闭该 Bash，重新开一个 Bash，再执行上面的 `mkvirtualenv` 和 `pip install`。

安装完成后可执行 `pip list` 确认能看到 `flask`、`pandas`、`openpyxl`。

### 步骤 2：在 Web 里绑定虚拟环境

1. 打开 **Web** 标签。
2. 找到 **Virtualenv** 这一项，点 **Enter path to a virtualenv**。
3. 输入（把 Douglas488 换成你的用户名）：

```
/home/Douglas488/.virtualenvs/tkpy
```

4. 点右侧绿色勾保存。

### 步骤 3：修正 WSGI 文件内容（解决 IndentationError）

WSGI 配置文件是 **Web 里点进去编辑** 的那个（例如 `/var/www/douglas488_pythonanywhere_com_wsgi.py`），**不要**从 README 里复制带「\`\`\`python」的整段，否则会报「意外的缩进」。

1. 在 **Web** 里点 **WSGI configuration file** 的链接，打开编辑器。
2. **删除里面全部内容**，只保留下面这几行。**只复制代码本身，不要复制「\`\`\`」或「\`\`\`python」**；每行从行首开始，第一行是 `import sys` 且前面没有空格：

```python
import sys
import os

this_dir = '/home/Douglas488/TkPy'
if this_dir not in sys.path:
    sys.path.insert(0, this_dir)
os.chdir(this_dir)

from app import app as application
```

3. 把上面代码里的 `Douglas488` 和 `TkPy` 改成你实际的项目路径（即放 `app.py`、`generate_summary.py` 的目录）。  
   也可以直接打开项目里的 **`wsgi_for_pythonanywhere.txt`**，复制全部内容到 WSGI 编辑器，再改路径。
4. 保存后回到 **Web**，点 **Reload** 重载应用。

若你的 TkPy 在别的位置（例如克隆在 `EtiquetaFull-main/TkPy`），则 `this_dir` 应改为类似：`/home/Douglas488/EtiquetaFull-main/TkPy`。

### 步骤 4：确认

- 再访问：https://Douglas488.pythonanywhere.com/
- 应看到 JSON，且其中有 `"generate_summary": "ok"`。若为 `"failed"`，把 `error` 里的内容发出来再排查。

---

## 四、打开站点报错时怎么查（其他情况）

1. **先看浏览器里打开的首页**  
   访问：`https://Douglas488.pythonanywhere.com/`  
   - 若返回 JSON 且里有 `"generate_summary": "failed"` 和 `"error": "..."`，说明 Flask 已起来，但依赖或 `generate_summary` 导入失败，根据 `error` 里的内容修（例如缺模块就 `pip install pandas openpyxl`）。
   - 若直接 500 或打不开，看下面第 2 步。

2. **看 PythonAnywhere 的 Error log**  
   登录 PythonAnywhere → **Web** → 页面下方 **Log files** 里点 **Error log**，看最后几行的报错（例如 `ModuleNotFoundError: No module named 'xxx'` 或 `ImportError`），按提示安装依赖或修正路径。

3. **确认 WSGI 指向和目录**  
   - **WSGI configuration file** 必须指到包含 `app.py` 和 `generate_summary.py` 的目录下的 `wsgi.py`（或你手改的 WSGI 文件），且该文件里 `sys.path.insert(0, 该目录)`。
   - **Project** 的 **Working directory** 设为同一 TkPy 目录。

4. **确认虚拟环境里已装依赖**  
   在 **Consoles** 里激活同一虚拟环境后执行：  
   `pip install -r requirements.txt`  
   或：  
   `pip install flask pandas openpyxl numpy`
