import pandas as pd
import json
import numpy as np
import re
from tkinter import Tk, filedialog

# 创建 Tkinter 根窗口（不显示）
root = Tk()
root.withdraw()

# 打开文件对话框让用户选择 Excel 文件
file_path = filedialog.askopenfilename(
    title="选择 Excel 文件",
    filetypes=[("Excel files", "*.xlsx")]
)

if file_path:
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 确保 SKU 列的数据类型为字符串
    if 'SKU' in df.columns:
        df['SKU'] = df['SKU'].astype(str)

    # 找出所有产品图 / 变种图列，并按编号排序
    def _sorted_img_cols(prefix: str):
        cols = [c for c in df.columns if c.startswith(prefix)]
        def _key(c):
            m = re.search(r'\d+', c)
            return int(m.group()) if m else 0
        return sorted(cols, key=_key)

    product_img_cols = _sorted_img_cols("产品图")
    variant_img_cols = _sorted_img_cols("变种图")

    # 将 DataFrame 转换为字典列表
    data = df.replace({np.nan: None}).to_dict(orient='records')

    # 处理每一行，把多行 URL 拆分并顺延填充到产品图 / 变种图字段
    for row in data:
        # 如果这一行所有图片字段里都没有换行符，则认为已经按列填好，直接跳过不动
        has_multiline = False
        for col in product_img_cols + variant_img_cols:
            val = row.get(col)
            if isinstance(val, str) and re.search(r'[\r\n]+', val):
                has_multiline = True
                break

        if not has_multiline:
            continue

        all_urls = []

        # 先从所有图片字段中收集 URL，拆分多行
        for col in product_img_cols + variant_img_cols:
            val = row.get(col)
            if not val:
                continue
            if isinstance(val, str):
                parts = re.split(r'[\r\n]+', val)
                parts = [p.strip() for p in parts if p.strip()]
                all_urls.extend(parts)
            else:
                all_urls.append(str(val).strip())

        # 依次回填到产品图和变种图字段中
        idx = 0
        total = len(all_urls)

        for col in product_img_cols:
            row[col] = all_urls[idx] if idx < total else None
            idx += 1

        for col in variant_img_cols:
            row[col] = all_urls[idx] if idx < total else None
            idx += 1

    # 将字典列表保存为 JSON 文件
    with open('products.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print("转换成功，JSON 文件已保存。")
else:
    print("没有选择文件。")

