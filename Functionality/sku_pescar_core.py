# -*- coding: utf-8 -*-
"""库存导 Excel 为 JSON（无 GUI），供 Web API 调用。"""
import re
import pandas as pd
import numpy as np


def _sorted_img_cols(df, prefix: str):
    cols = [c for c in df.columns if c.startswith(prefix)]
    def _key(c):
        m = re.search(r"\d+", c)
        return int(m.group()) if m else 0
    return sorted(cols, key=_key)


def run(excel_path: str):
    """
    读取 Excel，处理产品图/变种图多行 URL，返回字典列表。
    若失败返回 None。
    """
    try:
        df = pd.read_excel(excel_path)
        if "SKU" in df.columns:
            df["SKU"] = df["SKU"].astype(str)
        product_img_cols = _sorted_img_cols(df, "产品图")
        variant_img_cols = _sorted_img_cols(df, "变种图")
        data = df.replace({np.nan: None}).to_dict(orient="records")
        for row in data:
            has_multiline = False
            for col in product_img_cols + variant_img_cols:
                val = row.get(col)
                if isinstance(val, str) and re.search(r"[\r\n]+", val):
                    has_multiline = True
                    break
            if not has_multiline:
                continue
            all_urls = []
            for col in product_img_cols + variant_img_cols:
                val = row.get(col)
                if not val:
                    continue
                if isinstance(val, str):
                    parts = re.split(r"[\r\n]+", val)
                    parts = [p.strip() for p in parts if p.strip()]
                    all_urls.extend(parts)
                else:
                    all_urls.append(str(val).strip())
            idx = 0
            total = len(all_urls)
            for col in product_img_cols:
                row[col] = all_urls[idx] if idx < total else None
                idx += 1
            for col in variant_img_cols:
                row[col] = all_urls[idx] if idx < total else None
                idx += 1
        return data
    except Exception:
        return None
