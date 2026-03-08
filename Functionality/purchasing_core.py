# -*- coding: utf-8 -*-
"""采购 Excel 转 JSON（无 GUI），供 Web API 调用。"""
import os
import json
from typing import Dict, Any
from openpyxl import load_workbook


def extract_hyperlinks_from_excel(excel_file_path: str) -> Dict[str, Any]:
    """从 Excel 提取数据与超链接，返回可序列化的字典。"""
    try:
        workbook = load_workbook(excel_file_path, data_only=False)
        result = {
            "sheets": {},
            "metadata": {
                "file_name": os.path.basename(excel_file_path),
                "total_sheets": len(workbook.sheetnames),
            },
        }
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            sheet_data = []
            hyperlinks = {}
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    cell_value = cell.value
                    if cell.hyperlink:
                        hyperlink_info = {
                            "target": cell.hyperlink.target,
                            "tooltip": getattr(cell.hyperlink, "tooltip", "") or "",
                            "display_text": str(cell_value) if cell_value else "",
                        }
                        hyperlinks[cell.coordinate] = hyperlink_info
                        cell_value = f"{cell_value} [HYPERLINK: {cell.hyperlink.target}]" if cell_value else f"[HYPERLINK: {cell.hyperlink.target}]"
                    row_data.append(cell_value)
                if any(c is not None for c in row_data):
                    sheet_data.append(row_data)
            result["sheets"][sheet_name] = {
                "data": sheet_data,
                "hyperlinks": hyperlinks,
                "total_rows": len(sheet_data),
                "total_hyperlinks": len(hyperlinks),
            }
        return result
    except Exception:
        return None
