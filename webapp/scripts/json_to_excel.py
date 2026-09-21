#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 parse_pdf_fields.py 生成的 JSON 解析为 Excel 表格。

输入 JSON 结构 (parse_pdf_fields.py 输出):
    {
      "labelled_fields": {
        "commodity": "GARMENT AND TEXTILES",
        ...
      },
      "po_details": [
        {
          "po_no": "389500",
          "style": "222200",
          "entries": [
            {
              "style": "222200",
              "sub_style": "2451-021-L",
              "qty": {"value": "64", "unit": "PCS"},
              "pkg_count": {"value": "4", "unit": "CTN"},
              "weight": {"value": "24.000", "unit": "KG"},
              "volume": {"value": "0.288", "unit": "CUMTR"},
              "qty_per_pkg": "16"
            },
            ...
          ],
          "sub_total": {...}
        }
      ]
    }

Excel 列定义 (每个 entry 一行):
    po number        : po_details[i].po_no
    item number      : po_details[i].entries[j].style + po_details[i].entries[j].sub_style
    goods description: labelled_fields.commodity
    ctns             : po_details[i].entries[j].pkg_count.value
    pieces           : po_details[i].entries[j].qty.value
    gross weight     : po_details[i].entries[j].weight.value
    cbm              : po_details[i].entries[j].volume.value

用法:
    python json_to_excel.py <json_path> [-o output.xlsx]

依赖:
    pip install openpyxl
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# 列定义: (header, 值提取函数)
# 值提取函数接收 (po, entry, labelled_fields) 并返回值
def _col(po, entry, lf, key_default=""):
    return key_default


def _get_entry_value(entry: dict[str, Any], key: str) -> str:
    """从 entries[j] 中抽取 .value (容错, 空返回空字符串)."""
    v = entry.get(key)
    if isinstance(v, dict):
        return str(v.get("value", "") or "")
    return str(v or "")


COLUMNS: list[tuple[str, Any]] = [
    ("po number", lambda po, e, lf: po.get("po_no", "")),
    ("item number", lambda po, e, lf: f"{e.get('style', '')}{e.get('sub_style', '')}"),
    # goods description: 优先 description_of_packages; 但若其为 "NO" (空字段) 则回退到 commodity (类别)
    ("goods description", lambda po, e, lf: (
        lf["description_of_packages"]
        if lf.get("description_of_packages") and lf["description_of_packages"] != "NO"
        else lf.get("commodity", "")
    )),
    ("ctns", lambda po, e, lf: _get_entry_value(e, "pkg_count")),
    ("pieces", lambda po, e, lf: _get_entry_value(e, "qty")),
    ("gross weight", lambda po, e, lf: _get_entry_value(e, "weight")),
    ("cbm", lambda po, e, lf: _get_entry_value(e, "volume")),
]


def build_rows(data: dict[str, Any]) -> list[list[str]]:
    """从 JSON 数据构建 Excel 行 (header + data rows)."""
    labelled = data.get("labelled_fields", {}) or {}
    pos = data.get("po_details", []) or []

    rows: list[list[str]] = []
    headers = [c[0] for c in COLUMNS]
    rows.append(headers)
    for po in pos:
        entries = po.get("entries", []) or []
        if not entries:
            # PO 块无 entry 时, 仍输出一行 (含 PO 总览)
            row = [fn(po, {}, labelled) for (_, fn) in COLUMNS]
            rows.append(row)
            continue
        for entry in entries:
            row = [fn(po, entry, labelled) for (_, fn) in COLUMNS]
            rows.append(row)
    return rows


def write_excel(rows: list[list[str]], output_path: Path) -> None:
    """写入 .xlsx 文件.

    样式与 Booking Form - Air.xlsx 模板保持一致:
        - 字体 Arial 12 (与模板一致)
        - 表头加粗, 但无蓝色填充背景 (匹配模板的"无填充"风格)
        - 默认对齐 (不强制居中)
        - 不冻结表头 (与模板一致)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "PO Details"

    # 字体: Arial 12, 表头加粗
    data_font = Font(name="Arial", size=12)
    header_font = Font(name="Arial", size=12, bold=True)

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.font = header_font if r_idx == 1 else data_font
            # 默认对齐, 不强制居中

    # 列宽: 自动 + 较保守的上限 (与模板的小列宽一致)
    for c_idx, (header, _) in enumerate(COLUMNS, start=1):
        col_letter = get_column_letter(c_idx)
        max_len = max(
            [len(str(header))] + [len(str(r[c_idx - 1])) for r in rows[1:] if r[c_idx - 1]]
        )
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 32)

    # 不冻结表头 (与 Booking Form - Air.xlsx 模板保持一致)

    wb.save(str(output_path))


def parse_json_to_excel(json_path: Path, output_path: Path) -> int:
    """主流程."""
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = build_rows(data)
    write_excel(rows, output_path)
    print(f"已写入 {len(rows) - 1} 行数据到 {output_path}")
    return 0


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python json_to_excel.py <json_path> [-o output.xlsx]")
        print("  默认输出: <json_path>.xlsx (同主文件名, 后缀改为 .xlsx)")
        return 1

    json_path = Path(sys.argv[1])
    if not json_path.exists():
        print(f"找不到文件: {json_path}")
        return 1

    output_path = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            output_path = Path(sys.argv[idx + 1])

    if output_path is None:
        output_path = json_path.with_suffix(".xlsx")

    return parse_json_to_excel(json_path, output_path)


if __name__ == "__main__":
    sys.exit(main())
