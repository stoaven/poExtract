#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 字段解析程序

策略:
1. 优先尝试提取标准 AcroForm 字段 (交互式 PDF 表单)
2. 对于版式生成类 PDF (如 Oracle BI Publisher 渲染的),
   使用 pdfplumber.extract_words 拿到每个单词的 (x0, top) 坐标,
   按"标签: 值"在列上对齐的方式重新组合字段。

用法:
    python parse_pdf_fields.py <pdf_path> [-o output.json]

依赖:
    pip install pypdf pdfplumber
"""

# 支持 Python 3.9 的 PEP 604 类型注解
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Optional

from pypdf import PdfReader

try:
    import pdfplumber  # type: ignore

    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False


# ============================================================
# AcroForm 字段
# ============================================================

def _field_type_name(ft: Optional[str]) -> str:
    return {
        "/Tx": "text", "/Btn": "button", "/Ch": "choice", "/Sig": "signature"
    }.get(ft or "", "unknown")


def extract_acroform_fields(reader: PdfReader) -> dict[str, Any]:
    fields: dict[str, Any] = {}
    raw = reader.get_fields() or {}
    for name, info in raw.items():
        fields[name] = {
            "value": info.get("/V"),
            "default_value": info.get("/DV"),
            "type": _field_type_name(info.get("/FT")),
            "tooltip": info.get("/TU"),
            "options": info.get("/Opt"),
        }
    return fields


# ============================================================
# 标签定义
# ============================================================

INLINE_LABELS: list[tuple[list[str], str]] = [
    (["ATTN", "(BUYER):"], "attn_buyer"),
    (["BP_GSTAR", "DELIVERY", "NO:"], "bp_gstar_delivery_no"),
    (["CARGO", "READY", "DATE:"], "cargo_ready_date"),
    (["EMAIL:"], "email"),
    (["TEL:"], "tel"),
    (["FAX:"], "fax"),
    (["CONTACT:"], "contact"),
    (["Scan", "Flag:"], "scan_flag"),
    (["Cargo", "cutoff"], "cargo_cutoff"),
    (["Verified", "by:"], "verified_by"),
    (["DELIVERY", "NO:"], "delivery_no"),
    # 用于 rpt_*.pdf 等: 行内左侧标签 + 同行的右侧长描述
    (["ARTICLEID"], "description_of_packages"),
    (["ORDER", "NR.:"], "order_nr"),
]

# 对个别行内标签, 使用更大的 x-gap 阈值 (默认值 25 px).
# 用于 ARTICLEID / ORDER NR. 这类左侧标签-右侧值的"横向堆叠"场景.
WIDE_GAP_INLINE_KEYS: set[str] = {"description_of_packages", "order_nr"}

STACKED_LABELS: list[tuple[list[str], str, int]] = [
    (["SELLER", "PAYER", "ON", "APLL", "LOCAL", "CHARGE", "INVOICE"], "seller_payer", 5),
    (["BUYER", "NOTIFY", "PARTY"], "buyer_notify_party", 5),
    (["ALSO", "NOTIFY", "PARTY"], "also_notify_party", 5),
    (["VESSEL", "VOYAGE", "LOAD", "TYPE", "/", "CARGO", "TYPE"], "vessel_voyage_load_type", 1),
    (["PORT", "OF", "LOADING"], "port_of_loading", 1),
    (["PLACE", "OF", "RECEIPT"], "place_of_receipt", 1),
    (["PORT", "OF", "DISCHARGE"], "port_of_discharge", 1),
    (["PLACE", "OF", "DELIVERY"], "place_of_delivery", 1),
    (["COMMODITY"], "commodity", 1),
    (["COUNTRY", "OF", "ORIGIN"], "country_of_origin", 1),
    (["CPO#"], "cpo_no", 1),
    (["INCOTERM"], "incoterm", 1),
    (["MODE"], "mode", 1),
    (["PALLETS", "STACKABLE"], "pallets_stackable", 1),
    (["HAZARDOUS", "MATERIALS"], "hazardous_materials", 1),
    (["TAX", "ID:"], "tax_id", 1),
    (["REFERENCE", "NO.:"], "reference_no", 1),
    (["LETTER", "OF", "CREDIT", "NO."], "letter_of_credit_no", 1),
    (["PO", "NO."], "po_no", 1),
    (["SolidWoodPackingMaterial:"], "solid_wood_packing_material", 1),
    (["FDA", "Regulated", "Product:"], "fda_regulated_product", 1),
    (["SHIPMENT", "#"], "shipment_no", 1),
    (["SPECIAL", "PROGRAM"], "special_program_scan_flag", 1),
    # 用于 624303.pdf: 跨 2 行的完整标签, 配合 _find_label_span_with_skip 处理同行其他列的词
    (["DESCRIPTION", "OF", "PACKAGES", "AND", "GOODS", "PARTICULARS",
      "FURNISHED", "BY", "SHIPPER"], "description_of_packages", 1),
]

# 部分标签需要 y 容差 (允许跨行匹配) 或更大的 max_search_distance (值距离较远)
LABEL_EXT_OVERRIDES: dict[str, dict[str, Any]] = {
    "description_of_packages": {
        "y_tol": 12.0,            # 允许标签跨 2 行匹配
        "allow_skip": True,       # 跳过同行其他列的词 (GROSS/MEASUREMENT 等)
        "max_search_distance": 25.0,  # 标签最后一行到值行可能有较大间隔
    },
}

INDICATOR_LABELS: list[tuple[list[str], str]] = [
    (["FIRST", "MILE", "INDICATOR"], "first_mile_indicator"),
    (["LAST", "MILE", "INDICATOR"], "last_mile_indicator"),
    (["EXPORT", "CUSTOM", "INDICATOR"], "export_custom_indicator"),
    (["IMPORT", "CUSTOM", "INDICATOR"], "import_custom_indicator"),
]


# ============================================================
# 词级坐标提取
# ============================================================

def _extract_words(pdf_path: Path) -> list[dict[str, Any]]:
    all_words: list[dict[str, Any]] = []
    if not _HAS_PDFPLUMBER:
        return all_words
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            for w in page.extract_words(use_text_flow=False):
                all_words.append(
                    {
                        "text": w["text"],
                        "x0": round(w["x0"], 1),
                        "x1": round(w["x1"], 1),
                        "top": round(w["top"], 1),
                        "bottom": round(w["bottom"], 1),
                        "page": page_idx,
                    }
                )
    return all_words


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().rstrip(":.")).upper()


# ============================================================
# 标签匹配: 允许同一行跨多列
# ============================================================

def _find_label_span(
    words: list[dict[str, Any]], label_tokens: list[str], y_tol: float = 5.0
) -> Optional[dict[str, Any]]:
    """匹配标签: 默认要求连续 token 在同一行 (y_tol=5).

    对特殊标签, 通过 y_tol 参数允许跨行匹配 (例如 624303.pdf 中的
    "DESCRIPTION OF PACKAGES AND GOODS PARTICULARS / FURNISHED BY SHIPPER").
    """
    label_tokens_norm = [_norm(t) for t in label_tokens]
    n = len(label_tokens_norm)
    for i in range(len(words) - n + 1):
        if words[i]["page"] != words[i + n - 1]["page"]:
            continue
        matched = True
        for k in range(n):
            if _norm(words[i + k]["text"]) != label_tokens_norm[k]:
                matched = False
                break
        if not matched:
            continue
        row_ok = True
        last_x = -float("inf")
        label_max_top = words[i]["top"]
        for k in range(n):
            if abs(words[i + k]["top"] - words[i]["top"]) > y_tol:
                row_ok = False
                break
            if words[i + k]["x0"] < last_x:
                row_ok = False
                break
            last_x = words[i + k]["x0"]
            label_max_top = max(label_max_top, words[i + k]["top"])
        if not row_ok:
            continue
        return {
            "start_idx": i,
            "end_idx": i + n - 1,
            "x_start": words[i]["x0"],
            "x_end": words[i + n - 1]["x1"],
            "top": words[i]["top"],
            "max_top": label_max_top,
            "page": words[i]["page"],
        }
    return None


def _find_label_span_with_skip(
    words: list[dict[str, Any]], label_tokens: list[str], y_tol: float = 12.0
) -> Optional[dict[str, Any]]:
    """允许跨行匹配 + 同列跳过非标签词 (适用于模板, 例如 description 标签跨行)."""
    label_tokens_norm = [_norm(t) for t in label_tokens]
    n = len(label_tokens_norm)
    search_window = 30  # 每个 token 最多向后看多少个 word

    for i in range(len(words)):
        if not label_tokens_norm:
            break
        # 尝试从这个位置开始匹配第一个 token
        if _norm(words[i]["text"]) != label_tokens_norm[0]:
            continue
        if words[i]["page"] != words[i]["page"]:
            continue
        consumed: list[int] = []
        cursor = i
        ok = True
        last_row_top = words[i]["top"]
        label_max_top = last_row_top
        for k in range(n):
            # 在 cursor 之后 search_window 范围内找 label_tokens_norm[k]
            found = False
            for j in range(cursor, min(cursor + search_window, len(words))):
                w = words[j]
                if _norm(w["text"]) != label_tokens_norm[k]:
                    continue
                # 必须在同一页
                if w["page"] != words[i]["page"]:
                    break
                # 必须在同一行同一标签的水平范围内 (用第一个 token 的 x_start 容差)
                if abs(w["x0"] - words[i]["x0"]) > 150:
                    continue
                # 必须在同一行或下一行 (y_tol 范围内)
                if abs(w["top"] - last_row_top) <= 4:
                    pass  # same row
                elif abs(w["top"] - last_row_top) <= y_tol:
                    # next row, update last_row_top
                    last_row_top = w["top"]
                else:
                    continue
                if w["x0"] < words[i]["x0"] - 5 and len(consumed) > 0:
                    continue
                consumed.append(j)
                cursor = j + 1
                label_max_top = max(label_max_top, w["top"])
                found = True
                break
            if not found:
                ok = False
                break
        if not ok or len(consumed) != n:
            continue
        start_idx = consumed[0]
        end_idx = consumed[-1]
        return {
            "start_idx": start_idx,
            "end_idx": end_idx,
            "x_start": words[start_idx]["x0"],
            "x_end": words[end_idx]["x1"],
            "top": words[start_idx]["top"],
            "max_top": label_max_top,
            "page": words[start_idx]["page"],
        }
    return None


# ============================================================
# 值提取
# ============================================================

def _gather_value_below(
    words: list[dict[str, Any]],
    span: dict[str, Any],
    end_idx: int,
    max_rows: int = 1,
    y_tol: float = 4.0,
    x_tol: float = 8.0,
    row_gap_threshold: float = 25.0,
    max_search_distance: float = 10.0,
    multirow_max_distance: float = 80.0,
) -> str:
    """从标签段下方收集同一 x 列的值, max_rows=1 时只收集 1 行.

    `span["max_top"]` 是标签最后一行 top, 用于跳过整个标签 (含跨行延伸).
    `span["top"]` 是标签第一行 top, 仍用于相对距离.
    """
    label_x_start = span["x_start"]
    label_max_top = span.get("max_top", span["top"])

    # 跳过标签的所有行 (含跨行延伸)
    target_top: Optional[float] = None
    target_idx: int = end_idx
    search_top_max = label_max_top + 4 + max_search_distance
    for j in range(end_idx + 1, len(words)):
        w = words[j]
        if w["page"] != span["page"]:
            break
        if w["top"] <= label_max_top:
            continue
        if w["top"] > search_top_max:
            break
        if abs(w["x0"] - label_x_start) <= x_tol:
            target_top = w["top"]
            target_idx = j
            break
    if target_top is None:
        return ""

    # 在 target_top 行上, 从 label_x_start - x_tol 开始向右收所有词 (直到出现大跳跃)
    row_words: list[dict[str, Any]] = []
    for j in range(target_idx, len(words)):
        w = words[j]
        if w["page"] != span["page"]:
            break
        if abs(w["top"] - target_top) > y_tol:
            break
        if w["x0"] < label_x_start - x_tol:
            continue
        row_words.append(w)

    row_words.sort(key=lambda w: w["x0"])
    trimmed_row: list[dict[str, Any]] = []
    last_x1 = -float("inf")
    for w in row_words:
        if trimmed_row and (w["x0"] - last_x1) > row_gap_threshold:
            break
        trimmed_row.append(w)
        last_x1 = w["x1"]
    flat_text = " ".join(w["text"] for w in trimmed_row).strip()

    if max_rows <= 1:
        return flat_text

    # 多行模式: 用更宽的距离限制收集多行值 (地址等).
    multi_top_max = label_max_top + 4 + multirow_max_distance
    lines_out: list[str] = [flat_text] if flat_text else []
    cur_top: Optional[float] = None
    cur_words: list[dict[str, Any]] = []
    for j in range(target_idx, len(words)):
        w = words[j]
        if w["page"] != span["page"]:
            break
        if w["top"] <= label_max_top:
            continue
        if w["top"] > multi_top_max:
            break
        if abs(w["top"] - target_top) <= y_tol:
            continue  # 跳过 target_top 行 (单行已收集)
        if abs(w["x0"] - label_x_start) > x_tol:
            continue
        if cur_words is not None and cur_top is not None and abs(w["top"] - cur_top) > y_tol:
            if len(lines_out) >= max_rows:
                break
            sorted_line = sorted(cur_words, key=lambda x: x["x0"])
            line_text = " ".join(x["text"] for x in sorted_line).strip()
            if line_text and line_text not in lines_out:
                lines_out.append(line_text)
            cur_words = []
            cur_top = w["top"]
        if not cur_words:
            cur_top = w["top"]
        cur_words.append(w)
    if cur_words and len(lines_out) < max_rows:
        sorted_line = sorted(cur_words, key=lambda x: x["x0"])
        line_text = " ".join(x["text"] for x in sorted_line).strip()
        if line_text and line_text not in lines_out:
            lines_out.append(line_text)

    return " ".join(lines_out).strip()


def _gather_inline_value_after_label(
    words: list[dict[str, Any]],
    span: dict[str, Any],
    end_idx: int,
    x_gap_threshold: float = 25.0,
) -> str:
    """对 LABEL: VALUE 同行的标签, 提取同行紧随其后的值."""
    label_x_end = span["x_end"]
    label_top = span["top"]
    parts: list[tuple[float, str]] = []
    last_x_end = label_x_end
    for j in range(end_idx + 1, len(words)):
        w = words[j]
        if w["page"] != span["page"]:
            break
        if abs(w["top"] - label_top) > 4:
            break
        if w["x0"] < label_x_end - 2:
            continue
        raw = w["text"].strip()
        # 退出条件: 当前词以 ":" 结尾 (视为新标签)
        if parts and raw.endswith(":") and not raw[0].isdigit():
            break
        # 退出条件: 当前词与上一个词的 x 跳跃过大
        if parts and (w["x0"] - last_x_end) > x_gap_threshold:
            break
        parts.append((w["x0"], w["text"]))
        last_x_end = w["x1"]
    parts.sort(key=lambda p: p[0])
    return " ".join(t for _, t in parts).strip()


# ============================================================
# 4 个 INDICATOR 组 (共享 N 值)
# ============================================================

def _extract_indicator_group(
    words: list[dict[str, Any]], existing: dict[str, Any]
) -> None:
    spans: list[tuple[list[str], str, dict[str, Any]]] = []
    for tokens, key in INDICATOR_LABELS:
        if key in existing:
            continue
        span = _find_label_span(words, tokens)
        if span:
            spans.append((tokens, key, span))
    if len(spans) < 4:
        return
    last_span = max(spans, key=lambda s: s[2]["top"])
    last_bottom = max(
        words[i]["bottom"]
        for i in range(last_span[2]["start_idx"], last_span[2]["end_idx"] + 1)
    )
    candidates: list[dict[str, Any]] = []
    for j in range(last_span[2]["end_idx"] + 1, len(words)):
        w = words[j]
        if w["page"] != last_span[2]["page"]:
            break
        if w["top"] < last_bottom - 1:
            continue
        if abs(w["top"] - last_span[2]["top"]) > 25:
            break
        t = w["text"].strip().upper()
        if t in {"N", "Y"}:
            candidates.append(w)
        elif candidates and len(candidates) >= 4:
            break
    if len(candidates) >= 4:
        candidates.sort(key=lambda c: c["x0"])
        for (_, key, _), cand in zip(spans, candidates[:4]):
            existing[key] = cand["text"]


# ============================================================
# PO 明细行 (基于 pdfplumber 的词级坐标)
# ============================================================

_PO_HEADER_RE = re.compile(r"PO#\s*[: ]\s*(\d+)", re.IGNORECASE)
_BOOK_TOTAL_RE = re.compile(
    r"Booking\s+Total\s+(.+)$", re.IGNORECASE | re.MULTILINE
)
_NUM_UNIT_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(PCS|EA|CTN|KG|CUMTR|CBM)", re.IGNORECASE
)

_UNIT_SET = {"EA", "CTN", "KG", "CUMTR", "CBM"}

# 表头关键字 / 维度注解
_HEADER_TOKENS = {"STYLE", "SUB", "QTY", "PKGS", "WEIGHT", "VOLUME",
                  "QTY/PKGS", "QTY/PKG", "LENGTH", "WIDTH", "HEIGHT"}

# PO 主数据行模板: STYLE_CODE QTY PKG WEIGHT VOLUME QTY/PKG
# 兼容字母开头 (D3003) 与纯数字 (222200)
_PO_DATA_PATTERN = re.compile(
    r"^([A-Z0-9][A-Z0-9]{2,})\s+(\d+)\s+(\d+)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+)$"
)

# "0- EA CTN KG CUMTR" / "2451- PCS CTN KG CUMTR" 模式: sub_style 前缀 + 4 个单位
_PO_UNITS_PATTERN = re.compile(
    r"^(\S+)\s+(EA|PCS)\s+(CTN)\s+(KG)\s+(CUMTR|CBM)$"
)


def _split_block_by_words(block_text: str) -> list[str]:
    """把每个 PO 块切成行 (PDF 文本按行)。"""
    return [l for l in block_text.split("\n") if l.strip()]


def _extract_sub_total(lines: list[str], start_idx: int) -> Optional[str]:
    """从 "Sub-Total per" 行起, 合并 'PO:' 后续单位."""
    parts: list[str] = []
    for j in range(start_idx, len(lines)):
        line = lines[j].strip()
        if line.lower().startswith("sub-total per"):
            parts.append(line[len("Sub-Total per"):].strip() if line.lower().startswith("sub-total per") else line)
            # 检查下一行是否是 "PO: UNIT"
            if j + 1 < len(lines):
                nxt = lines[j + 1].strip()
                m = re.match(r"^PO\s*:\s*(\S+)$", nxt)
                if m:
                    parts.append(m.group(1))
            return " ".join(parts).strip()
    return None


def _parse_po_block(po_no: str, block: str) -> dict[str, Any]:
    """解析一个 PO 块: 内含 1..N 个 entries (每个 entry 对应一个子样式 / 尺码行)."""
    item: dict[str, Any] = {"po_no": po_no, "entries": []}
    lines = _split_block_by_words(block)

    cur_entry: Optional[dict[str, Any]] = None
    common_style: Optional[str] = None

    def _finalize_entry():
        nonlocal cur_entry
        if cur_entry is not None:
            item["entries"].append(cur_entry)
            cur_entry = None

    for line in lines:
        s = line.strip()
        if not s:
            continue
        # 跳过的内容
        if re.match(r"^PO#\s*[: ]", s, re.IGNORECASE):
            continue
        if re.match(r"^STYLE\s+SUB\s+STYLE", s, re.IGNORECASE):
            continue
        if re.match(r"^(Length|Width|Height)\s*:", s, re.IGNORECASE):
            # 一个条目的边界
            _finalize_entry()
            continue
        if s.lower().startswith("sub-total") or s.startswith("PO:"):
            continue
        if s.lower().startswith("booking total"):
            continue

        # 主数据行
        m = _PO_DATA_PATTERN.match(s)
        if m:
            _finalize_entry()
            common_style = m.group(1)
            cur_entry = {
                "style": m.group(1),
                "qty": {"value": m.group(2), "unit": "EA"},
                "pkg_count": {"value": m.group(3), "unit": "CTN"},
                "weight": {"value": m.group(4), "unit": "KG"},
                "volume": {"value": m.group(5), "unit": "CUMTR"},
                "qty_per_pkg": m.group(6),
            }
            continue

        # 单位行 (含 sub_style 前缀)
        mu = _PO_UNITS_PATTERN.match(s)
        if mu:
            if cur_entry:
                cur_entry["sub_style"] = mu.group(1)
                cur_entry["qty"]["unit"] = mu.group(2)
                cur_entry["pkg_count"]["unit"] = mu.group(3)
                cur_entry["weight"]["unit"] = mu.group(4)
                cur_entry["volume"]["unit"] = mu.group(5)
            continue

        # sub_style 延续单行. 允许:
        #   - 字母 + 数字 + 短横 + 斜杠 (例如 "E627-" "J384" "021-L" "100-L")
        #   - 跨行的 "134/14"+"0" 组成 "134/140" 这种"短行延续"
        if cur_entry and re.match(r"^[A-Z0-9][A-Z0-9\-/]*$", s):
            cur_entry["sub_style"] = cur_entry.get("sub_style", "") + s
            continue

    # 处理最后一个 entry
    _finalize_entry()

    # PO 级别汇总
    if common_style and all(e["style"] == common_style for e in item["entries"]):
        item["style"] = common_style

    sub_total_text = _extract_sub_total(lines, 0)
    if sub_total_text:
        item["sub_total"] = _parse_value_pairs(sub_total_text)

    return item


def _parse_value_pairs(text: str) -> dict[str, Any]:
    result: dict[str, Any] = {"raw": text}
    pairs = _NUM_UNIT_RE.findall(text)
    if len(pairs) >= 4:
        result["qty"] = {"value": pairs[0][0], "unit": pairs[0][1].upper()}
        result["pkg_count"] = {
            "value": pairs[1][0], "unit": pairs[1][1].upper(),
        }
        result["weight"] = {
            "value": pairs[2][0], "unit": pairs[2][1].upper(),
        }
        result["volume"] = {
            "value": pairs[3][0], "unit": pairs[3][1].upper(),
        }
    return result


def extract_po_details(all_text: str) -> dict[str, Any]:
    po_blocks: list[dict[str, Any]] = []
    pos = 0
    while True:
        m = _PO_HEADER_RE.search(all_text, pos)
        if not m:
            break
        po_no = m.group(1)
        candidate_ends = [len(all_text)]
        nxt = _PO_HEADER_RE.search(all_text, m.end())
        if nxt:
            candidate_ends.append(nxt.start())
        bt = _BOOK_TOTAL_RE.search(all_text, m.end())
        if bt:
            candidate_ends.append(bt.start())
        end_idx = min(candidate_ends)
        block = all_text[m.start(): end_idx]
        po_blocks.append(_parse_po_block(po_no, block))
        pos = end_idx

    booking_total = None
    bt = _BOOK_TOTAL_RE.search(all_text)
    if bt:
        booking_total = _parse_value_pairs(bt.group(1).strip())
    return {"po_details": po_blocks, "booking_total": booking_total}


# ============================================================
# 主入口
# ============================================================

def parse_pdf(pdf_path: str | Path) -> dict[str, Any]:
    pdf_path = Path(pdf_path)
    reader = PdfReader(str(pdf_path))

    result: dict[str, Any] = {
        "file": pdf_path.name,
        "page_count": len(reader.pages),
        "metadata": {k: str(v) for k, v in (reader.metadata or {}).items()},
        "form_fields": extract_acroform_fields(reader),
        "labelled_fields": {},
        "po_details": [],
        "booking_total": None,
    }

    if result["form_fields"]:
        return result

    if not _HAS_PDFPLUMBER:
        raise RuntimeError(
            "未找到 pdfplumber, 请先 `pip install pdfplumber`, 或改用 AcroForm PDF."
        )

    words = _extract_words(pdf_path)
    out: dict[str, Any] = {}
    used_word_indices: set[int] = set()

    def _mark_used(span):
        for i in range(span["start_idx"], span["end_idx"] + 1):
            used_word_indices.add(i)

    for tokens, key in INLINE_LABELS:
        span = _find_label_span(words, tokens)
        if not span:
            continue
        if any(i in used_word_indices for i in range(span["start_idx"], span["end_idx"] + 1)):
            continue
        x_gap = 200.0 if key in WIDE_GAP_INLINE_KEYS else 25.0
        out[key] = _gather_inline_value_after_label(
            words, span, span["end_idx"], x_gap_threshold=x_gap,
        )
        _mark_used(span)

    for tokens, key, max_rows in STACKED_LABELS:
        if key in out:
            continue
        overrides = LABEL_EXT_OVERRIDES.get(key, {})
        # 决定使用哪个匹配器
        if overrides.get("allow_skip"):
            span = _find_label_span_with_skip(
                words, tokens, y_tol=overrides.get("y_tol", 12.0),
            )
        else:
            span = _find_label_span(words, tokens, y_tol=overrides.get("y_tol", 5.0))
        if not span:
            continue
        if any(i in used_word_indices for i in range(span["start_idx"], span["end_idx"] + 1)):
            continue
        out[key] = _gather_value_below(
            words,
            span,
            span["end_idx"],
            max_rows=max_rows,
            max_search_distance=overrides.get("max_search_distance", 10.0),
        )
        _mark_used(span)

    _extract_indicator_group(words, out)

    result["labelled_fields"] = out

    # PO 明细: 基于普通文本顺序扫描
    plain_text_parts: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            plain_text_parts.append(page.extract_text() or "")
    plain_text = "\n".join(plain_text_parts)
    po_info = extract_po_details(plain_text)
    result["po_details"] = po_info["po_details"]
    result["booking_total"] = po_info["booking_total"]

    return result


# ============================================================
# CLI
# ============================================================

def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python parse_pdf_fields.py <pdf_path> [-o output.json]")
        print("  默认输出: <pdf_path>.json (同主文件名, 后缀改为 .json)")
        return 1

    pdf_path = sys.argv[1]
    output_path = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            output_path = sys.argv[idx + 1]

    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        print(f"找不到文件: {pdf_path}")
        return 1

    # 默认输出: 同主文件名, 后缀改为 .json
    if output_path is None:
        output_path = pdf_file.with_suffix(".json")

    parsed = parse_pdf(pdf_file)
    json_text = json.dumps(parsed, ensure_ascii=False, indent=2)
    Path(output_path).write_text(json_text, encoding="utf-8")
    print(f"已写入: {output_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
