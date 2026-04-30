#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
現_打鍵前 / 現_打鍵後 の CSV を、テンプレート Excel と同じ形式で展開するツール。

今回の主な仕様:
- 現_打鍵前 / 現_打鍵後 の CSV をファイル名で突合する。
- 片側にしかない CSV は、存在する側のヘッダーを使い、欠損側は0件扱いにする。
- CSV 内に「無効な識別」がある場合は、そのCSVを0件扱いにする。
- 現_打鍵前件数=現_打鍵後件数=0 の表はテーブルシートへ出力せず、テーブル検証 のみに出力する。
- テーブル検証 には CRUD / 出力番号 / 出力Sheet / テーブル名 / 現・新の件数チェック / 差分チェック / 備考 を出力する。
- テーブル検証 の件数・差異・チェック項目は、固定値ではなく Excel 計算式で出力する。
- CRUD が空の テーブル検証 行は既定で非表示にする。
- 既定で 1 sheet につき 5 テーブルを出力し、超過時は テーブル1, テーブル2... を作成する。
- テーブル見出しは「1.テーブル名」のように番号を付ける。見出しの .csv は削除する。
- CSV 列数が 15 列を超える場合はテンプレート末尾列のスタイルを右方向へ拡張する。
- CSV 列数が 15 列未満の場合は、余った列の内容・スタイル・罫線を消す。
- フォントは全セル Yu Gothic、通常行高 18.75、ヘッダー行高 25.5 に統一し、上下中央にする。
- 比較式・条件付き書式を保持する。
"""

import argparse
import csv
import json
import io
import os
import re
import tempfile
import uuid
import warnings
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formatting.rule import CellIsRule
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


TEMPLATE_COLS = 17
DATA_START_COL = 3
FONT_NAME = "Yu Gothic"
BODY_ROW_HEIGHT = 18.75
HEADER_ROW_HEIGHT = 25.5

INVALID_MARKER = "無効な識別"
VALID_DELIMITERS = [",", "\t", ";", "|"]

GRAY_FILL = PatternFill(fill_type="solid", fgColor="D9D9D9")
SUMMARY_HEADER_FILL = PatternFill(fill_type="solid", fgColor="5B9BD5")
SUMMARY_C_FILL = PatternFill(fill_type="solid", fgColor="E2F0D9")
SUMMARY_U_FILL = PatternFill(fill_type="solid", fgColor="DDEBF7")
SUMMARY_D_FILL = PatternFill(fill_type="solid", fgColor="FCE4D6")
SUMMARY_SKIP_FILL = PatternFill(fill_type="solid", fgColor="E7E6E6")

TRUE_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
FALSE_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
FALSE_FONT = Font(name=FONT_NAME, color="9C0006")
WARNING_FILL = PatternFill(fill_type="solid", fgColor="FFFF00")

THIN_SIDE = Side(style="thin", color="000000")
MEDIUM_SIDE = Side(style="medium", color="000000")
NO_SIDE = Side(style=None)

SUMMARY_SHEET_NAME = "テーブル検証"
COMPARE_EXCLUDE_SHEET_NAME = "現新比較対象外"
TRIANGLE_MARK = "△"

warnings.filterwarnings(
    "ignore",
    message="Conditional Formatting extension is not supported and will be removed",
)


@dataclass
class CsvInfo:
    path: Path | None
    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    exists: bool = False
    invalid: bool = False

    @property
    def count(self) -> int:
        return len(self.rows)


@dataclass
class TableRecord:
    key: str
    display_name: str
    sql_name: str
    before: CsvInfo
    after: CsvInfo
    header: list[str]
    col_count: int
    crud: str
    count_diff: int
    notes: list[str]
    skip_detail: bool
    output_no: int | None = None
    output_sheet: str = ""
    detail_refs: dict = field(default_factory=dict)


# ===== CSV =====
def _normalize_csv_rows(rows):
    return [["" if v is None else str(v) for v in row] for row in rows]


def _infer_delimiter(text):
    sample_lines = []
    for line in text.splitlines():
        if line.strip():
            sample_lines.append(line)
        if len(sample_lines) >= 20:
            break
    sample = "\n".join(sample_lines)
    if sample:
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters="".join(VALID_DELIMITERS))
            delimiter = getattr(dialect, "delimiter", None)
            if isinstance(delimiter, str) and len(delimiter) == 1 and delimiter in VALID_DELIMITERS:
                return delimiter
        except (csv.Error, TypeError, ValueError):
            pass
    best_delimiter = ","
    best_count = -1
    for delimiter in VALID_DELIMITERS:
        count = sum(line.count(delimiter) for line in sample_lines)
        if count > best_count:
            best_count = count
            best_delimiter = delimiter
    return best_delimiter


def _read_csv_from_text(text):
    delimiter = _infer_delimiter(text)
    try:
        return list(csv.reader(io.StringIO(text), delimiter=delimiter))
    except (csv.Error, ValueError):
        return list(csv.reader(io.StringIO(text), delimiter=","))


def _decode_csv_text(path: Path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp932", "shift_jis", "utf-8", "gb18030"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("cp932", errors="replace"), "cp932-replace"


def read_csv_info(path: Path | None) -> CsvInfo:
    if path is None:
        return CsvInfo(path=None, exists=False)
    text, _enc = _decode_csv_text(path)
    rows = _normalize_csv_rows(_read_csv_from_text(text))
    header = rows[0] if rows else []
    invalid = INVALID_MARKER in text
    data = [] if invalid else rows[1:]
    return CsvInfo(path=path, header=header, rows=data, exists=True, invalid=invalid)


def csv_files(folder: Path):
    if not folder.exists():
        raise FileNotFoundError(f"フォルダが存在しません: {folder}")
    return {p.name.lower(): p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".csv"}


# ===== template / style =====
def table_name_from_file(path: Path, mode: str):
    return path.name if mode == "filename" else path.stem


def row_text(ws, row):
    values = []
    max_c = min(max(ws.max_column, TEMPLATE_COLS), TEMPLATE_COLS)
    for c in range(1, max_c + 1):
        v = ws.cell(row, c).value
        if v is not None:
            values.append(str(v))
    return " ".join(values)


def find_after_sql_template_row(tws):
    max_r = min(tws.max_row, 25)
    for r in range(14, max_r + 1):
        text = row_text(tws, r)
        upper = text.upper()
        if "SELECT" in upper and ("WHERE" in upper or "抽出条件" in text or "テーブル" in text):
            return r
    text16 = row_text(tws, 16)
    if "更新" in text16 and "件数" in text16:
        return 17
    return 16


def template_layout(tws):
    after_sql = find_after_sql_template_row(tws)
    return {
        "after_sql": after_sql,
        "after_header": after_sql + 1,
        "after_data": after_sql + 2,
        "after_new": after_sql + 5,
        "after_compare": after_sql + 8,
        "count_title": after_sql + 11,
        "before_count": after_sql + 12,
        "after_count": after_sql + 13,
        "count_diff": after_sql + 14,
        "pre_cross_title": after_sql + 15,
        "cross_title": after_sql + 16,
        "cross_current": after_sql + 17,
        "cross_new": after_sql + 20,
    }


def is_update_count_template_row(tws, row):
    text = row_text(tws, row)
    return "更新" in text and "件数" in text


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def _copy_style_component(cell, attr_name, default_value):
    try:
        value = getattr(cell, attr_name)
        return copy(value) if value is not None else copy(default_value)
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        return copy(default_value)


def force_font_name(cell, font_name=FONT_NAME):
    if isinstance(cell, MergedCell):
        return
    f = _copy_style_component(cell, "font", Font())
    f.name = font_name
    cell.font = f
    a = _copy_style_component(cell, "alignment", Alignment())
    a.vertical = "center"
    cell.alignment = a


def copy_cell_format(src, dst):
    if isinstance(dst, MergedCell):
        return
    # workbookを跨ぐ場合に _style を直接コピーすると、openpyxl の style index が壊れて
    # cell.border 参照時に IndexError になることがあるため、各style要素を値としてコピーする。
    dst.font = _copy_style_component(src, "font", Font(name=FONT_NAME))
    dst.fill = _copy_style_component(src, "fill", PatternFill(fill_type=None))
    dst.border = _copy_style_component(src, "border", Border())
    dst.alignment = _copy_style_component(src, "alignment", Alignment(vertical="center"))
    try:
        dst.number_format = src.number_format or "General"
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        dst.number_format = "General"
    try:
        dst.protection = copy(src.protection)
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        pass
    try:
        if src.comment:
            dst.comment = copy(src.comment)
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        pass
    force_font_name(dst)


def clear_cell_format(cell):
    if isinstance(cell, MergedCell):
        return
    cell.value = None
    cell.border = Border()
    cell.fill = PatternFill(fill_type=None)
    cell.font = Font(name=FONT_NAME)
    cell.alignment = Alignment(vertical="center")
    cell.number_format = "General"


def clear_extra_columns(ws, row, first_extra_col, last_col):
    if first_extra_col > last_col:
        return
    for c in range(first_extra_col, last_col + 1):
        clear_cell_format(ws.cell(row, c))


def template_cell_for_col(tws, template_row, dest_col):
    return tws.cell(template_row, dest_col if dest_col <= TEMPLATE_COLS else TEMPLATE_COLS)


def translate_formula(formula, origin, target):
    try:
        return Translator(formula, origin=origin).translate_formula(target)
    except Exception:
        return formula


def translate_cf_formula(formula, origin, target):
    if not isinstance(formula, str):
        return formula
    if not re.search(r"\$?[A-Za-z]{1,3}\$?\d+", formula):
        return formula
    had_equal = formula.startswith("=")
    expr = formula if had_equal else "=" + formula
    try:
        translated = Translator(expr, origin=origin).translate_formula(target)
        return translated if had_equal else translated[1:]
    except Exception:
        return formula


def copy_conditional_formatting_for_row(tws, ows, template_row, out_row, block_last_col):
    cf_rules = getattr(tws.conditional_formatting, "_cf_rules", None)
    if not cf_rules:
        return
    for cf_obj, rules in cf_rules.items():
        for src_range in cf_obj.sqref.ranges:
            if not (src_range.min_row <= template_row <= src_range.max_row):
                continue
            dest_min_col = src_range.min_col
            dest_max_col = block_last_col if src_range.max_col >= TEMPLATE_COLS else min(src_range.max_col, block_last_col)
            if dest_min_col > block_last_col or dest_min_col > dest_max_col:
                continue
            dest_ref = f"{get_column_letter(dest_min_col)}{out_row}:{get_column_letter(dest_max_col)}{out_row}"
            if dest_min_col == dest_max_col:
                dest_ref = f"{get_column_letter(dest_min_col)}{out_row}"
            src_anchor = f"{get_column_letter(src_range.min_col)}{template_row}"
            dst_anchor = f"{get_column_letter(dest_min_col)}{out_row}"
            for rule in rules:
                new_rule = copy(rule)
                if getattr(rule, "formula", None):
                    new_rule.formula = [translate_cf_formula(f, src_anchor, dst_anchor) for f in rule.formula]
                ows.conditional_formatting.add(dest_ref, new_rule)


def copy_template_row_style(tws, ows, template_row, out_row, block_last_col):
    for c in range(1, block_last_col + 1):
        copy_cell_format(template_cell_for_col(tws, template_row, c), ows.cell(out_row, c))
    copy_conditional_formatting_for_row(tws, ows, template_row, out_row, block_last_col)


def set_cell_border(cell, left=None, right=None, top=None, bottom=None):
    if isinstance(cell, MergedCell):
        return
    try:
        b = copy(cell.border)
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        b = Border()
    cell.border = Border(
        left=left if left is not None else b.left,
        right=right if right is not None else b.right,
        top=top if top is not None else b.top,
        bottom=bottom if bottom is not None else b.bottom,
        diagonal=b.diagonal,
        diagonal_direction=b.diagonal_direction,
        diagonalUp=b.diagonalUp,
        diagonalDown=b.diagonalDown,
        outline=b.outline,
        vertical=b.vertical,
        horizontal=b.horizontal,
    )


def apply_standard_grid(ws, row_start, row_count, col_start, col_end):
    if row_count <= 0 or col_end < col_start:
        return
    for r in range(row_start, row_start + row_count):
        for c in range(col_start, col_end + 1):
            set_cell_border(ws.cell(r, c), left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def apply_horizontal_border(ws, row, col_start, col_end, side, where="top"):
    if col_end < col_start:
        return
    for c in range(col_start, col_end + 1):
        if where == "bottom":
            set_cell_border(ws.cell(row, c), bottom=side)
        else:
            set_cell_border(ws.cell(row, c), top=side)


def clear_horizontal_border(ws, row, col_start, col_end, where="top"):
    if col_end < col_start:
        return
    for c in range(col_start, col_end + 1):
        if where == "bottom":
            set_cell_border(ws.cell(row, c), bottom=NO_SIDE)
        else:
            set_cell_border(ws.cell(row, c), top=NO_SIDE)


def apply_title_row_fill(tws, ows, out_row, block_last_col):
    src_fill = copy(tws.cell(1, 1).fill)
    for c in range(1, block_last_col + 1):
        ows.cell(out_row, c).fill = copy(src_fill)


def write_static_from_template(tws, ows, template_row, out_row, block_last_col):
    copy_template_row_style(tws, ows, template_row, out_row, block_last_col)
    for c in range(1, min(TEMPLATE_COLS, block_last_col) + 1):
        src = tws.cell(template_row, c)
        if src.value is not None:
            dst = ows.cell(out_row, c)
            dst.value = translate_formula(src.value, src.coordinate, dst.coordinate) if is_formula(src.value) else src.value
            force_font_name(dst)


def write_static_data_limited_from_template(tws, ows, template_row, out_row, data_last_col, block_last_col):
    write_static_from_template(tws, ows, template_row, out_row, block_last_col)
    clear_extra_columns(ows, out_row, data_last_col + 1, block_last_col)


def merge_b_label(ows, row_start, row_count, label):
    if row_count <= 0:
        return
    ows.cell(row_start, 2).value = label
    force_font_name(ows.cell(row_start, 2))
    if row_count > 1:
        ows.merge_cells(start_row=row_start, start_column=2, end_row=row_start + row_count - 1, end_column=2)


def fill_gray_row(ows, row, col_start, col_end):
    for c in range(col_start, col_end + 1):
        ows.cell(row, c).fill = copy(GRAY_FILL)


def build_compare_formula(row_a, row_b, col, table_aliases=None, item_name=None, compare_exclusions=None, blank_skip=True):
    letter = get_column_letter(col)
    exact_expr = f"EXACT({letter}{row_a},{letter}{row_b})"
    if is_compare_excluded(table_aliases or set(), item_name, compare_exclusions or set()):
        body = f'IF({exact_expr},TRUE,"{TRIANGLE_MARK}")'
    else:
        body = exact_expr
    if blank_skip:
        return f'=IF(AND({letter}{row_a}="",{letter}{row_b}=""),"",{body})'
    return f"={body}"


def build_result_formula(row, first_col, last_col):
    return f'=IF(COUNTIF({get_column_letter(first_col)}{row}:{get_column_letter(last_col)}{row},FALSE)>0,"差分あり","差分なし")'


def prepare_columns(tws, ows, block_last_col):
    for c in range(1, block_last_col + 1):
        letter = get_column_letter(c)
        src_letter = letter if c <= TEMPLATE_COLS else get_column_letter(TEMPLATE_COLS)
        ows.column_dimensions[letter].width = tws.column_dimensions[src_letter].width or 12


def normalize_matrix(rows, row_count, col_count):
    out = []
    for i in range(row_count):
        src = rows[i] if i < len(rows) else []
        out.append(list(src[:col_count]) + [""] * max(0, col_count - len(src)))
    return out


def add_compare_conditional_formatting(ws, row_start, row_count, data_last_col):
    if row_count <= 0 or data_last_col < DATA_START_COL:
        return
    ref = f"{get_column_letter(DATA_START_COL)}{row_start}:{get_column_letter(data_last_col)}{row_start + row_count - 1}"
    ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=["TRUE"], fill=TRUE_FILL))
    ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=["FALSE"], fill=FALSE_FILL, font=FALSE_FONT))
    ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=[f'"{TRIANGLE_MARK}"'], fill=WARNING_FILL))


def write_header_row(tws, ows, template_row, out_row, header, col_count, block_last_col, header_rows):
    data_last_col = DATA_START_COL + col_count - 1
    copy_template_row_style(tws, ows, template_row, out_row, block_last_col)
    for idx in range(col_count):
        cell = ows.cell(out_row, DATA_START_COL + idx)
        cell.value = header[idx] if idx < len(header) else ""
        force_font_name(cell)
    clear_extra_columns(ows, out_row, data_last_col + 1, block_last_col)
    apply_standard_grid(ows, out_row, 1, DATA_START_COL, data_last_col)
    header_rows.add(out_row)


def write_data_rows(tws, ows, template_row, start_row, rows, display_count, col_count, block_last_col, gray_if_empty=False):
    data_last_col = DATA_START_COL + col_count - 1
    matrix = normalize_matrix(rows, display_count, col_count)
    for i in range(display_count):
        r = start_row + i
        copy_template_row_style(tws, ows, template_row, r, block_last_col)
        for idx in range(col_count):
            cell = ows.cell(r, DATA_START_COL + idx)
            cell.value = matrix[i][idx]
            force_font_name(cell)
        clear_extra_columns(ows, r, data_last_col + 1, block_last_col)
        if gray_if_empty:
            fill_gray_row(ows, r, DATA_START_COL, data_last_col)
    apply_standard_grid(ows, start_row, display_count, DATA_START_COL, data_last_col)


def write_blank_rows(tws, ows, template_row, start_row, display_count, col_count, block_last_col):
    data_last_col = DATA_START_COL + col_count - 1
    for i in range(display_count):
        r = start_row + i
        copy_template_row_style(tws, ows, template_row, r, block_last_col)
        for idx in range(col_count):
            cell = ows.cell(r, DATA_START_COL + idx)
            cell.value = None
            force_font_name(cell)
        clear_extra_columns(ows, r, data_last_col + 1, block_last_col)
    apply_standard_grid(ows, start_row, display_count, DATA_START_COL, data_last_col)


def write_vertical_compare_rows(tws, ows, template_row, start_row, count, row_a_start, row_b_start, col_count, block_last_col, table_aliases=None, header=None, compare_exclusions=None):
    data_last_col = DATA_START_COL + col_count - 1
    table_aliases = table_aliases or set()
    header = header or []
    compare_exclusions = compare_exclusions or set()
    for i in range(count):
        r = start_row + i
        copy_template_row_style(tws, ows, template_row, r, block_last_col)
        ows.cell(r, 2).value = build_result_formula(r, DATA_START_COL, data_last_col)
        force_font_name(ows.cell(r, 2))
        for idx in range(col_count):
            c = DATA_START_COL + idx
            item_name = header[idx] if idx < len(header) else ""
            cell = ows.cell(r, c)
            cell.value = build_compare_formula(row_a_start + i, row_b_start + i, c, table_aliases, item_name, compare_exclusions, blank_skip=True)
            force_font_name(cell)
        clear_extra_columns(ows, r, data_last_col + 1, block_last_col)
    apply_standard_grid(ows, start_row, count, DATA_START_COL, data_last_col)
    add_compare_conditional_formatting(ows, start_row, count, data_last_col)


def write_cross_compare_rows(tws, ows, template_row, start_row, count, row_a_start, row_b_start, label, col_count, block_last_col, table_aliases=None, header=None, compare_exclusions=None):
    data_last_col = DATA_START_COL + col_count - 1
    table_aliases = table_aliases or set()
    header = header or []
    compare_exclusions = compare_exclusions or set()
    for i in range(count):
        r = start_row + i
        copy_template_row_style(tws, ows, template_row, r, block_last_col)
        for idx in range(col_count):
            c = DATA_START_COL + idx
            item_name = header[idx] if idx < len(header) else ""
            cell = ows.cell(r, c)
            cell.value = build_compare_formula(row_a_start + i, row_b_start + i, c, table_aliases, item_name, compare_exclusions, blank_skip=False)
            force_font_name(cell)
        clear_extra_columns(ows, r, data_last_col + 1, block_last_col)
    apply_standard_grid(ows, start_row, count, DATA_START_COL, data_last_col)
    add_compare_conditional_formatting(ows, start_row, count, data_last_col)
    merge_b_label(ows, start_row, count, label)


# ===== records / block =====
def build_table_records(before_map, after_map, sql_name_mode):
    keys = sorted(set(before_map.keys()) | set(after_map.keys()))
    records = []
    for key in keys:
        before = read_csv_info(before_map.get(key))
        after = read_csv_info(after_map.get(key))
        display_path = before.path or after.path
        display_name = display_path.stem if display_path else key[:-4] if key.lower().endswith(".csv") else key
        sql_name = table_name_from_file(display_path, sql_name_mode) if display_path else display_name

        notes = []
        if not before.exists:
            notes.append("打鍵前CSVなし。0件扱い")
        if not after.exists:
            notes.append("打鍵後CSVなし。0件扱い")
        if before.invalid:
            notes.append("打鍵前CSVに無効な識別あり。0件扱い")
        if after.invalid:
            notes.append("打鍵後CSVに無効な識別あり。0件扱い")

        before_header = before.header[:]
        after_header = after.header[:]
        if not before_header:
            before_header = after_header[:]
        if not after_header:
            after_header = before_header[:]
        if not before.header:
            before.header = before_header
        if not after.header:
            after.header = after_header

        max_cols = max(len(before_header), len(after_header), *(len(r) for r in before.rows), *(len(r) for r in after.rows), 1)
        header = before_header if before_header else after_header
        header = list(header[:max_cols]) + [""] * max(0, max_cols - len(header))

        diff = after.count - before.count
        if after.count > before.count:
            crud = "C"
        elif after.count < before.count:
            crud = "D"
        elif before.count != 0:
            crud = "U"
        else:
            crud = ""

        skip_detail = before.count == 0 and after.count == 0
        if skip_detail:
            notes.append("打鍵前後とも0件のため明細出力なし")

        records.append(TableRecord(key, display_name, sql_name, before, after, header, max_cols, crud, diff, notes, skip_detail))
    return records


def write_block(tws, ows, start_row, record, layout, header_rows, cf_targets, compare_exclusions=None):
    before_header = record.before.header[:] if record.before.header else record.header[:]
    after_header = record.after.header[:] if record.after.header else record.header[:]
    before_rows = record.before.rows
    after_rows = record.after.rows
    col_count = record.col_count
    data_last_col = DATA_START_COL + col_count - 1
    block_last_col = max(TEMPLATE_COLS, data_last_col)
    table_aliases = make_table_aliases(record.display_name, record.sql_name, record.key)
    prepare_columns(tws, ows, block_last_col)

    before_count = len(before_rows)
    after_count = len(after_rows)
    before_disp = max(1, before_count)
    after_disp = max(1, after_count)
    compare_b_ranges = []
    r = start_row

    write_static_from_template(tws, ows, 1, r, block_last_col)
    ows.cell(r, 1).value = f"{record.output_no}.{record.display_name}" if record.output_no is not None else record.display_name
    force_font_name(ows.cell(r, 1))
    apply_title_row_fill(tws, ows, r, block_last_col)
    header_rows.add(r)
    r += 1

    write_static_from_template(tws, ows, 2, r, block_last_col)
    r += 1
    write_static_from_template(tws, ows, 3, r, block_last_col)
    ows.cell(r, 3).value = f"SELECT * FROM {record.sql_name} WHERE 抽出条件"
    force_font_name(ows.cell(r, 3))
    r += 1
    write_header_row(tws, ows, 4, r, before_header, col_count, block_last_col, header_rows)
    r += 1

    before_curr_start = r
    write_data_rows(tws, ows, 5, before_curr_start, before_rows, before_disp, col_count, block_last_col, gray_if_empty=(before_count == 0))
    merge_b_label(ows, before_curr_start, before_disp, "現")
    r += before_disp

    before_new_start = r
    write_blank_rows(tws, ows, 8, before_new_start, before_disp, col_count, block_last_col)
    merge_b_label(ows, before_new_start, before_disp, "新")
    apply_horizontal_border(ows, before_new_start, 2, data_last_col, MEDIUM_SIDE, where="top")
    clear_horizontal_border(ows, before_new_start, data_last_col + 1, block_last_col, where="top")
    r += before_disp

    before_cmp_start = r
    write_vertical_compare_rows(tws, ows, 11, before_cmp_start, before_disp, before_curr_start, before_new_start, col_count, block_last_col, table_aliases, before_header, compare_exclusions)
    cf_targets.append({"kind": "before_compare", "start": before_cmp_start, "end": before_cmp_start + before_disp - 1, "data_last_col": data_last_col})
    compare_b_ranges.append((before_cmp_start, before_cmp_start + before_disp - 1))
    r += before_disp

    for tmpl_row in range(14, layout["after_sql"]):
        if is_update_count_template_row(tws, tmpl_row):
            continue
        write_static_from_template(tws, ows, tmpl_row, r, block_last_col)
        r += 1

    write_static_from_template(tws, ows, layout["after_sql"], r, block_last_col)
    ows.cell(r, 3).value = f"SELECT * FROM {record.sql_name} WHERE 抽出条件"
    force_font_name(ows.cell(r, 3))
    r += 1
    write_header_row(tws, ows, layout["after_header"], r, after_header, col_count, block_last_col, header_rows)
    r += 1

    after_curr_start = r
    write_data_rows(tws, ows, layout["after_data"], after_curr_start, after_rows, after_disp, col_count, block_last_col, gray_if_empty=(after_count == 0))
    merge_b_label(ows, after_curr_start, after_disp, "現")
    r += after_disp

    after_new_start = r
    write_blank_rows(tws, ows, layout["after_new"], after_new_start, after_disp, col_count, block_last_col)
    merge_b_label(ows, after_new_start, after_disp, "新")
    apply_horizontal_border(ows, after_new_start, 2, data_last_col, MEDIUM_SIDE, where="top")
    clear_horizontal_border(ows, after_new_start, data_last_col + 1, block_last_col, where="top")
    r += after_disp

    after_cmp_start = r
    write_vertical_compare_rows(tws, ows, layout["after_compare"], after_cmp_start, after_disp, after_curr_start, after_new_start, col_count, block_last_col, table_aliases, after_header, compare_exclusions)
    cf_targets.append({"kind": "after_compare", "start": after_cmp_start, "end": after_cmp_start + after_disp - 1, "data_last_col": data_last_col})
    compare_b_ranges.append((after_cmp_start, after_cmp_start + after_disp - 1))
    r += after_disp

    write_static_from_template(tws, ows, layout["count_title"], r, block_last_col)
    r += 1
    write_static_from_template(tws, ows, layout["before_count"], r, block_last_col)
    ows.cell(r, 3).value = before_count
    force_font_name(ows.cell(r, 3))
    before_count_row = r
    r += 1
    write_static_from_template(tws, ows, layout["after_count"], r, block_last_col)
    ows.cell(r, 3).value = after_count
    force_font_name(ows.cell(r, 3))
    after_count_row = r
    r += 1
    write_static_from_template(tws, ows, layout["count_diff"], r, block_last_col)
    ows.cell(r, 3).value = f"=C{after_count_row}-C{before_count_row}"
    force_font_name(ows.cell(r, 3))
    count_diff = after_count - before_count
    r += 1

    write_static_from_template(tws, ows, layout["pre_cross_title"], r, block_last_col)
    r += 1

    if count_diff == 0:
        write_static_data_limited_from_template(tws, ows, layout["cross_title"], r, data_last_col, block_last_col)
        r += 1
        cross_count = max(1, before_count)
        cross_current_start = r
        write_cross_compare_rows(tws, ows, layout["cross_current"], cross_current_start, cross_count, before_curr_start, after_curr_start, "現", col_count, block_last_col, table_aliases, before_header, compare_exclusions)
        r += cross_count
        cross_new_start = r
        write_cross_compare_rows(tws, ows, layout["cross_new"], cross_new_start, cross_count, before_new_start, after_new_start, "新", col_count, block_last_col, table_aliases, before_header, compare_exclusions)
        apply_horizontal_border(ows, cross_new_start, 2, data_last_col, MEDIUM_SIDE, where="top")
        clear_horizontal_border(ows, cross_new_start, data_last_col + 1, block_last_col, where="top")
        r += cross_count
        cf_targets.append({"kind": "cross_compare", "start": cross_current_start, "end": cross_new_start + cross_count - 1, "data_last_col": data_last_col})
        compare_b_ranges.append((cross_current_start, cross_new_start + cross_count - 1))

    record.detail_refs = {
        "before_count_row": before_count_row,
        "after_count_row": after_count_row,
        "before_new_start": before_new_start,
        "before_new_end": before_new_start + before_disp - 1,
        "after_new_start": after_new_start,
        "after_new_end": after_new_start + after_disp - 1,
        "data_last_col": data_last_col,
        "compare_b_ranges": compare_b_ranges,
    }
    return r, block_last_col


# ===== x14 条件付き書式 =====
def resolve_sheet_xml_path(xlsx_path: Path, sheet_name=None):
    with zipfile.ZipFile(xlsx_path, "r") as z:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        sheets_el = wb_root.find(f"{ns_main}sheets")
        sheets = list(sheets_el) if sheets_el is not None else []
        target_sheet = None
        for sheet in sheets:
            if sheet_name is None or sheet.attrib.get("name") == sheet_name:
                target_sheet = sheet
                break
        if target_sheet is None:
            return None
        rid = target_sheet.attrib.get(f"{ns_rel}id")
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        for rel in rels_root.findall(f"{rel_ns}Relationship"):
            if rel.attrib.get("Id") == rid:
                target = rel.attrib.get("Target", "")
                if target.startswith("/"):
                    return target.lstrip("/")
                return str(PurePosixPath("xl") / target)
    raise RuntimeError("対象 worksheet XML を特定できません。")


def parse_sqref_rows(sqref):
    first = (sqref or "").split()[0]
    rows = [int(x) for x in re.findall(r"\$?[A-Za-z]{1,3}\$?(\d+)", first)]
    if not rows:
        return None, None
    return min(rows), max(rows)


def extract_x14_compare_cf_templates(template_path: Path, layout, sheet_name=None):
    result = {}
    try:
        sheet_xml_path = resolve_sheet_xml_path(template_path, sheet_name)
        if not sheet_xml_path:
            return result
        with zipfile.ZipFile(template_path, "r") as z:
            xml = z.read(sheet_xml_path).decode("utf-8")
    except Exception:
        return result
    blocks = re.findall(r"<x14:conditionalFormatting\b[\s\S]*?</x14:conditionalFormatting>", xml)
    for block in blocks:
        m = re.search(r"<xm:sqref>(.*?)</xm:sqref>", block)
        if not m:
            continue
        min_row, max_row = parse_sqref_rows(m.group(1))
        if min_row is None:
            continue
        item = {"block": block, "src_min_row": min_row, "src_max_row": max_row}
        if min_row <= 11 <= max_row:
            result.setdefault("before_compare", item)
        if min_row <= layout["after_compare"] <= max_row:
            result.setdefault("after_compare", item)
        if min_row <= layout["cross_current"] <= max_row or min_row <= layout["cross_new"] <= max_row:
            result.setdefault("cross_compare", item)
    return result


def render_x14_cf_block(template_item, dest_start, dest_end, priority_start):
    block = template_item["block"]
    src_anchor = f"A{template_item['src_min_row']}"
    dst_anchor = f"A{dest_start}"
    priority = priority_start
    block = block.replace("<x14:conditionalFormattings>", "").replace("</x14:conditionalFormattings>", "")

    def repl_formula(m):
        return f"<xm:f>{translate_cf_formula(m.group(1), src_anchor, dst_anchor)}</xm:f>"

    def repl_rule(m):
        nonlocal priority
        rule = m.group(0)
        rule = re.sub(r'priority="\d+"', f'priority="{priority}"', rule, count=1)
        priority += 1
        rule = re.sub(r'id="\{[^\"]+\}"', f'id="{{{str(uuid.uuid4()).upper()}}}"', rule, count=1)
        rule = re.sub(r"<xm:f>(.*?)</xm:f>", repl_formula, rule)
        return rule

    block = re.sub(r"<x14:cfRule\b[\s\S]*?</x14:cfRule>", repl_rule, block)
    block = re.sub(r"<xm:sqref>.*?</xm:sqref>", f"<xm:sqref>A{dest_start}:XFD{dest_end}</xm:sqref>", block, count=1)
    return block, priority


def inject_x14_conditional_formatting_for_sheet(xlsx_path: Path, template_path: Path, sheet_name, cf_targets, layout, template_sheet_name=None):
    if not cf_targets:
        return False
    templates = extract_x14_compare_cf_templates(template_path, layout, template_sheet_name)
    if not templates:
        return False
    blocks = []
    priority = 3
    for target in cf_targets:
        tmpl = templates.get(target["kind"])
        if not tmpl:
            continue
        block, priority = render_x14_cf_block(tmpl, target["start"], target["end"], priority)
        blocks.append(block)
    if not blocks:
        return False
    ext_xml = ('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" '
               'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
               '<x14:conditionalFormattings>' + "".join(blocks) + '</x14:conditionalFormattings></ext>')
    sheet_xml_path = resolve_sheet_xml_path(xlsx_path, sheet_name)
    if not sheet_xml_path:
        return False
    tmp_dir = xlsx_path.parent if xlsx_path.parent.exists() else Path(".")
    tmp = tempfile.NamedTemporaryFile(prefix=f".{xlsx_path.stem}_", suffix=".tmp.xlsx", dir=tmp_dir, delete=False)
    tmp_name = tmp.name
    tmp.close()
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == sheet_xml_path:
                    xml = data.decode("utf-8")
                    if "</extLst>" in xml:
                        xml = xml.replace("</extLst>", ext_xml + "</extLst>", 1)
                    else:
                        xml = xml.replace("</worksheet>", "<extLst>" + ext_xml + "</extLst></worksheet>", 1)
                    data = xml.encode("utf-8")
                zout.writestr(info, data)
        os.replace(tmp_name, xlsx_path)
        return True
    finally:
        try:
            Path(tmp_name).unlink(missing_ok=True)
        except (FileNotFoundError, PermissionError):
            pass


def inject_x14_conditional_formatting(xlsx_path: Path, template_path: Path, sheet_targets: dict[str, list[dict]], layout, template_sheet_name=None):
    for sheet_name, targets in sheet_targets.items():
        inject_x14_conditional_formatting_for_sheet(xlsx_path, template_path, sheet_name, targets, layout, template_sheet_name)


# ===== summary =====
def quote_sheet_name(sheet_name):
    return "'" + str(sheet_name).replace("'", "''") + "'"


def _normalize_compare_key(value):
    text = "" if value is None else str(value)
    return text.strip().lower()


def _normalize_table_key(value):
    text = _normalize_compare_key(value)
    if text.endswith(".csv"):
        text = text[:-4]
    return text


def read_compare_exclusions(wb):
    exclusions = set()
    if COMPARE_EXCLUDE_SHEET_NAME not in wb.sheetnames:
        return exclusions

    ws = wb[COMPARE_EXCLUDE_SHEET_NAME]
    table_col = 1
    item_col = 2
    for c in range(1, min(ws.max_column, 20) + 1):
        header = _normalize_compare_key(ws.cell(1, c).value)
        if header in {"テーブル名", "table", "table_name", "表名"}:
            table_col = c
        elif header in {"項目名", "项目名", "column", "column_name", "field", "field_name"}:
            item_col = c

    for r in range(2, ws.max_row + 1):
        table_name = ws.cell(r, table_col).value
        item_name = ws.cell(r, item_col).value
        item_key = _normalize_compare_key(item_name)
        if not item_key:
            continue
        table_key = _normalize_table_key(table_name)
        if not table_key:
            table_key = "*"
        exclusions.add((table_key, item_key))
    return exclusions


def make_table_aliases(*values):
    aliases = set()
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        aliases.add(_normalize_table_key(text))
        aliases.add(_normalize_compare_key(text))
        if text.lower().endswith(".csv"):
            aliases.add(_normalize_table_key(text[:-4]))
    return aliases


def is_compare_excluded(table_aliases, item_name, compare_exclusions):
    item_key = _normalize_compare_key(item_name)
    if not item_key or not compare_exclusions:
        return False
    if ("*", item_key) in compare_exclusions:
        return True
    aliases = table_aliases if isinstance(table_aliases, set) else make_table_aliases(*(table_aliases or []))
    return any((alias, item_key) in compare_exclusions for alias in aliases)


def xl_abs_cell(sheet_name, row, col):
    return f"{quote_sheet_name(sheet_name)}!${get_column_letter(col)}${row}"


def xl_abs_range(sheet_name, row1, col1, row2, col2):
    return f"{quote_sheet_name(sheet_name)}!${get_column_letter(col1)}${row1}:${get_column_letter(col2)}${row2}"


def build_non_empty_row_count_formula(sheet_name, row1, row2, col1, col2):
    data_range = xl_abs_range(sheet_name, row1, col1, row2, col2)
    first_row_range = xl_abs_range(sheet_name, row1, col1, row1, col2)
    return ("=IFERROR("
            f"SUMPRODUCT(--(MMULT(--({data_range}<>\"\"),TRANSPOSE(COLUMN({first_row_range})^0))>0)),"
            "0)")


def build_countif_diff_formula(ranges):
    if not ranges:
        return '="OK"'
    parts = [f'COUNTIF({rng},"差分あり")' for rng in ranges]
    return f'=IF(({"+".join(parts)})>0,"NG","OK")'


def write_summary_sheet(wb, records, sheet_index=0):
    for old_name in (SUMMARY_SHEET_NAME, "summary"):
        if old_name in wb.sheetnames:
            del wb[old_name]
    safe_index = min(max(0, sheet_index or 0), len(wb.worksheets))
    ws = wb.create_sheet(SUMMARY_SHEET_NAME, safe_index)
    headers = [
        "CRUD", "出力番号", "出力Sheet", "テーブル名",
        "現_打鍵前件数", "現_打鍵後件数", "現_打鍵前後件数差異",
        "新_打鍵前件数", "新_打鍵後件数", "新_打鍵前後件数差異",
        "現新件数チェック", "差分チェック", "備考",
    ]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.fill = copy(SUMMARY_HEADER_FILL)
        cell.font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    for r_idx, rec in enumerate(records, start=2):
        refs = rec.detail_refs or {}
        sheet_name = rec.output_sheet
        if refs:
            gen_before_formula = f"=IFERROR({xl_abs_cell(sheet_name, refs['before_count_row'], 3)},0)"
            gen_after_formula = f"=IFERROR({xl_abs_cell(sheet_name, refs['after_count_row'], 3)},0)"
            new_before_formula = build_non_empty_row_count_formula(sheet_name, refs["before_new_start"], refs["before_new_end"], DATA_START_COL, refs["data_last_col"])
            new_after_formula = build_non_empty_row_count_formula(sheet_name, refs["after_new_start"], refs["after_new_end"], DATA_START_COL, refs["data_last_col"])
            diff_ranges = [xl_abs_range(sheet_name, s, 2, e, 2) for s, e in refs.get("compare_b_ranges", []) if s and e and e >= s]
            diff_check_formula = build_countif_diff_formula(diff_ranges)
        else:
            gen_before_formula = "=0"
            gen_after_formula = "=0"
            new_before_formula = "=0"
            new_after_formula = "=0"
            diff_check_formula = '="OK"'

        values = [
            f'=IF(F{r_idx}>E{r_idx},"C",IF(F{r_idx}<E{r_idx},"D",IF(AND(F{r_idx}=E{r_idx},E{r_idx}<>0),"U","")))',
            rec.output_no if rec.output_no is not None else "",
            rec.output_sheet,
            rec.display_name,
            gen_before_formula,
            gen_after_formula,
            f"=F{r_idx}-E{r_idx}",
            new_before_formula,
            new_after_formula,
            f"=I{r_idx}-H{r_idx}",
            f'=IF(AND(E{r_idx}=H{r_idx},F{r_idx}=I{r_idx}),"差分なし","差分あり")',
            diff_check_formula,
            " / ".join(rec.notes),
        ]
        ws.append(values)
        row_fill = {"C": SUMMARY_C_FILL, "U": SUMMARY_U_FILL, "D": SUMMARY_D_FILL, "": SUMMARY_SKIP_FILL}.get(rec.crud, SUMMARY_SKIP_FILL)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r_idx, c)
            cell.fill = copy(row_fill)
            cell.font = Font(name=FONT_NAME)
            cell.alignment = Alignment(vertical="center", wrap_text=(c == len(headers)))
            cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
        if rec.crud == "":
            ws.row_dimensions[r_idx].hidden = True

    ws.auto_filter.ref = f"A1:M{max(1, ws.max_row)}"
    ws.freeze_panes = "A2"
    widths = [10, 12, 16, 32, 18, 18, 22, 18, 18, 22, 18, 16, 70]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    if ws.max_row >= 2:
        count_check_ref = f"K2:K{ws.max_row}"
        data_check_ref = f"L2:L{ws.max_row}"
        ws.conditional_formatting.add(count_check_ref, CellIsRule(operator="equal", formula=['"差分なし"'], fill=TRUE_FILL))
        ws.conditional_formatting.add(count_check_ref, CellIsRule(operator="equal", formula=['"差分あり"'], fill=FALSE_FILL, font=FALSE_FONT))
        ws.conditional_formatting.add(data_check_ref, CellIsRule(operator="equal", formula=['"OK"'], fill=TRUE_FILL))
        ws.conditional_formatting.add(data_check_ref, CellIsRule(operator="equal", formula=['"NG"'], fill=FALSE_FILL, font=FALSE_FONT))
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = HEADER_ROW_HEIGHT if r == 1 else BODY_ROW_HEIGHT
    return ws


# ===== workbook =====
def normalize_fonts_and_row_heights(ws, max_row, max_col, header_rows):
    for r in range(1, max_row + 1):
        ws.row_dimensions[r].height = HEADER_ROW_HEIGHT if r in header_rows else BODY_ROW_HEIGHT
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                force_font_name(cell)


def copy_sheet_page_settings(tws, ows):
    try:
        ows.sheet_view.showGridLines = tws.sheet_view.showGridLines
    except Exception:
        pass
    try:
        ows.freeze_panes = tws.freeze_panes
    except Exception:
        pass
    for attr in ("page_margins", "page_setup", "print_options", "sheet_properties"):
        try:
            setattr(ows, attr, copy(getattr(tws, attr)))
        except Exception:
            pass


def unique_sheet_name(wb, base_name):
    if base_name not in wb.sheetnames:
        return base_name
    i = 1
    while f"{base_name}{i}" in wb.sheetnames:
        i += 1
    return f"{base_name}{i}"


def assign_output_positions(records, base_sheet_name, tables_per_sheet):
    output_records = [r for r in records if not r.skip_detail]
    multi = len(output_records) > tables_per_sheet
    for idx, rec in enumerate(output_records, start=1):
        rec.output_no = idx
        rec.output_sheet = f"{base_sheet_name}{(idx - 1) // tables_per_sheet + 1}" if multi else base_sheet_name
    return output_records



# ===== 設定ファイル / 出力制御 =====
DEFAULT_CONFIG_FILE = "設定ファイル.json"
ACTION_REPLACE_FILE = "replace_file"
ACTION_ERROR = "error"
ACTION_REPLACE_SHEETS = "replace_sheets"


def load_settings_file(config_path: Path | None):
    if config_path is None:
        config_path = Path(DEFAULT_CONFIG_FILE)
    if not config_path.is_absolute():
        local = Path.cwd() / config_path
        script_local = Path(__file__).resolve().parent / config_path
        config_path = local if local.exists() else script_local
    if not config_path.exists():
        return {}
    with config_path.open("r", encoding="utf-8-sig") as f:
        return json.load(f)


def cfg_get(*dicts, keys, default=None):
    for d in dicts:
        if not isinstance(d, dict):
            continue
        for key in keys:
            if key in d and d[key] not in (None, ""):
                return d[key]
    return default


def bool_value(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "有", "あり", "する", "はい"}


def normalize_existing_action(value):
    text = str(value or ACTION_REPLACE_FILE).strip().lower()
    mapping = {
        "replace": ACTION_REPLACE_FILE,
        "overwrite": ACTION_REPLACE_FILE,
        "replace_file": ACTION_REPLACE_FILE,
        "上書き": ACTION_REPLACE_FILE,
        "ファイル上書き": ACTION_REPLACE_FILE,
        "error": ACTION_ERROR,
        "raise": ACTION_ERROR,
        "エラー": ACTION_ERROR,
        "replace_sheets": ACTION_REPLACE_SHEETS,
        "sheet": ACTION_REPLACE_SHEETS,
        "sheets": ACTION_REPLACE_SHEETS,
        "シート差替": ACTION_REPLACE_SHEETS,
        "シート差し替え": ACTION_REPLACE_SHEETS,
    }
    if text not in mapping:
        raise ValueError(f"既存ファイル処理が不正です: {value}（replace_file / error / replace_sheets）")
    return mapping[text]


def normalize_path_value(value):
    """設定ファイル上のパス表記を正規化する。

    JSON では通常どおり "." / ".." を使えるほか、説明用に
    「現在フォルダ」「上一层文件夹」などを書いた場合も解釈する。
    """
    if value is None:
        return "."
    text = str(value).strip()
    aliases = {
        "": ".",
        "現在フォルダ": ".",
        "現在フォルダー": ".",
        "カレントフォルダ": ".",
        "カレントフォルダー": ".",
        "current": ".",
        "current_folder": ".",
        "当前文件夹": ".",
        "当前目录": ".",
        "上一层文件夹": "..",
        "上一層文件夾": "..",
        "上位フォルダ": "..",
        "上位フォルダー": "..",
        "親フォルダ": "..",
        "親フォルダー": "..",
        "parent": "..",
        "parent_folder": "..",
    }
    return aliases.get(text, text)


def join_if_relative(base, value):
    value = normalize_path_value(value)
    base = normalize_path_value(base)
    path = Path(str(value))
    if path.is_absolute():
        return path
    return Path(str(base or ".")) / path


def resolve_wildcard_path(folder, filename, must_exist=False):
    base = Path(str(normalize_path_value(folder or ".")))
    pattern_path = join_if_relative(base, filename)
    pattern = str(pattern_path)
    if any(ch in pattern for ch in "*?["):
        matches = sorted(Path(pattern_path.parent).glob(pattern_path.name))
        matches = [m for m in matches if m.is_file()]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            raise FileNotFoundError(f"ワイルドカードに一致するファイルがありません: {pattern}")
        raise RuntimeError("ワイルドカードに一致するファイルが複数あります: " + ", ".join(str(m) for m in matches))
    if must_exist and not pattern_path.exists():
        raise FileNotFoundError(f"ファイルが存在しません: {pattern_path}")
    return pattern_path

def east_asian_len(value):
    import unicodedata
    text = "" if value is None else str(value)
    total = 0
    for ch in text:
        width = unicodedata.east_asian_width(ch)
        total += 2 if width in {"F", "W"} else 1
    return total


def _is_sql_like_text(text):
    upper = text.upper()
    return "SELECT" in upper and (" FROM " in upper or " WHERE " in upper or "抽出条件" in text)


def _should_skip_for_autofit(ws, cell):
    value = cell.value
    if value in (None, ""):
        return True
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return True
        # 数式や取得SQLは列幅を極端に広げる原因になるため、列幅計算から除外する。
        if text.startswith("=") or _is_sql_like_text(text):
            return True
    return False


def _effective_autofit_len(value):
    text = "" if value is None else str(value).replace("\r", "\n")
    # 改行がある場合は最長行だけを見る。備考などの長文が列幅を押し広げすぎないようにする。
    parts = text.split("\n")
    return max((east_asian_len(part) for part in parts), default=0)


def auto_fit_column_widths(wb, enabled=False, max_width=60, min_width=8, sheet_names=None):
    if not enabled:
        return
    try:
        max_width = float(max_width)
        min_width = float(min_width)
    except Exception:
        max_width, min_width = 60, 8
    if max_width < min_width:
        max_width, min_width = min_width, max_width

    target_names = set(sheet_names or wb.sheetnames)
    for ws in wb.worksheets:
        if ws.title not in target_names:
            continue
        for col in range(1, ws.max_column + 1):
            max_len = 0
            for row in range(1, ws.max_row + 1):
                if getattr(ws.row_dimensions[row], "hidden", False):
                    continue
                cell = ws.cell(row, col)
                if isinstance(cell, MergedCell) or _should_skip_for_autofit(ws, cell):
                    continue
                max_len = max(max_len, _effective_autofit_len(cell.value))

            if max_len:
                # Excel の列幅は実文字数より少し余裕を持たせるだけにする。
                # 以前の max_len+2 は SQL 行や長い式に引っ張られて 40 超になりやすかったため抑制する。
                width = max_len * 1.08 + 1.5
                ws.column_dimensions[get_column_letter(col)].width = min(max_width, max(min_width, width))

def is_table_output_sheet(name, prefix):
    return name == prefix or re.fullmatch(re.escape(prefix) + r"\d+", name or "") is not None


def remove_existing_output_sheets(wb, table_prefix):
    indexes = {ws.title: idx for idx, ws in enumerate(wb.worksheets)}
    table_indexes = [idx for idx, ws in enumerate(wb.worksheets) if is_table_output_sheet(ws.title, table_prefix)]
    summary_index = indexes.get(SUMMARY_SHEET_NAME, indexes.get("summary", 0))
    table_index = min(table_indexes) if table_indexes else len(wb.worksheets)
    delete_titles = [ws.title for ws in wb.worksheets if is_table_output_sheet(ws.title, table_prefix) or ws.title in {SUMMARY_SHEET_NAME, "summary"}]
    if len(delete_titles) >= len(wb.worksheets):
        wb.create_sheet("__一時退避__")
    for title in delete_titles:
        if title in wb.sheetnames:
            del wb[title]
    return min(table_index, len(wb.worksheets)), min(summary_index, len(wb.worksheets))


def save_workbook_atomic(wb, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = tempfile.NamedTemporaryFile(prefix=f".{output_path.stem}_", suffix=".tmp.xlsx", dir=output_path.parent, delete=False)
    tmp_name = Path(tmp.name)
    tmp.close()
    try:
        wb.save(tmp_name)
        os.replace(tmp_name, output_path)
    finally:
        try:
            tmp_name.unlink(missing_ok=True)
        except Exception:
            pass

def main():
    parser = argparse.ArgumentParser(description="現_打鍵前 / 現_打鍵後 CSV をテンプレート Excel 形式へ展開します。")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help="設定ファイル JSON。既定: 設定ファイル.json")
    parser.add_argument("--template", default=None, help="添付テンプレート Excel")
    parser.add_argument("--template-sheet", default=None, help="テンプレート本体のシート名。既定: テーブル")
    parser.add_argument("--before", default=None, help="更新前CSVフォルダ")
    parser.add_argument("--after", default=None, help="更新後CSVフォルダ")
    parser.add_argument("--output", default=None, help="出力Excel")
    parser.add_argument("--sheet-name", default=None, help="出力シート名")
    parser.add_argument("--sql-name-mode", choices=["stem", "filename"], default=None, help="SQL の FROM に使う名前")
    parser.add_argument("--tables-per-sheet", type=int, default=None, help="1シートに出力するテーブル数。既定値は5。")
    args = parser.parse_args()

    settings = load_settings_file(Path(args.config) if args.config else None)
    common_cfg = settings.get("共通", {}) if isinstance(settings, dict) else {}
    old_cfg = settings.get("現CSV出力", {}) if isinstance(settings, dict) else {}

    # 各インプット/アウトプットは個別に指定できる。
    # 後方互換のため、未設定時は従来の「入力パス」「テンプレートパス」「出力パス」も読む。
    before_base = cfg_get(
        old_cfg, common_cfg,
        keys=["現_打鍵前入力パス", "現_打鍵前CSV入力パス", "入力パス", "入力フォルダパス", "基準パス"],
        default=".",
    )
    after_base = cfg_get(
        old_cfg, common_cfg,
        keys=["現_打鍵後入力パス", "現_打鍵後CSV入力パス", "入力パス", "入力フォルダパス", "基準パス"],
        default=".",
    )
    template_base = cfg_get(
        old_cfg, common_cfg,
        keys=["テンプレート入力パス", "テンプレートパス", "テンプレートフォルダパス"],
        default=".",
    )
    output_base = cfg_get(old_cfg, common_cfg, keys=["出力パス", "出力フォルダパス"], default=".")

    template_name = args.template or cfg_get(old_cfg, keys=["テンプレートファイル名", "テンプレート", "template"], default="テーブル.xlsx")
    before_name = args.before or cfg_get(old_cfg, keys=["現_打鍵前フォルダ名", "before", "before_folder"], default="現_打鍵前")
    after_name = args.after or cfg_get(old_cfg, keys=["現_打鍵後フォルダ名", "after", "after_folder"], default="現_打鍵後")
    output_name = args.output or cfg_get(old_cfg, keys=["出力ファイル名", "出力ファイル", "output"], default="比較結果.xlsx")

    template_path = join_if_relative(template_base, template_name)
    before_dir = join_if_relative(before_base, before_name)
    after_dir = join_if_relative(after_base, after_name)
    output_path = resolve_wildcard_path(output_base, output_name, must_exist=any(ch in str(output_name) for ch in "*?["))

    template_sheet = args.template_sheet or cfg_get(old_cfg, keys=["テンプレートシート名", "template_sheet"], default="テーブル")
    sheet_name_arg = args.sheet_name or cfg_get(old_cfg, common_cfg, keys=["テーブルシート名", "sheet_name"], default="テーブル")
    sql_name_mode = args.sql_name_mode or cfg_get(old_cfg, keys=["SQL名モード", "sql_name_mode"], default="stem")
    tables_per_sheet = args.tables_per_sheet or int(cfg_get(old_cfg, common_cfg, keys=["1シートテーブル数", "tables_per_sheet"], default=5))
    if tables_per_sheet <= 0:
        raise ValueError("--tables-per-sheet は 1 以上を指定してください。")

    existing_action = normalize_existing_action(cfg_get(old_cfg, common_cfg, keys=["既存ファイル処理", "既存時動作", "existing_action"], default=ACTION_REPLACE_FILE))
    auto_width_enabled = bool_value(cfg_get(old_cfg, common_cfg, keys=["列幅自動調整", "自動列幅調整", "auto_fit_columns"], default=False))
    max_width = cfg_get(old_cfg, common_cfg, keys=["最大列幅", "max_column_width"], default=60)
    min_width = cfg_get(old_cfg, common_cfg, keys=["最小列幅", "min_column_width"], default=8)

    if not template_path.exists():
        raise FileNotFoundError(f"テンプレート Excel が存在しません: {template_path}")
    if output_path.exists() and existing_action == ACTION_ERROR:
        raise FileExistsError(f"出力ファイルが既に存在します: {output_path}")

    template_wb = load_workbook(template_path)
    compare_exclusions = read_compare_exclusions(template_wb)
    if template_sheet in template_wb.sheetnames:
        tws = template_wb[template_sheet]
    else:
        template_candidates = [ws for ws in template_wb.worksheets if ws.title != COMPARE_EXCLUDE_SHEET_NAME]
        tws = template_candidates[0] if template_candidates else template_wb.active
    layout = template_layout(tws)

    records = build_table_records(csv_files(before_dir), csv_files(after_dir), sql_name_mode)
    if not records:
        raise RuntimeError("CSVファイルが見つかりません。")
    output_records = assign_output_positions(records, sheet_name_arg, tables_per_sheet)

    replace_sheets_mode = output_path.exists() and existing_action == ACTION_REPLACE_SHEETS
    if replace_sheets_mode:
        wb = load_workbook(output_path)
        table_insert_index, summary_insert_index = remove_existing_output_sheets(wb, sheet_name_arg)
        # テーブル検証シートがテーブルシートより前に戻される場合、
        # 後でテーブル検証を挿入するとテーブル位置が右へずれるため、先に1つ左へ補正する。
        adjusted_table_insert_index = max(0, table_insert_index - 1) if summary_insert_index <= table_insert_index else table_insert_index
    else:
        wb = template_wb
        table_insert_index = None
        adjusted_table_insert_index = None
        summary_insert_index = 0

    original_sheet_title = tws.title
    table_sheet_targets = {}
    table_sheets_to_rename = []

    sheet_names = sorted(set(r.output_sheet for r in output_records), key=lambda s: (len(s), s))
    for sheet_name in sheet_names:
        tmp_name = unique_sheet_name(wb, f"__{sheet_name}__")
        create_index = None if adjusted_table_insert_index is None else adjusted_table_insert_index + len(table_sheets_to_rename)
        ws = wb.create_sheet(tmp_name, create_index)
        copy_sheet_page_settings(tws, ws)
        sheet_records = [r for r in output_records if r.output_sheet == sheet_name]
        row = 1
        max_col_used = TEMPLATE_COLS
        header_rows = set()
        cf_targets = []
        for i, rec in enumerate(sheet_records):
            row, block_last_col = write_block(tws, ws, row, rec, layout, header_rows, cf_targets, compare_exclusions)
            max_col_used = max(max_col_used, block_last_col)
            if i != len(sheet_records) - 1:
                row += 1
        last_row = max(1, row - 1)
        normalize_fonts_and_row_heights(ws, last_row, max_col_used, header_rows)
        table_sheets_to_rename.append((ws, sheet_name))
        table_sheet_targets[sheet_name] = cf_targets

    write_summary_sheet(wb, records, summary_insert_index)

    if not replace_sheets_mode:
        remove_sheet_titles = {original_sheet_title, COMPARE_EXCLUDE_SHEET_NAME}
        for ws in list(wb.worksheets):
            if ws.title in remove_sheet_titles:
                wb.remove(ws)
    else:
        if "__一時退避__" in wb.sheetnames and len(wb.worksheets) > 1:
            del wb["__一時退避__"]

    for ws, sheet_name in table_sheets_to_rename:
        ws.title = sheet_name

    auto_fit_column_widths(
        wb,
        enabled=auto_width_enabled,
        max_width=max_width,
        min_width=min_width,
        sheet_names=[SUMMARY_SHEET_NAME] + [s for s in table_sheet_targets.keys()],
    )

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    save_workbook_atomic(wb, output_path)
    inject_x14_conditional_formatting(output_path, template_path, table_sheet_targets, layout, tws.title)
    print(f"出力しました: {output_path} / summary={SUMMARY_SHEET_NAME} / mode={existing_action}")


if __name__ == "__main__":
    main()
