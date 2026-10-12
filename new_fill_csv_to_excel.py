#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
比較結果.xlsx に、新_打鍵前 / 新_打鍵後 フォルダの CSV 内容を反映するツール。

想定する既存 Excel:
- generate_table_excel.py で作成した「比較結果.xlsx」
- テーブル明細シート名: テーブル / テーブル1 / テーブル2 ...
- テーブル見出し: A列に「1.テーブル名」の形式
- 各テーブル内の B列「新」エリアへ新システム CSV を転記する
- テーブル検証シートがある場合は更新する。summary シートの場合はテーブル検証へ改名/再作成する

注意:
- 既存の明細行数を前提に転記するため、CSV行数が明細シート上の「新」欄より多い場合、
  明細表示は収まる分だけ転記し、件数はヘルパーシート経由で CSV の実件数を計算式参照する。
- CSV入力フォルダは複数指定でき、同名CSVは先に指定した有効CSVを優先する。
- 優先フォルダ側のCSVが存在し「無効な識別」を含まない場合は、ヘッダーのみでも採用する。
- 「無効な識別」を含む場合のみ、後続フォルダの同名CSVを採用する。
- 既存 Excel 内に残っている x14 条件付き書式は保存前に退避し、保存後に再注入する。
"""

import argparse
import csv
import json
import io
import logging
import os
import re
import tempfile
import warnings
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path, PurePosixPath

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


DATA_START_COL = 3  # C列
FONT_NAME = "Yu Gothic"
BODY_ROW_HEIGHT = 18.75
HEADER_ROW_HEIGHT = 25.5
INVALID_MARKER = "無効な識別"
VALID_DELIMITERS = [",", "\t", ";", "|"]

SUMMARY_SHEET_NAME = "テーブル検証"
OLD_SUMMARY_SHEET_NAME = "summary"
HELPER_SHEET_NAME = "_新CSV読込"
COMPARE_EXCLUDE_SHEET_NAME = "現新比較対象外"
TRIANGLE_MARK = "△"

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
class Section:
    start: int
    end: int
    label: str

    @property
    def height(self) -> int:
        return max(0, self.end - self.start + 1)


@dataclass
class TableBlock:
    sheet_name: str
    title_row: int
    end_row: int
    output_no: int | None
    table_name: str
    before_current: Section
    before_new: Section
    after_current: Section
    after_new: Section
    header_before_row: int
    header_after_row: int
    before_count_row: int
    after_count_row: int
    count_diff_row: int
    data_last_col: int
    before_compare_start: int
    before_compare_end: int
    after_compare_start: int
    after_compare_end: int
    cross_ranges: list[tuple[int, int]] = field(default_factory=list)
    note: str = ""
    new_before_count: int = 0
    new_after_count: int = 0


@dataclass
class FillResult:
    table_name: str
    before_info: CsvInfo
    after_info: CsvInfo
    before_written: int = 0
    after_written: int = 0
    notes: list[str] = field(default_factory=list)


# ------------------------- CSV 読み込み -------------------------

def _normalize_csv_rows(rows):
    return [["" if v is None else str(v).strip() for v in row] for row in rows]


def _infer_delimiter(text: str) -> str:
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


def _read_csv_from_text(text: str):
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


def is_effective_csv_info(info: CsvInfo) -> bool:
    """優先採用できるCSVか判定する。

    条件:
    - CSVが存在する
    - 「無効な識別」を含まない

    ヘッダーのみでも採用し、無効な識別を含む場合だけ後続候補へフォールバックする。
    """
    return bool(info.exists and not info.invalid)


def csv_files(folder) -> dict[str, Path]:
    folders = folder if isinstance(folder, (list, tuple)) else [folder]
    result: dict[str, Path] = {}
    first_bad: dict[str, Path] = {}

    def register_bad(key: str, path: Path):
        first_bad.setdefault(key, path)

    for f in folders:
        f = Path(f)
        if not f.exists():
            raise FileNotFoundError(f"フォルダが存在しません: {f}")
        for p in f.iterdir():
            if not (p.is_file() and p.suffix.lower() == ".csv"):
                continue

            keys = [p.stem.lower(), p.name.lower()]
            # stem/name のどちらかで有効CSVが採用済みなら、このファイルは低優先候補なのでスキップ。
            if all(key in result for key in keys):
                continue

            info = read_csv_info(p)
            if is_effective_csv_info(info):
                for key in keys:
                    result.setdefault(key, p)
            else:
                # 無効な識別などの場合は、後続フォルダの同名CSVを探す。
                # ただし全候補が無効だった場合に備え、最初の無効候補は保持する。
                for key in keys:
                    register_bad(key, p)

    for key, p in first_bad.items():
        result.setdefault(key, p)
    return result


def find_csv(csv_map: dict[str, Path], table_name: str) -> Path | None:
    key = table_name.strip().lower()
    if key in csv_map:
        return csv_map[key]
    if f"{key}.csv" in csv_map:
        return csv_map[f"{key}.csv"]
    return None


# ------------------------- Excel 共通 -------------------------

def _copy_style_component(cell, attr_name, default_value):
    try:
        value = getattr(cell, attr_name)
        return copy(value) if value is not None else copy(default_value)
    except (IndexError, KeyError, TypeError, ValueError, AttributeError):
        return copy(default_value)


def force_font_name(cell):
    if isinstance(cell, MergedCell):
        return
    f = _copy_style_component(cell, "font", Font())
    f.name = FONT_NAME
    cell.font = f

    a = _copy_style_component(cell, "alignment", Alignment())
    a.vertical = "center"
    cell.alignment = a


def quote_sheet_name(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _normalize_compare_key(value):
    text = "" if value is None else str(value)
    return text.strip().lower()


def _normalize_table_key(value):
    text = _normalize_compare_key(value)
    if text.endswith(".csv"):
        text = text[:-4]
    return text


def _normalize_mark_text(value):
    text = "" if value is None else str(value).strip()
    return "CIRCLE" if text in {"\u25cb", "\u3007"} else text


def is_selected_mark(value, expected):
    return _normalize_mark_text(value) == _normalize_mark_text(expected)


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


def read_compare_exclusions_from_path(path: Path | None):
    if path is None or not path.exists():
        return set()
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            return read_compare_exclusions(wb)
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        return set()


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


def build_result_formula(row: int, first_col: int, last_col: int) -> str:
    return f'=IF(COUNTIF({get_column_letter(first_col)}{row}:{get_column_letter(last_col)}{row},FALSE)>0,"差分あり","差分なし")'


def build_compare_formula(row_a: int, row_b: int, col: int, table_aliases=None, item_name=None, compare_exclusions=None, blank_skip=True) -> str:
    letter = get_column_letter(col)
    exact_expr = f"EXACT({letter}{row_a},{letter}{row_b})"
    if is_compare_excluded(table_aliases or set(), item_name, compare_exclusions or set()):
        body = f'IF({exact_expr},TRUE,"{TRIANGLE_MARK}")'
    else:
        body = exact_expr
    if blank_skip:
        return f'=IF(AND({letter}{row_a}="",{letter}{row_b}=""),"",{body})'
    return f"={body}"


def rewrite_vertical_compare_rows(ws, start_row: int, end_row: int, row_a_start: int, row_b_start: int, header_row: int, data_last_col: int, table_aliases: set, compare_exclusions: set):
    if start_row <= 0 or end_row < start_row:
        return
    for offset, r in enumerate(range(start_row, end_row + 1)):
        ws.cell(r, 2).value = build_result_formula(r, DATA_START_COL, data_last_col)
        force_font_name(ws.cell(r, 2))
        for c in range(DATA_START_COL, data_last_col + 1):
            item_name = ws.cell(header_row, c).value
            cell = ws.cell(r, c)
            cell.value = build_compare_formula(row_a_start + offset, row_b_start + offset, c, table_aliases, item_name, compare_exclusions, blank_skip=True)
            force_font_name(cell)


def rewrite_existing_exact_formula(cell, table_aliases: set, item_name, compare_exclusions: set):
    value = cell.value
    if not (isinstance(value, str) and value.startswith("=")):
        return
    refs = re.findall(r"\$?[A-Z]{1,3}\$?\d+", value)
    if len(refs) < 2:
        return
    exact_expr = f"EXACT({refs[0]},{refs[1]})"
    if is_compare_excluded(table_aliases, item_name, compare_exclusions):
        cell.value = f'=IF({exact_expr},TRUE,"{TRIANGLE_MARK}")'
    else:
        cell.value = f"={exact_expr}"
    force_font_name(cell)


def rewrite_cross_compare_rows(ws, row_start: int, row_end: int, header_row: int, data_last_col: int, table_aliases: set, compare_exclusions: set):
    if row_start <= 0 or row_end < row_start:
        return
    for r in range(row_start, row_end + 1):
        for c in range(DATA_START_COL, data_last_col + 1):
            item_name = ws.cell(header_row, c).value
            rewrite_existing_exact_formula(ws.cell(r, c), table_aliases, item_name, compare_exclusions)


def rewrite_compare_formulas_for_block(ws, block: TableBlock, compare_exclusions: set):
    table_aliases = make_table_aliases(block.table_name)
    rewrite_vertical_compare_rows(
        ws,
        block.before_compare_start,
        block.before_compare_end,
        block.before_current.start,
        block.before_new.start,
        block.header_before_row,
        block.data_last_col,
        table_aliases,
        compare_exclusions,
    )
    rewrite_vertical_compare_rows(
        ws,
        block.after_compare_start,
        block.after_compare_end,
        block.after_current.start,
        block.after_new.start,
        block.header_after_row,
        block.data_last_col,
        table_aliases,
        compare_exclusions,
    )
    for s, e in block.cross_ranges:
        rewrite_cross_compare_rows(ws, s, e, block.header_before_row, block.data_last_col, table_aliases, compare_exclusions)


def copy_style_if_needed(ws, row: int, src_col: int, dst_col: int):
    if dst_col <= src_col:
        return
    src = ws.cell(row, src_col)
    dst = ws.cell(row, dst_col)
    # _style の直接コピーは既存ブックの style index 不整合を引き継ぐことがあるため避ける。
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
    force_font_name(dst)


def clear_cell_value_and_keep_style(cell):
    if isinstance(cell, MergedCell):
        return
    cell.value = None
    force_font_name(cell)


def apply_gray_to_row(ws, row: int, col_start: int, col_end: int):
    for c in range(col_start, col_end + 1):
        ws.cell(row, c).fill = copy(GRAY_FILL)


def apply_thin_grid(ws, row_start: int, row_end: int, col_start: int, col_end: int):
    if row_end < row_start or col_end < col_start:
        return
    border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            ws.cell(r, c).border = copy(border)


def is_table_sheet_name(name: str, prefix: str) -> bool:
    return name == prefix or re.fullmatch(re.escape(prefix) + r"\d+", name) is not None


def title_value_parts(value):
    if value is None:
        return None
    text = str(value).strip()
    m = re.match(r"^\s*(\d+)\.(.+?)\s*$", text)
    if not m:
        return None
    return int(m.group(1)), m.group(2).strip()


def header_last_col(ws, header_row: int, min_last_col: int) -> int:
    last = min_last_col
    for c in range(DATA_START_COL, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v not in (None, ""):
            last = c
    return max(last, DATA_START_COL)


def row_has_any_value(ws, row: int, col_start: int, col_end: int) -> bool:
    for c in range(col_start, col_end + 1):
        v = ws.cell(row, c).value
        if v not in (None, ""):
            return True
    return False


def get_b_merged_rows(ws) -> set[int]:
    rows = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= 2 <= rng.max_col:
            for r in range(rng.min_row, rng.max_row + 1):
                rows.add(r)
    return rows


def get_label_sections(ws, start_row: int, end_row: int, label: str) -> list[Section]:
    sections: list[Section] = []
    covered = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= 2 <= rng.max_col and start_row <= rng.min_row <= end_row:
            value = ws.cell(rng.min_row, rng.min_col).value
            if str(value).strip() == label:
                s = max(start_row, rng.min_row)
                e = min(end_row, rng.max_row)
                sections.append(Section(s, e, label))
                covered.update(range(s, e + 1))

    for r in range(start_row, end_row + 1):
        if r in covered:
            continue
        value = ws.cell(r, 2).value
        if str(value).strip() == label:
            sections.append(Section(r, r, label))

    sections.sort(key=lambda x: x.start)
    return sections


def discover_table_blocks(wb, table_sheet_prefix: str) -> list[TableBlock]:
    blocks: list[TableBlock] = []

    for ws in wb.worksheets:
        if ws.title in (SUMMARY_SHEET_NAME, OLD_SUMMARY_SHEET_NAME, HELPER_SHEET_NAME):
            continue
        if not is_table_sheet_name(ws.title, table_sheet_prefix):
            continue

        title_rows = []
        for r in range(1, ws.max_row + 1):
            parts = title_value_parts(ws.cell(r, 1).value)
            if parts:
                title_rows.append((r, parts[0], parts[1]))

        for i, (title_row, output_no, table_name) in enumerate(title_rows):
            end_row = title_rows[i + 1][0] - 1 if i + 1 < len(title_rows) else ws.max_row
            current_sections = get_label_sections(ws, title_row, end_row, "現")
            new_sections = get_label_sections(ws, title_row, end_row, "新")

            if len(current_sections) < 2 or len(new_sections) < 2:
                # 明細ブロックではない/想定外の場合はスキップ
                continue

            before_current = current_sections[0]
            after_current = current_sections[1]
            before_new = new_sections[0]
            after_new = new_sections[1]

            before_cmp_start = before_new.end + 1
            before_cmp_end = before_cmp_start + before_new.height - 1
            after_cmp_start = after_new.end + 1
            after_cmp_end = after_cmp_start + after_new.height - 1
            before_count_row = after_cmp_end + 2
            after_count_row = after_cmp_end + 3
            count_diff_row = after_cmp_end + 4

            data_last_col = max(
                header_last_col(ws, before_current.start - 1, DATA_START_COL),
                header_last_col(ws, after_current.start - 1, DATA_START_COL),
            )

            cross_ranges = []
            # 3つ目以降の「現」「新」は打鍵前後データ対比エリア
            for sec in current_sections[2:] + new_sections[2:]:
                cross_ranges.append((sec.start, sec.end))

            blocks.append(
                TableBlock(
                    sheet_name=ws.title,
                    title_row=title_row,
                    end_row=end_row,
                    output_no=output_no,
                    table_name=table_name,
                    before_current=before_current,
                    before_new=before_new,
                    after_current=after_current,
                    after_new=after_new,
                    header_before_row=before_current.start - 1,
                    header_after_row=after_current.start - 1,
                    before_count_row=before_count_row,
                    after_count_row=after_count_row,
                    count_diff_row=count_diff_row,
                    data_last_col=data_last_col,
                    before_compare_start=before_cmp_start,
                    before_compare_end=before_cmp_end,
                    after_compare_start=after_cmp_start,
                    after_compare_end=after_cmp_end,
                    cross_ranges=cross_ranges,
                )
            )

    blocks.sort(key=lambda b: (b.sheet_name, b.title_row))
    return blocks


# ------------------------- x14 条件付き書式の退避/復元 -------------------------

def sheet_xml_paths(xlsx_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(xlsx_path, "r") as z:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        sheets_el = wb_root.find(f"{ns_main}sheets")
        sheets = list(sheets_el) if sheets_el is not None else []

        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rid_to_target = {}
        for rel in rels_root.findall(f"{rel_ns}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = str(PurePosixPath("xl") / target)
            rid_to_target[rid] = path

        result = {}
        for sheet in sheets:
            name = sheet.attrib.get("name")
            rid = sheet.attrib.get(f"{ns_rel}id")
            if name and rid in rid_to_target:
                result[name] = rid_to_target[rid]
        return result


def extract_extlst_by_sheet(xlsx_path: Path) -> dict[str, str]:
    result = {}
    try:
        paths = sheet_xml_paths(xlsx_path)
        with zipfile.ZipFile(xlsx_path, "r") as z:
            for sheet_name, xml_path in paths.items():
                try:
                    xml = z.read(xml_path).decode("utf-8")
                except KeyError:
                    continue
                m = re.search(r"<extLst>[\s\S]*?</extLst>", xml)
                if m and "x14:conditionalFormatting" in m.group(0):
                    result[sheet_name] = m.group(0)
    except Exception:
        return {}
    return result


def reinject_extlst_by_sheet(xlsx_path: Path, ext_map: dict[str, str]):
    if not ext_map:
        return

    paths = sheet_xml_paths(xlsx_path)
    path_to_ext = {paths[name]: ext for name, ext in ext_map.items() if name in paths}
    if not path_to_ext:
        return

    tmp_dir = xlsx_path.parent if xlsx_path.parent.exists() else Path(".")
    tmp = tempfile.NamedTemporaryFile(prefix=f".{xlsx_path.stem}_", suffix=".tmp.xlsx", dir=tmp_dir, delete=False)
    tmp_name = tmp.name
    tmp.close()

    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in path_to_ext:
                    xml = data.decode("utf-8")
                    if "x14:conditionalFormatting" not in xml:
                        ext = path_to_ext[info.filename]
                        if "</extLst>" in xml:
                            # 既存 extLst がある場合は中身だけ追加
                            inner = ext.replace("<extLst>", "", 1).rsplit("</extLst>", 1)[0]
                            xml = xml.replace("</extLst>", inner + "</extLst>", 1)
                        else:
                            xml = xml.replace("</worksheet>", ext + "</worksheet>", 1)
                        data = xml.encode("utf-8")
                zout.writestr(info, data)
        os.replace(tmp_name, xlsx_path)
    finally:
        try:
            Path(tmp_name).unlink(missing_ok=True)
        except Exception:
            pass


# ------------------------- 新 CSV 反映 -------------------------

def max_csv_cols(info: CsvInfo) -> int:
    m = len(info.header)
    for row in info.rows:
        m = max(m, len(row))
    return m


def fill_new_section(ws, section: Section, info: CsvInfo, data_last_col: int) -> tuple[int, list[str]]:
    notes = []
    display_rows = section.height
    actual_rows = info.rows
    write_count = min(len(actual_rows), display_rows)

    if info.invalid:
        notes.append("CSVに無効な識別あり。0件扱い")
    if not info.exists:
        notes.append("CSVなし。0件扱い")
    if info.exists and not info.invalid and info.count == 0:
        notes.append("CSVがヘッダーのみ。0件扱い")
    if len(actual_rows) > display_rows:
        notes.append(f"CSV件数{len(actual_rows)}件のうち、明細欄に収まる{display_rows}件のみ転記")

    csv_col_count = max(1, max_csv_cols(info))
    last_col = max(data_last_col, DATA_START_COL + csv_col_count - 1)

    # 必要に応じて右側列にも左隣のスタイルをコピーする
    if last_col > data_last_col:
        for r in range(section.start, section.end + 1):
            for c in range(data_last_col + 1, last_col + 1):
                copy_style_if_needed(ws, r, data_last_col, c)

    for offset, row_no in enumerate(range(section.start, section.end + 1)):
        source = actual_rows[offset] if offset < write_count else []
        is_blank_line = offset >= write_count
        for c in range(DATA_START_COL, last_col + 1):
            idx = c - DATA_START_COL
            cell = ws.cell(row_no, c)
            cell.value = source[idx] if idx < len(source) else None
            force_font_name(cell)
        if is_blank_line:
            apply_gray_to_row(ws, row_no, DATA_START_COL, last_col)

    apply_thin_grid(ws, section.start, section.end, DATA_START_COL, last_col)
    return write_count, notes


def add_standard_compare_cf(ws, start_row: int, end_row: int, data_last_col: int):
    if start_row <= 0 or end_row < start_row or data_last_col < DATA_START_COL:
        return
    ref = f"{get_column_letter(DATA_START_COL)}{start_row}:{get_column_letter(data_last_col)}{end_row}"
    ws.conditional_formatting.add(
        ref,
        CellIsRule(operator="equal", formula=["TRUE"], fill=TRUE_FILL),
    )
    ws.conditional_formatting.add(
        ref,
        CellIsRule(operator="equal", formula=["FALSE"], fill=FALSE_FILL, font=FALSE_FONT),
    )
    ws.conditional_formatting.add(
        ref,
        CellIsRule(operator="equal", formula=[f'"{TRIANGLE_MARK}"'], fill=WARNING_FILL),
    )


def normalize_rows_and_fonts(ws, max_col: int):
    for r in range(1, ws.max_row + 1):
        # 既存行高を尊重しつつ、未設定なら標準化
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = BODY_ROW_HEIGHT
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                force_font_name(cell)


def trim_text_cells(ws, max_col: int):
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            value = cell.value
            if isinstance(value, str) and not value.startswith("="):
                cell.value = value.strip()


# ------------------------- テーブル検証 -------------------------

def collect_existing_summary_info(wb) -> dict[str, dict[str, str]]:
    ws = None
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        ws = wb[SUMMARY_SHEET_NAME]
    elif OLD_SUMMARY_SHEET_NAME in wb.sheetnames:
        ws = wb[OLD_SUMMARY_SHEET_NAME]
    if ws is None or ws.max_row < 2:
        return {}

    headers = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value is not None}
    table_col = headers.get("テーブル名")
    note_col = headers.get("備考")
    current_before_csv_col = headers.get("現_打鍵前CSVパス") or headers.get("現_打鍵前CSV出处")
    current_after_csv_col = headers.get("現_打鍵後CSVパス") or headers.get("現_打鍵後CSV出处")
    if not table_col:
        return {}

    result = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, table_col).value
        if not name:
            continue
        note = ws.cell(r, note_col).value if note_col else ""
        current_before_csv = ws.cell(r, current_before_csv_col).value if current_before_csv_col else ""
        current_after_csv = ws.cell(r, current_after_csv_col).value if current_after_csv_col else ""
        result[str(name).strip()] = {
            "note": str(note or ""),
            "current_before_csv": str(current_before_csv or ""),
            "current_after_csv": str(current_after_csv or ""),
        }
    return result


def write_helper_sheet(wb, helper_rows: list[dict], selected_table_names=None):
    preserved_rows = []
    if HELPER_SHEET_NAME in wb.sheetnames:
        if selected_table_names is not None:
            for values in wb[HELPER_SHEET_NAME].iter_rows(min_row=2, max_col=6, values_only=True):
                if values[0] and not (make_table_aliases(values[0]) & selected_table_names):
                    preserved_rows.append(values)
        del wb[HELPER_SHEET_NAME]
    ws = wb.create_sheet(HELPER_SHEET_NAME)
    headers = [
        "テーブル名",
        "新_打鍵前件数",
        "新_打鍵後件数",
        "新_打鍵前CSV",
        "新_打鍵後CSV",
        "備考",
    ]
    ws.append(headers)
    for row in helper_rows:
        ws.append([
            row.get("table_name", ""),
            row.get("new_before_count", 0),
            row.get("new_after_count", 0),
            row.get("new_before_csv", ""),
            row.get("new_after_csv", ""),
            row.get("note", ""),
        ])
    for values in preserved_rows:
        ws.append(values)

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = HEADER_ROW_HEIGHT if r == 1 else BODY_ROW_HEIGHT
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            force_font_name(cell)
            cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
            if r == 1:
                cell.fill = copy(SUMMARY_HEADER_FILL)
                cell.font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
    widths = [32, 16, 16, 35, 35, 70]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.sheet_state = "hidden"
    return ws


def _build_validation_data(wb, blocks: list[TableBlock], existing_summary_info: dict[str, dict[str, str]], helper_names: set[str]):
    ws = wb.create_sheet("_validation_update")
    headers = [
        "CRUD",
        "出力番号",
        "出力Sheet",
        "テーブル名",
        "現_打鍵前件数",
        "現_打鍵後件数",
        "現_打鍵前後件数差異",
        "新_打鍵前件数",
        "新_打鍵後件数",
        "新_打鍵前後件数差異",
        "現_打鍵前CSVパス",
        "現_打鍵後CSVパス",
        "新_打鍵前CSVパス",
        "新_打鍵後CSVパス",
        "現新件数チェック",
        "差分チェック",
        "備考",
    ]
    ws.append(headers)

    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.fill = copy(SUMMARY_HEADER_FILL)
        cell.font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    row_no = 2
    seen = set()
    for block in sorted(blocks, key=lambda b: (b.output_no or 10**9, b.sheet_name, b.title_row)):
        seen.add(block.table_name.lower())
        sheet_ref = quote_sheet_name(block.sheet_name)
        helper_ref = quote_sheet_name(HELPER_SHEET_NAME)

        note_parts = []
        existing_info = existing_summary_info.get(block.table_name, {})
        old_note = existing_info.get("note", "")
        if old_note:
            note_parts.extend(old_note.split(" / "))
        if block.note:
            note_parts.extend(block.note.split(" / "))
        note = " / ".join(dict.fromkeys(n for n in note_parts if n))

        ws.cell(row_no, 1).value = f'=IF(AND(E{row_no}=0,F{row_no}=0),"",IF(F{row_no}>E{row_no},"C",IF(F{row_no}<E{row_no},"D","U")))'
        ws.cell(row_no, 2).value = block.output_no if block.output_no is not None else ""
        ws.cell(row_no, 3).value = block.sheet_name
        ws.cell(row_no, 4).value = block.table_name
        ws.cell(row_no, 5).value = f"={sheet_ref}!C{block.before_count_row}"
        ws.cell(row_no, 6).value = f"={sheet_ref}!C{block.after_count_row}"
        ws.cell(row_no, 7).value = f"=F{row_no}-E{row_no}"
        ws.cell(row_no, 8).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,2,FALSE),0)'
        ws.cell(row_no, 9).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,3,FALSE),0)'
        ws.cell(row_no, 10).value = f"=I{row_no}-H{row_no}"
        ws.cell(row_no, 11).value = existing_info.get("current_before_csv", "")
        ws.cell(row_no, 12).value = existing_info.get("current_after_csv", "")
        ws.cell(row_no, 13).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,4,FALSE),"")'
        ws.cell(row_no, 14).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,5,FALSE),"")'
        ws.cell(row_no, 15).value = f'=IF(AND(E{row_no}=H{row_no},F{row_no}=I{row_no}),"差分なし","差分あり")'
        ws.cell(row_no, 16).value = f'=IF(COUNTIF({sheet_ref}!B{block.title_row}:B{block.end_row},"差分あり")>0,"NG","OK")'
        ws.cell(row_no, 17).value = note

        current_before = _safe_numeric_formula_source(wb[block.sheet_name].cell(block.before_count_row, 3).value)
        current_after = _safe_numeric_formula_source(wb[block.sheet_name].cell(block.after_count_row, 3).value)
        total_new = block.new_before_count + block.new_after_count
        if current_before == 0 and current_after == 0 and total_new == 0:
            ws.row_dimensions[row_no].hidden = True
        row_no += 1

    # 明細シートに存在しないが、新CSVだけ存在するテーブルも検証シートへ表示する
    for table_name in sorted(helper_names - seen):
        helper_ref = quote_sheet_name(HELPER_SHEET_NAME)
        note = "明細シートなし。新CSVのみ存在"
        ws.cell(row_no, 1).value = f'=IF(AND(E{row_no}=0,F{row_no}=0),"",IF(F{row_no}>E{row_no},"C",IF(F{row_no}<E{row_no},"D","U")))'
        ws.cell(row_no, 2).value = ""
        ws.cell(row_no, 3).value = ""
        ws.cell(row_no, 4).value = table_name
        ws.cell(row_no, 5).value = "=0"
        ws.cell(row_no, 6).value = "=0"
        ws.cell(row_no, 7).value = f"=F{row_no}-E{row_no}"
        ws.cell(row_no, 8).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,2,FALSE),0)'
        ws.cell(row_no, 9).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,3,FALSE),0)'
        ws.cell(row_no, 10).value = f"=I{row_no}-H{row_no}"
        ws.cell(row_no, 11).value = ""
        ws.cell(row_no, 12).value = ""
        ws.cell(row_no, 13).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,4,FALSE),"")'
        ws.cell(row_no, 14).value = f'=IFERROR(VLOOKUP(D{row_no},{helper_ref}!$A:$F,5,FALSE),"")'
        ws.cell(row_no, 15).value = f'=IF(AND(E{row_no}=H{row_no},F{row_no}=I{row_no}),"差分なし","差分あり")'
        ws.cell(row_no, 16).value = ""
        ws.cell(row_no, 17).value = note
        row_no += 1

    max_row = max(1, ws.max_row)
    max_col = len(headers)
    for r in range(1, max_row + 1):
        ws.row_dimensions[r].height = HEADER_ROW_HEIGHT if r == 1 else BODY_ROW_HEIGHT
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            force_font_name(cell)
            cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
            if r > 1:
                cell.alignment = Alignment(vertical="center", wrap_text=(c in {11, 12, 13, 14, 17}))

    # 行背景は数式結果が未計算でも見やすいように出力時点の現件数で軽く分類する
    for r in range(2, max_row + 1):
        crud_guess = ws.cell(r, 1).value
        # 数式なので表示色は基本グレー。Excelでフィルタ確認しやすくする
        fill = SUMMARY_SKIP_FILL
        for c in range(1, max_col + 1):
            ws.cell(r, c).fill = copy(fill)

    if max_row >= 2:
        count_check_ref = f"O2:O{max_row}"
        data_check_ref = f"P2:P{max_row}"
        ws.conditional_formatting.add(count_check_ref, CellIsRule(operator="equal", formula=['"差分なし"'], fill=TRUE_FILL))
        ws.conditional_formatting.add(count_check_ref, CellIsRule(operator="equal", formula=['"差分あり"'], fill=FALSE_FILL, font=FALSE_FONT))
        ws.conditional_formatting.add(data_check_ref, CellIsRule(operator="equal", formula=['"OK"'], fill=TRUE_FILL))
        ws.conditional_formatting.add(data_check_ref, CellIsRule(operator="equal", formula=['"NG"'], fill=FALSE_FILL, font=FALSE_FONT))

    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    ws.freeze_panes = "A2"
    widths = [10, 12, 16, 32, 18, 18, 22, 18, 18, 22, 45, 45, 45, 45, 18, 18, 75]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    return ws


def write_validation_sheet(wb, blocks: list[TableBlock], existing_summary_info: dict[str, dict[str, str]], helper_names: set[str]):
    ws = next((wb[name] for name in (SUMMARY_SHEET_NAME, OLD_SUMMARY_SHEET_NAME)
               if name in wb.sheetnames), None)
    data = _build_validation_data(wb, blocks, existing_summary_info, helper_names)
    if ws is None:
        data.title = SUMMARY_SHEET_NAME
        wb.move_sheet(data, offset=-wb.index(data))
        return data

    if ws.cell(1, 13).value != "新_打鍵前CSVパス":
        old_max_col = ws.max_column
        # openpyxl inserts cells but does not move dimensions or formatting ranges.
        dimensions = {c: copy(ws.column_dimensions[get_column_letter(c)])
                      for c in range(13, old_max_col + 1)}
        merges = [copy(area) for area in ws.merged_cells.ranges]
        for area in merges:
            ws.unmerge_cells(str(area))
        rules = [(copy(cf), list(cf_rules))
                 for cf, cf_rules in ws.conditional_formatting._cf_rules.items()]
        ws.conditional_formatting._cf_rules.clear()
        ws.insert_cols(13, 2)
        for c, dimension in dimensions.items():
            letter = get_column_letter(c + 2)
            dimension.index = letter
            dimension.min = dimension.max = c + 2
            ws.column_dimensions[letter] = dimension
        for area in merges:
            if area.min_col >= 13:
                area.shift(col_shift=2)
            elif area.max_col >= 13:
                area.max_col += 2
            ws.merge_cells(str(area))
        for cf, cf_rules in rules:
            for area in cf.sqref.ranges:
                if area.min_col >= 13:
                    area.shift(col_shift=2)
                elif area.max_col >= 13:
                    area.max_col += 2
            for rule in cf_rules:
                ws.conditional_formatting.add(str(cf.sqref), rule)
        for target, source in ((13, 11), (14, 12)):
            letter = get_column_letter(target)
            dimension = copy(ws.column_dimensions[get_column_letter(source)])
            dimension.index = letter
            dimension.min = dimension.max = target
            ws.column_dimensions[letter] = dimension
            for r in range(1, ws.max_row + 1):
                ws.cell(r, target)._style = copy(ws.cell(r, source)._style)
        if ws.auto_filter.ref:
            from openpyxl.worksheet.cell_range import CellRange
            area = CellRange(ws.auto_filter.ref)
            area.max_col += 2
            ws.auto_filter.ref = str(area)
            for column in ws.auto_filter.filterColumn:
                if column.colId >= 12:
                    column.colId += 2

    for c in (11, 12, 13, 14):
        ws.cell(1, c).value = data.cell(1, c).value
    existing_rows = {str(ws.cell(r, 4).value).strip().lower(): r
                     for r in range(2, ws.max_row + 1) if ws.cell(r, 4).value}
    for source_row in range(2, data.max_row + 1):
        key = str(data.cell(source_row, 4).value).strip().lower()
        target_row = existing_rows.get(key)
        if target_row is None:
            target_row = ws.max_row + 1
            for c in range(1, 18):
                ws.cell(target_row, c)._style = copy(ws.cell(target_row - 1, c)._style)
            ws.row_dimensions[target_row].height = ws.row_dimensions[target_row - 1].height
            columns = range(1, 18)
        else:
            columns = (8, 9, 10, 13, 14, 15, 17)
        for c in columns:
            value = data.cell(source_row, c).value
            if isinstance(value, str) and value.startswith("="):
                from openpyxl.formula import Tokenizer
                from openpyxl.formula.translate import Translator
                tokens = Tokenizer(value).items
                for token in tokens:
                    if token.subtype == "RANGE" and "!" not in token.value:
                        token.value = Translator("=" + token.value,
                                                 origin=data.cell(source_row, c).coordinate).translate_formula(
                                                     ws.cell(target_row, c).coordinate)[1:]
                value = "=" + "".join(token.value for token in tokens)
            ws.cell(target_row, c).value = value
    if ws.auto_filter.ref:
        from openpyxl.worksheet.cell_range import CellRange
        area = CellRange(ws.auto_filter.ref)
        area.max_row = max(area.max_row, ws.max_row)
        ws.auto_filter.ref = str(area)
    wb.remove(data)
    return ws


def _safe_numeric_formula_source(value):
    if isinstance(value, (int, float)):
        return value
    return None


# ------------------------- メイン処理 -------------------------

def make_helper_rows(blocks: list[TableBlock], before_map: dict[str, Path], after_map: dict[str, Path], fill_results: dict[str, FillResult], selected_table_names=None):
    rows = []
    names = {b.table_name for b in blocks}
    # 新CSVだけ存在するものも追加
    for key, path in before_map.items():
        if key.endswith(".csv"):
            continue
        names.add(path.stem)
    for key, path in after_map.items():
        if key.endswith(".csv"):
            continue
        names.add(path.stem)

    for name in sorted(names, key=lambda x: x.lower()):
        if selected_table_names is not None and not (make_table_aliases(name) & selected_table_names):
            continue
        res = fill_results.get(name.lower())
        before_info = res.before_info if res else read_csv_info(find_csv(before_map, name))
        after_info = res.after_info if res else read_csv_info(find_csv(after_map, name))
        notes = []
        if not before_info.exists:
            notes.append("新_打鍵前CSVなし。0件扱い")
        if not after_info.exists:
            notes.append("新_打鍵後CSVなし。0件扱い")
        if before_info.invalid:
            notes.append("新_打鍵前CSVに無効な識別あり。0件扱い")
        if after_info.invalid:
            notes.append("新_打鍵後CSVに無効な識別あり。0件扱い")
        if before_info.exists and not before_info.invalid and before_info.count == 0:
            notes.append("新_打鍵前CSVがヘッダーのみ。0件扱い")
        if after_info.exists and not after_info.invalid and after_info.count == 0:
            notes.append("新_打鍵後CSVがヘッダーのみ。0件扱い")
        if res:
            notes.extend(res.notes)
        rows.append({
            "table_name": name,
            "new_before_count": before_info.count,
            "new_after_count": after_info.count,
            "new_before_csv": str(before_info.path) if before_info.path else "",
            "new_after_csv": str(after_info.path) if after_info.path else "",
            "note": " / ".join(dict.fromkeys([n for n in notes if n])),
        })
    return rows



# ------------------------- 設定ファイル / 出力制御 -------------------------
DEFAULT_CONFIG_FILE = "設定ファイル.json"
ACTION_REPLACE_FILE = "replace_file"
ACTION_ERROR = "error"
ACTION_REPLACE_SHEETS = "replace_sheets"


def setup_logging(script_name: str):
    log_dir = Path.cwd() / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"{script_name}_{timestamp}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8")],
        force=True,
    )
    return logging.getLogger(script_name), log_path


def print_and_log(message: str, logger=None):
    print(message)
    if logger:
        logger.info(message)


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



def _as_ordered_values(value):
    """設定値を優先順のリストへ変換する。

    JSON 配列のほか、文字列では ; / ； / 改行区切りも利用できる。
    Windows の C:/... を壊さないため、コロンでは分割しない。
    """
    if value is None:
        return ["."]
    if isinstance(value, (list, tuple)):
        result = []
        for item in value:
            result.extend(_as_ordered_values(item))
        return result or ["."]
    text = str(value).strip()
    if text == "":
        return ["."]
    parts = re.split(r"[;；\r\n]+", text)
    parts = [p.strip() for p in parts if p and p.strip()]
    return parts or [text]


def _first_ordered_value(value, default=None):
    values = _as_ordered_values(value)
    if not values:
        return default
    return values[0]


def has_any_key(*dicts, keys):
    for d in dicts:
        if not isinstance(d, dict):
            continue
        for key in keys:
            if key in d and d[key] not in (None, ""):
                return True
    return False


def _dedupe_paths(paths):
    result = []
    seen = set()
    for p in paths:
        p = Path(str(normalize_path_value(p)))
        try:
            key = str(p.resolve()) if p.exists() else str(p.absolute())
        except Exception:
            key = str(p)
        if key in seen:
            continue
        seen.add(key)
        result.append(p)
    return result


def resolve_path_list(path_value):
    """設定ファイル上のCSVフォルダパスを優先順リストとして解決する。

    例:
      "現_打鍵前CSVパス": ["優先1/現_打鍵前", "優先2/現_打鍵前"]
      => 優先1/現_打鍵前, 優先2/現_打鍵前 の順で検索する。
    """
    return _dedupe_paths(_as_ordered_values(path_value))


def resolve_folder_list(base_value, folder_value):
    """旧設定互換: 入力パス + フォルダ名を優先順リストとして解決する。"""
    bases = _as_ordered_values(base_value)
    folders = _as_ordered_values(folder_value)
    return _dedupe_paths(join_if_relative(base, folder) for base in bases for folder in folders)


def resolve_file_path(path_value, must_exist=False):
    """設定ファイル上のファイルパスを解決する。ワイルドカードは1件一致のみ許可する。"""
    raw = _first_ordered_value(path_value, default=".")
    file_path = Path(str(normalize_path_value(raw)))
    pattern = str(file_path)
    if any(ch in pattern for ch in "*?["):
        parent = file_path.parent if str(file_path.parent) else Path(".")
        matches = sorted(parent.glob(file_path.name))
        matches = [m for m in matches if m.is_file()]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            raise FileNotFoundError(f"ワイルドカードに一致するファイルがありません: {pattern}")
        raise RuntimeError("ワイルドカードに一致するファイルが複数あります: " + ", ".join(str(m) for m in matches))
    if must_exist and not file_path.exists():
        raise FileNotFoundError(f"ファイルが存在しません: {file_path}")
    return file_path

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

def column_index_from_config(value, label):
    """設定値の列指定を 1 始まりの列番号へ変換する。例: K / k / K列 / 11。"""
    if value is None or str(value).strip() == "":
        raise ValueError(f"{label} が設定されていません。")
    text = str(value).strip()
    if text.endswith("列"):
        text = text[:-1].strip()
    if text.isdigit():
        idx = int(text)
        if idx <= 0:
            raise ValueError(f"{label} は 1 以上を指定してください: {value}")
        return idx
    text = text.upper()
    if not re.fullmatch(r"[A-Z]+", text):
        raise ValueError(f"{label} の列指定が不正です: {value}")
    idx = 0
    for ch in text:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def read_selected_table_names_from_workbook(path: Path, sheet_name, mark_col, table_col, mark_value="〇"):
    """出力済みExcelの指定シートから、〇が付いた表名を読む。

    指定シートが存在しない場合は None を返し、呼び出し側はフィルタなしとして扱う。
    """
    if not sheet_name or path is None or not path.exists():
        return None
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        mark_col_idx = column_index_from_config(mark_col, "対象表選択判定列")
        table_col_idx = column_index_from_config(table_col, "対象表選択表名列")
        expected = "" if mark_value is None else str(mark_value).strip()
        selected = set()
        for r in range(1, ws.max_row + 1):
            mark = ws.cell(r, mark_col_idx).value
            table_name = ws.cell(r, table_col_idx).value
            if not is_selected_mark(mark, expected):
                continue
            key = _normalize_table_key(table_name)
            if key:
                selected.add(key)
        return selected
    finally:
        try:
            wb.close()
        except Exception:
            pass


def read_selected_table_conditions_from_workbook(path: Path, sheet_name, mark_col, table_col, before_current_col, before_new_col, after_current_col, after_new_col, mark_value="〇"):
    """出力済みExcelの指定シートから、〇が付いた表名ごとの4種類の抽出条件を読む。"""
    condition_cols = [before_current_col, before_new_col, after_current_col, after_new_col]
    if not sheet_name or not any(condition_cols) or path is None or not path.exists():
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    formula_wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in wb.sheetnames:
            return {}
        ws = wb[sheet_name]
        formula_ws = formula_wb[sheet_name] if sheet_name in formula_wb.sheetnames else None
        mark_col_idx = column_index_from_config(mark_col, "対象表選択判定列")
        table_col_idx = column_index_from_config(table_col, "対象表選択表名列")
        before_current_idx = column_index_from_config(before_current_col, "対象表選択打鍵前現条件列") if before_current_col else None
        before_new_idx = column_index_from_config(before_new_col, "対象表選択打鍵前新条件列") if before_new_col else None
        after_current_idx = column_index_from_config(after_current_col, "対象表選択打鍵後現条件列") if after_current_col else None
        after_new_idx = column_index_from_config(after_new_col, "対象表選択打鍵後新条件列") if after_new_col else None
        expected = "" if mark_value is None else str(mark_value).strip()
        conditions = {}

        def read_condition(row, col_idx):
            if not col_idx:
                return ""
            value = ws.cell(row, col_idx).value
            if value in (None, "") and formula_ws is not None:
                formula_value = formula_ws.cell(row, col_idx).value
                if isinstance(formula_value, str) and formula_value.startswith("="):
                    # openpyxl は数式を計算できないため、キャッシュが空の数式は条件なしとして扱う。
                    return ""
                value = formula_value
            return "" if value is None else str(value).strip()

        for r in range(1, ws.max_row + 1):
            mark = ws.cell(r, mark_col_idx).value
            table_name = ws.cell(r, table_col_idx).value
            if not is_selected_mark(mark, expected):
                continue
            key = _normalize_table_key(table_name)
            if not key:
                continue
            conditions[key] = {
                "before_current": read_condition(r, before_current_idx),
                "before_new": read_condition(r, before_new_idx),
                "after_current": read_condition(r, after_current_idx),
                "after_new": read_condition(r, after_new_idx),
            }
        return conditions
    finally:
        try:
            wb.close()
        except Exception:
            pass
        try:
            formula_wb.close()
        except Exception:
            pass


def block_matches_selected_tables(block: TableBlock, selected_table_names):
    if selected_table_names is None:
        return True
    aliases = make_table_aliases(block.table_name)
    return any(alias in selected_table_names for alias in aliases)


def table_aliases_with_csv_data(before_map: dict[str, Path], after_map: dict[str, Path]):
    aliases = set()
    for key in sorted(set(before_map.keys()) | set(after_map.keys())):
        before_info = read_csv_info(before_map.get(key))
        after_info = read_csv_info(after_map.get(key))
        if before_info.count <= 0 and after_info.count <= 0:
            continue
        display_path = before_info.path or after_info.path
        display_name = display_path.stem if display_path else key[:-4] if key.lower().endswith(".csv") else key
        aliases.update(make_table_aliases(display_name, key))
    return aliases


def find_or_create_selector_result_col(ws, header_name="CSVデータ不一致", header_row=3):
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(header_row, c).value or "").strip() == header_name:
            return c
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value or "").strip() == header_name:
            ws.cell(1, c).value = None
            ws.cell(header_row, c).value = header_name
            ws.cell(header_row, c).font = Font(name=FONT_NAME, bold=True)
            ws.cell(header_row, c).alignment = Alignment(horizontal="center", vertical="center")
            return c
    col = ws.max_column + 1
    ws.cell(header_row, col).value = header_name
    ws.cell(header_row, col).font = Font(name=FONT_NAME, bold=True)
    ws.cell(header_row, col).alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[get_column_letter(col)].width = max(ws.column_dimensions[get_column_letter(col)].width or 0, 28)
    return col


def mark_selector_csv_mismatches(wb, sheet_name, mark_col, table_col, mark_value, data_table_aliases, selected_table_names=None):
    if not sheet_name or sheet_name not in wb.sheetnames:
        return 0
    ws = wb[sheet_name]
    mark_col_idx = column_index_from_config(mark_col, "対象表選択判定列")
    table_col_idx = column_index_from_config(table_col, "対象表選択表名列")
    result_col_idx = find_or_create_selector_result_col(ws, header_row=3)
    expected = "" if mark_value is None else str(mark_value).strip()
    mismatch_count = 0
    for r in range(4, ws.max_row + 1):
        table_name = ws.cell(r, table_col_idx).value
        table_key = _normalize_table_key(table_name)
        if not table_key:
            ws.cell(r, result_col_idx).value = None
            continue
        if selected_table_names is not None:
            selected = table_key in selected_table_names
        else:
            selected = is_selected_mark(ws.cell(r, mark_col_idx).value, expected)
        has_csv_data = table_key in data_table_aliases
        result = ""
        if selected and not has_csv_data:
            result = "選択あり/CSVデータなし"
        elif not selected and has_csv_data:
            result = "CSVデータあり/選択なし"
        ws.cell(r, result_col_idx).value = result
        force_font_name(ws.cell(r, result_col_idx))
        if result:
            mismatch_count += 1
    return mismatch_count


def sql_text_for_condition(table_name, condition):
    condition = "" if condition is None else str(condition).strip()
    base = f"SELECT * FROM {table_name}"
    return f"{base} WHERE {condition}" if condition else base


def find_sql_rows_in_range(ws, start_row, end_row):
    result = {}
    for r in range(max(1, start_row), min(ws.max_row, end_row) + 1):
        label = str(ws.cell(r, 2).value or "").strip()
        sql_text = str(ws.cell(r, 3).value or "").strip().upper()
        if "取得SQL" not in label and not sql_text.startswith("SELECT"):
            continue
        if "新" in label:
            result.setdefault("new", r)
        elif "現" in label:
            result.setdefault("current", r)
        else:
            result.setdefault("current", r)
    return result


def find_sql_rows_near_section(ws, title_row, section_start):
    result = find_sql_rows_in_range(ws, title_row, section_start - 1)
    if result:
        return result
    return find_sql_rows_in_range(ws, max(1, section_start - 8), section_start - 1)


def conditions_for_block(block: TableBlock, selected_table_conditions):
    aliases = make_table_aliases(block.table_name)
    for alias in aliases:
        if alias in selected_table_conditions:
            return selected_table_conditions[alias]
    return {}


def write_extraction_conditions_for_block(ws, block: TableBlock, conditions):
    if not conditions:
        return 0
    before_rows = find_sql_rows_near_section(ws, block.title_row, block.before_current.start)
    after_rows = find_sql_rows_near_section(ws, block.before_compare_end + 1, block.after_current.start)
    targets = [
        (before_rows.get("new"), conditions.get("before_new", "")),
        (after_rows.get("new"), conditions.get("after_new", "")),
    ]
    count = 0
    for row_no, condition in targets:
        if not row_no:
            continue
        if condition in (None, "") or str(condition).strip() == "":
            continue
        ws.cell(row_no, 3).value = sql_text_for_condition(block.table_name, condition)
        force_font_name(ws.cell(row_no, 3))
        count += 1
    return count


def write_extraction_conditions_for_blocks(wb, blocks, selected_table_conditions):
    if not selected_table_conditions:
        return 0
    count = 0
    for block in blocks:
        conditions = conditions_for_block(block, selected_table_conditions)
        if not conditions:
            continue
        count += write_extraction_conditions_for_block(wb[block.sheet_name], block, conditions)
    return count

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

def fill_workbook(workbook_path: Path, before_dir, after_dir, output_path: Path, table_sheet_prefix: str, template_path: Path | None = None, auto_width=False, max_width=60, min_width=8, selected_table_names=None, selected_table_conditions=None, selector_info=None):
    if not workbook_path.exists():
        raise FileNotFoundError(f"Excelファイルが存在しません: {workbook_path}")

    before_map = csv_files(before_dir)
    after_map = csv_files(after_dir)
    data_table_aliases = table_aliases_with_csv_data(before_map, after_map)

    preserved_ext = extract_extlst_by_sheet(workbook_path)
    compare_exclusions = read_compare_exclusions_from_path(template_path)

    wb = load_workbook(workbook_path)
    if not compare_exclusions:
        compare_exclusions = read_compare_exclusions(wb)
    existing_summary_info = collect_existing_summary_info(wb)
    blocks = discover_table_blocks(wb, table_sheet_prefix)
    if not blocks:
        raise RuntimeError("テーブル明細ブロックが見つかりません。先に old_fill_csv_to_excel.py を実行し、シート名・タイトル行を確認してください。")
    if selected_table_names is not None:
        blocks = [block for block in blocks if block_matches_selected_tables(block, selected_table_names)]

    condition_write_count = write_extraction_conditions_for_blocks(wb, blocks, selected_table_conditions or {})

    fill_results: dict[str, FillResult] = {}

    for block in blocks:
        ws = wb[block.sheet_name]
        before_info = read_csv_info(find_csv(before_map, block.table_name))
        after_info = read_csv_info(find_csv(after_map, block.table_name))

        # CSV列数が既存より多い場合は data_last_col を拡張して転記する
        csv_last_col = DATA_START_COL + max(max_csv_cols(before_info), max_csv_cols(after_info), 1) - 1
        block.data_last_col = max(block.data_last_col, csv_last_col)

        before_written, before_notes = fill_new_section(ws, block.before_new, before_info, block.data_last_col)
        after_written, after_notes = fill_new_section(ws, block.after_new, after_info, block.data_last_col)

        block.new_before_count = before_info.count
        block.new_after_count = after_info.count
        notes = []
        for n in before_notes:
            notes.append("新_打鍵前" + n)
        for n in after_notes:
            notes.append("新_打鍵後" + n)
        block.note = " / ".join(dict.fromkeys([n for n in notes if n]))

        fill_results[block.table_name.lower()] = FillResult(
            table_name=block.table_name,
            before_info=before_info,
            after_info=after_info,
            before_written=before_written,
            after_written=after_written,
            notes=notes,
        )

        rewrite_compare_formulas_for_block(ws, block, compare_exclusions)

        # 通常の TRUE/FALSE/△ 条件付き書式を補う。
        add_standard_compare_cf(ws, block.before_compare_start, block.before_compare_end, block.data_last_col)
        add_standard_compare_cf(ws, block.after_compare_start, block.after_compare_end, block.data_last_col)
        for s, e in block.cross_ranges:
            add_standard_compare_cf(ws, s, e, block.data_last_col)
        trim_text_cells(ws, max(block.data_last_col, ws.max_column))
        normalize_rows_and_fonts(ws, max(block.data_last_col, ws.max_column))

    helper_rows = make_helper_rows(blocks, before_map, after_map, fill_results, selected_table_names)
    write_helper_sheet(wb, helper_rows, selected_table_names)
    helper_names = {row["table_name"].lower() for row in helper_rows}
    write_validation_sheet(wb, blocks, existing_summary_info, helper_names)

    auto_fit_column_widths(wb, enabled=auto_width, max_width=max_width, min_width=min_width, sheet_names=[ws.title for ws in wb.worksheets if (is_table_sheet_name(ws.title, table_sheet_prefix) or ws.title == HELPER_SHEET_NAME) and ws.title not in {SUMMARY_SHEET_NAME, OLD_SUMMARY_SHEET_NAME}])

    mismatch_count = None
    if selector_info:
        mismatch_count = mark_selector_csv_mismatches(
            wb,
            selector_info.get("sheet_name"),
            selector_info.get("mark_col"),
            selector_info.get("table_col"),
            selector_info.get("mark_value"),
            data_table_aliases,
            selected_table_names,
        )

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

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

    # x14 条件付き書式を可能な限り復元する
    reinject_extlst_by_sheet(output_path, preserved_ext)

    return blocks, helper_rows, mismatch_count, condition_write_count


def main():
    parser = argparse.ArgumentParser(description="新_打鍵前 / 新_打鍵後 CSV を比較結果.xlsx の『新』欄へ反映します。")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help="設定ファイル JSON。既定: 設定ファイル.json")
    parser.add_argument("--workbook", default=None, help="更新対象の比較結果 Excel。既定: 比較結果.xlsx")
    parser.add_argument("--new-before", default=None, help="新システム 打鍵前 CSV フォルダ。既定: 新_打鍵前")
    parser.add_argument("--new-after", default=None, help="新システム 打鍵後 CSV フォルダ。既定: 新_打鍵後")
    parser.add_argument("--output", default=None, help="出力先 Excel。省略時は --workbook を上書きします。")
    parser.add_argument("--table-sheet-prefix", default=None, help="明細シート名の接頭辞。既定: テーブル")
    parser.add_argument("--template", default=None, help="比較対象外項目定義を読むテンプレート Excel。既定: テーブル.xlsx")
    args = parser.parse_args()
    logger, log_path = setup_logging(Path(__file__).stem)
    logger.info("処理開始")
    logger.info("ログファイル: %s", log_path)
    logger.info("設定ファイル: %s", args.config)

    settings = load_settings_file(Path(args.config) if args.config else None)
    common_cfg = settings.get("共通", {}) if isinstance(settings, dict) else {}
    new_cfg = settings.get("新CSV反映", {}) if isinstance(settings, dict) else {}

    # 新形式: パス + フォルダ名/ファイル名を分けず、1項目に完全パスを書く。
    # 旧形式も互換のため残す。新形式が設定されている場合は新形式を優先する。
    before_path_value = args.new_before or cfg_get(
        new_cfg, common_cfg,
        keys=["新_打鍵前CSVパス", "新_打鍵前CSVフォルダ", "新_打鍵前パス", "new_before_csv_path", "new_before_path"],
        default=None,
    )
    after_path_value = args.new_after or cfg_get(
        new_cfg, common_cfg,
        keys=["新_打鍵後CSVパス", "新_打鍵後CSVフォルダ", "新_打鍵後パス", "new_after_csv_path", "new_after_path"],
        default=None,
    )
    workbook_path_value = args.workbook or cfg_get(
        new_cfg, common_cfg,
        keys=["対象Excelファイルパス", "対象Excelパス", "入力Excelファイルパス", "workbook_path"],
        default=None,
    )
    if workbook_path_value is None and not has_any_key(new_cfg, common_cfg, keys=["対象Excel入力パス", "対象Excelフォルダパス"]):
        workbook_path_value = cfg_get(new_cfg, keys=["対象Excelファイル", "入力Excelファイル", "workbook"], default=None)

    template_path_value = args.template or cfg_get(
        new_cfg, common_cfg,
        keys=["テンプレートファイルパス", "テンプレートExcelパス", "template_path"],
        default=None,
    )
    if template_path_value is None and not has_any_key(new_cfg, common_cfg, keys=["テンプレート入力パス", "テンプレートフォルダパス"]):
        template_path_value = cfg_get(new_cfg, keys=["テンプレートファイル", "テンプレート", "template"], default=None)

    output_path_value = args.output or cfg_get(
        new_cfg, common_cfg,
        keys=["出力ファイルパス", "出力Excelパス", "output_path"],
        default=None,
    )
    if output_path_value is None and not has_any_key(new_cfg, common_cfg, keys=["出力パス", "出力フォルダパス"]):
        output_path_value = cfg_get(new_cfg, keys=["出力ファイル", "output"], default=None)

    if before_path_value is not None:
        before_dir = resolve_path_list(before_path_value)
    else:
        before_base = cfg_get(new_cfg, common_cfg, keys=["新_打鍵前入力パス", "新_打鍵前CSV入力パス", "入力パス", "入力フォルダパス", "基準パス"], default=".")
        before_name = cfg_get(new_cfg, keys=["新_打鍵前フォルダ名", "new_before", "new_before_folder"], default="新_打鍵前")
        before_dir = resolve_folder_list(before_base, before_name)

    if after_path_value is not None:
        after_dir = resolve_path_list(after_path_value)
    else:
        after_base = cfg_get(new_cfg, common_cfg, keys=["新_打鍵後入力パス", "新_打鍵後CSV入力パス", "入力パス", "入力フォルダパス", "基準パス"], default=".")
        after_name = cfg_get(new_cfg, keys=["新_打鍵後フォルダ名", "new_after", "new_after_folder"], default="新_打鍵後")
        after_dir = resolve_folder_list(after_base, after_name)

    if workbook_path_value is not None:
        workbook_path = resolve_file_path(workbook_path_value, must_exist=True)
    else:
        workbook_base = cfg_get(new_cfg, common_cfg, keys=["対象Excel入力パス", "対象Excelパス", "対象Excelフォルダパス"], default=cfg_get(new_cfg, common_cfg, keys=["出力パス", "出力フォルダパス"], default="."))
        workbook_name = cfg_get(new_cfg, keys=["対象Excelファイル名", "入力Excelファイル名", "workbook"], default="比較結果.xlsx")
        workbook_path = resolve_wildcard_path(workbook_base, workbook_name, must_exist=True)

    if output_path_value is not None:
        output_path = resolve_file_path(output_path_value, must_exist=any(ch in str(_first_ordered_value(output_path_value, "")) for ch in "*?["))
    else:
        output_base = cfg_get(new_cfg, common_cfg, keys=["出力パス", "出力フォルダパス"], default=workbook_path.parent)
        output_name = cfg_get(new_cfg, keys=["出力ファイル名", "出力ファイル", "output"], default=None)
        if output_name:
            output_path = resolve_wildcard_path(output_base, output_name, must_exist=any(ch in str(output_name) for ch in "*?["))
        else:
            output_path = workbook_path

    if template_path_value is not None:
        template_path = resolve_file_path(template_path_value, must_exist=False)
    else:
        template_base = cfg_get(new_cfg, common_cfg, keys=["テンプレート入力パス", "テンプレートパス", "テンプレートフォルダパス"], default=".")
        template_name = cfg_get(new_cfg, keys=["テンプレートファイル名", "テンプレート", "template"], default="テーブル.xlsx")
        template_path = join_if_relative(template_base, template_name)

    table_sheet_prefix = args.table_sheet_prefix or cfg_get(new_cfg, common_cfg, keys=["テーブルシート名", "table_sheet_prefix"], default="テーブル")

    existing_action = normalize_existing_action(cfg_get(new_cfg, common_cfg, keys=["既存ファイル処理", "既存時動作", "existing_action"], default=ACTION_REPLACE_FILE))
    if output_path.exists() and output_path.resolve() != workbook_path.resolve() and existing_action == ACTION_ERROR:
        raise FileExistsError(f"出力ファイルが既に存在します: {output_path}")
    # 新CSV反映では replace_sheets は「既存ブック内の検証/ヘルパーシートを再作成して保存」として扱う。
    # 明細テーブルシート自体は削除せず、新欄だけ更新する。

    auto_width_enabled = bool_value(cfg_get(new_cfg, common_cfg, keys=["列幅自動調整", "自動列幅調整", "auto_fit_columns"], default=False))
    max_width = cfg_get(new_cfg, common_cfg, keys=["最大列幅", "max_column_width"], default=60)
    min_width = cfg_get(new_cfg, common_cfg, keys=["最小列幅", "min_column_width"], default=8)
    selector_enabled = bool_value(cfg_get(new_cfg, common_cfg, keys=["対象表選択有効", "対象表選択を使用", "enable_table_selector"], default=False))
    selector_sheet = cfg_get(new_cfg, common_cfg, keys=["対象表選択シート名", "選択シート名", "table_selector_sheet"], default=None)
    selector_mark_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択判定列", "選択判定列", "table_selector_mark_column"], default=None)
    selector_table_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択表名列", "選択表名列", "table_selector_table_column"], default=None)
    selector_legacy_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択条件列", "選択条件列", "table_selector_condition_column"], default=None)
    selector_before_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵前条件列", "打鍵前条件列", "table_selector_before_condition_column"], default=selector_legacy_condition_col)
    selector_after_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵後条件列", "打鍵後条件列", "table_selector_after_condition_column"], default=selector_legacy_condition_col)
    selector_before_current_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵前現条件列", "打鍵前現条件列", "table_selector_before_current_condition_column"], default=selector_before_condition_col)
    selector_before_new_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵前新条件列", "打鍵前新条件列", "table_selector_before_new_condition_column"], default=selector_before_condition_col)
    selector_after_current_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵後現条件列", "打鍵後現条件列", "table_selector_after_current_condition_column"], default=selector_after_condition_col)
    selector_after_new_condition_col = cfg_get(new_cfg, common_cfg, keys=["対象表選択打鍵後新条件列", "打鍵後新条件列", "table_selector_after_new_condition_column"], default=selector_after_condition_col)
    selector_mark_value = cfg_get(new_cfg, common_cfg, keys=["対象表選択判定値", "選択判定値", "table_selector_mark_value"], default="〇")
    logger.info("新_打鍵前CSVパス: %s", [str(p) for p in before_dir])
    logger.info("新_打鍵後CSVパス: %s", [str(p) for p in after_dir])
    logger.info("対象Excelファイル: %s", workbook_path)
    logger.info("テンプレートファイル: %s", template_path)
    logger.info("出力ファイル: %s", output_path)
    logger.info("既存ファイル処理: %s", existing_action)
    logger.info("対象表選択有効: %s", selector_enabled)
    selected_table_names = None
    selected_table_conditions = {}
    if selector_enabled:
        selected_table_names = read_selected_table_names_from_workbook(
            output_path,
            selector_sheet,
            selector_mark_col,
            selector_table_col,
            selector_mark_value,
        )
        selected_table_conditions = read_selected_table_conditions_from_workbook(
            output_path,
            selector_sheet,
            selector_mark_col,
            selector_table_col,
            None,
            selector_before_new_condition_col,
            None,
            selector_after_new_condition_col,
            selector_mark_value,
        )
        logger.info(
            "対象表選択シート: %s / 判定列: %s / 表名列: %s / 打鍵前現条件列: %s / 打鍵前新条件列: %s / 打鍵後現条件列: %s / 打鍵後新条件列: %s / 判定値: %s",
            selector_sheet,
            selector_mark_col,
            selector_table_col,
            selector_before_current_condition_col,
            selector_before_new_condition_col,
            selector_after_current_condition_col,
            selector_after_new_condition_col,
            selector_mark_value,
        )
        non_empty_condition_count = sum(
            1
            for conditions in selected_table_conditions.values()
            for key, value in conditions.items()
            if key in {"before_new", "after_new"} and str(value or "").strip()
        )
        logger.info("対象表選択件数: %s / 新抽出条件対象表数: %s / 非空新抽出条件数: %s", 0 if selected_table_names is None else len(selected_table_names), len(selected_table_conditions), non_empty_condition_count)

    selector_info = None
    if selector_enabled:
        selector_info = {
            "sheet_name": selector_sheet,
            "mark_col": selector_mark_col,
            "table_col": selector_table_col,
            "mark_value": selector_mark_value,
        }

    blocks, helper_rows, mismatch_count, condition_write_count = fill_workbook(
        workbook_path=workbook_path,
        before_dir=before_dir,
        after_dir=after_dir,
        output_path=output_path,
        table_sheet_prefix=table_sheet_prefix,
        template_path=template_path,
        auto_width=auto_width_enabled,
        max_width=max_width,
        min_width=min_width,
        selected_table_names=selected_table_names,
        selected_table_conditions=selected_table_conditions,
        selector_info=selector_info,
    )

    print_and_log(f"出力しました: {output_path} / mode={existing_action}", logger)
    print_and_log(f"明細更新テーブル数: {len(blocks)}", logger)
    print_and_log(f"新CSV検証件数: {len(helper_rows)}", logger)
    print_and_log(f"取得SQL条件更新セル数: {condition_write_count}", logger)
    if mismatch_count is not None:
        print_and_log(f"対象表選択CSV不一致件数: {mismatch_count}", logger)
    print_and_log(f"ログファイル: {log_path}", logger)
    if selected_table_names is not None:
        print_and_log(f"対象表選択: {len(selected_table_names)}件指定 / {len(blocks)}件更新", logger)
    logger.info("処理終了")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logging.exception("処理中にエラーが発生しました。")
        raise
