from __future__ import annotations

import codecs
import csv
import json
import os
import posixpath
import queue
import re
import sys
import tempfile
import threading
import time
import traceback
import zipfile
from contextlib import contextmanager
from copy import copy, deepcopy
from datetime import date, datetime, time as datetime_time, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path, PurePosixPath
from typing import Any, Callable
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
from xml.parsers import expat

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, MergedCell, WriteOnlyCell
from openpyxl.formula.translate import Translator
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.datetime import from_excel, WINDOWS_EPOCH
from wps_automation_workflow import (DEFAULT_WORKFLOW, MODULE_LABELS, add_workflow_module, delete_profile, list_profiles,
                                    load_profile, matching_profiles, move_workflow_module,
                                    remove_workflow_module, save_profile, validate_workflow)

APP_NAME = "WPS 表格多源汇总处理器"
CONFIG_PATH = Path.home() / ".wps_sheet_automation.json"
PROFILE_DIR = Path.home() / ".wps_sheet_automation_profiles"
DEFAULT_WINDOW_WIDTH = 1500
DEFAULT_WINDOW_HEIGHT = 960
MIN_WINDOW_WIDTH = 1024
MIN_WINDOW_HEIGHT = 680


def _window_size_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
        return base / "WpsSheetAutomation" / "window_size.json"
    xdg_config = os.environ.get("XDG_CONFIG_HOME", "")
    base = Path(xdg_config) if xdg_config and Path(xdg_config).is_absolute() else Path.home() / ".config"
    return base / "wps_sheet_automation" / "window_size.json"


WINDOW_SIZE_PATH = _window_size_path()
CONFIG_VERSION = 18
LEGACY_FORMULA_ROW_KEYS = ("formula_seed_row", "formula_template_row", "formula_row", "seed_row")
MAX_PROCESSING_CELLS = 500_000
MAX_REGION_XLSX_CELLS = 100_000
MAX_REGION_XLSX_CHARS = 2_000_000
MAX_TEXT_CELLS = 100_000
MAX_TEXT_CHARS = 2_000_000
TEXT_COPY_CHUNK_ROWS = 500
MAX_XLSX_MEMBERS = 20_000
MAX_XLSX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_XLSX_WORKSHEET_XML_BYTES = 64 * 1024 * 1024
MAX_XLSX_METADATA_XML_BYTES = 8 * 1024 * 1024
MAX_CSV_FILE_BYTES = 512 * 1024 * 1024
DEFAULT_SETTINGS: dict[str, Any] = {
    "config_version": CONFIG_VERSION,
    "config_name": "",
    "template_prefix_length": 5,
    "workflow": list(DEFAULT_WORKFLOW),
    "template_path": "",
    "output_dir": "",
    "output_filename": "处理结果.xlsx",
    "output_filename_custom": None,
    "merge_only": False,
    "merge_output_filename": "合并结果.xlsx",
    "merge_sources": [],
    "retain_merged_sources": False,
    "delete_processed_sources_after_success": False,
    # Deprecated compatibility alias; normalized to retain_merged_sources.
    "retain_edited_sources": False,
    "copy_start_col": "A",
    "copy_end_col": "",
    "pre_mappings": [],
    "post_mappings": [],
    "region_xlsx_outputs": [],
    "text_outputs": [],
    "CustomFormula": {"enabled": False, "rules": []},
    "TextToColumns": {"enabled": False, "rules": []},
    "Deduplicate": {"enabled": False, "rules": []},
    "FormatConvert": {"enabled": False, "rules": []},
    "RegionFilter": {"enabled": False, "rules": []},
    # Kept for migration of first-version configurations.
    "source_start_row": 2,
    "jobs": [
        {"template_sheet": "1当天派单", "merged_source_filename": "", "header_rows": 2, "start_row": 3, "formula_range": "EB:EN", "mappings": [],
         "sources": [{"path": "", "source_sheet": "", "copy_start_row": 2, "copy_end_row": None, "keep_first_row": True, "_optional_placeholder": True}]},
        {"template_sheet": "2当天归档", "merged_source_filename": "", "header_rows": 2, "start_row": 3, "formula_range": "EB:EU", "mappings": [],
         "sources": [{"path": "", "source_sheet": "", "copy_start_row": 2, "copy_end_row": None, "keep_first_row": True, "_optional_placeholder": True}]},
        {"template_sheet": "3装移全业务在途", "merged_source_filename": "", "header_rows": 3, "start_row": 4, "formula_range": "EB:FV", "mappings": [],
         "sources": [{"path": "", "source_sheet": "", "copy_start_row": 2, "copy_end_row": None, "keep_first_row": True, "_optional_placeholder": True}]},
    ],
}


class ProcessingError(RuntimeError):
    """An actionable error safe to show to the user."""


def window_size_limits(screen_width: Any = None, screen_height: Any = None) -> tuple[int, int, int, int]:
    """Return minimum and maximum window dimensions, safely bounded to the display."""
    def positive_dimension(value: Any, fallback: int) -> int:
        try:
            dimension = int(value)
        except (TypeError, ValueError, OverflowError):
            return fallback
        return dimension if dimension > 0 else fallback

    screen_width = positive_dimension(screen_width, DEFAULT_WINDOW_WIDTH + 64)
    screen_height = positive_dimension(screen_height, DEFAULT_WINDOW_HEIGHT + 96)
    max_width = screen_width - 48 if screen_width >= 400 else screen_width
    max_height = screen_height - 80 if screen_height >= 320 else screen_height
    max_width, max_height = max(1, max_width), max(1, max_height)
    return (min(MIN_WINDOW_WIDTH, max_width), min(MIN_WINDOW_HEIGHT, max_height), max_width, max_height)


def normalize_window_size(saved: Any, screen_width: Any = None,
                          screen_height: Any = None) -> tuple[int, int]:
    """Validate saved width/height and clamp them to the current display."""
    min_width, min_height, max_width, max_height = window_size_limits(screen_width, screen_height)

    def parse_dimension(value: Any, minimum: int) -> int | None:
        if isinstance(value, bool):
            return None
        if isinstance(value, int):
            dimension = value
        elif isinstance(value, str) and re.fullmatch(r"[0-9]{1,5}", value.strip()):
            dimension = int(value.strip())
        else:
            return None
        return dimension if minimum <= dimension <= 16_384 else None

    width = height = None
    if isinstance(saved, dict):
        width = parse_dimension(saved.get("width"), 320)
        height = parse_dimension(saved.get("height"), 240)
    if width is None or height is None:
        width, height = DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT
    return (min(max(width, min_width), max_width), min(max(height, min_height), max_height))


def load_window_size(path: Path | str | None = None, screen_width: Any = None,
                     screen_height: Any = None) -> tuple[int, int]:
    """Load only the prior window dimensions; malformed or missing state uses defaults."""
    state_path = Path(path) if path is not None else WINDOW_SIZE_PATH
    try:
        saved = json.loads(state_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        saved = None
    return normalize_window_size(saved, screen_width, screen_height)


def save_window_size(width: Any, height: Any, path: Path | str | None = None,
                     screen_width: Any = None, screen_height: Any = None) -> bool:
    """Persist only validated dimensions, atomically; return False on I/O failure."""
    state_path = Path(path) if path is not None else WINDOW_SIZE_PATH
    size = normalize_window_size({"width": width, "height": height}, screen_width, screen_height)
    temporary = state_path.with_name(state_path.name + ".tmp")
    try:
        state_path.parent.mkdir(parents=True, exist_ok=True)
        temporary.write_text(json.dumps({"width": size[0], "height": size[1]}), encoding="utf-8")
        os.replace(temporary, state_path)
        return True
    except OSError:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        return False


def column_number(value: str) -> int:
    value = str(value).strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", value):
        raise ProcessingError(f"列号“{value or '(空)'}”无效；请使用 A、DG、XFD 这样的 Excel 列名。")
    try:
        number = column_index_from_string(value)
    except ValueError as exc:
        raise ProcessingError(f"列号“{value}”超出 Excel 列范围。") from exc
    if number > 16384:
        raise ProcessingError(f"列号“{value}”超出 Excel 最大列 XFD。")
    return number


def parse_column_range(value: str) -> tuple[int, int]:
    parts = str(value).strip().upper().split(":")
    if len(parts) != 2:
        raise ProcessingError(f"列范围“{value}”格式错误；请按 A:C 或 EB:EN 填写。")
    start, end = (column_number(part) for part in parts)
    if start > end:
        raise ProcessingError(f"列范围“{value}”的起始列不能在末列之后。")
    return start, end


def ranges_overlap(first: tuple[int, int], second: tuple[int, int]) -> bool:
    return first[0] <= second[1] and second[0] <= first[1]


def _positive_row(value: Any, label: str, allow_zero: bool = False) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise ProcessingError(f"{label}必须是{'非负' if allow_zero else '正'}整数。") from exc
    lower = 0 if allow_zero else 1
    if result < lower or result > 1_048_576:
        raise ProcessingError(f"{label}必须介于 {lower} 和 1048576。")
    return result


def _formula_template_row(job: dict[str, Any]) -> int:
    """Formula templates always come from the row immediately above pasted data."""
    start_row = _positive_row(job.get("start_row"), "粘贴起始行")
    if start_row <= 1:
        raise ProcessingError("公式填充要求粘贴起始行至少为第 2 行；公式模板行自动取粘贴起始行的上一行。")
    return start_row - 1


def configuration_migration_notes(raw: Any) -> list[str]:
    """Explain legacy row semantics without trusting any saved seed-row field."""
    if not isinstance(raw, dict):
        return []
    try:
        version = int(raw.get("config_version", 1))
    except (TypeError, ValueError):
        version = 1
    jobs = raw.get("jobs", [])
    if not isinstance(jobs, list):
        jobs = []
    notes: list[str] = []
    if version < 7:
        notes.append(
            f"旧配置版本 v{version} 已按 v7 规则迁移：每个目标表的 start_row 保持为粘贴起始行，"
            "公式模板行自动取 start_row - 1；旧版本中将起始行当作公式种子行的含义不再沿用。"
        )
    if version < 11:
        notes.append("配置已升级到 v11：源文件规则先在临时副本中由 WPS 执行；目标表可不关联源文件并直接保留模板内容处理。可选保留的源副本会写入输出目录，原件不变。")
    if version < 12:
        notes.append("配置已升级到 v12：新增 WPS 公式填充规则，支持输入自定义公式或从指定单元格取公式，并按相对引用填充所选范围；规则可用于源文件临时副本或模板工作簿。")
    if version < 13:
        notes.append("配置已升级到 v13：结果映射已恢复为可选步骤，与后映射分别保留；旧配置中的结果列设置会迁移为结果映射规则。")
    if version < 14:
        notes.append("配置已升级到 v14：源文件先合并后处理；旧配置中的后续源文件保留首行以保持原有结果，可在源文件设置中调整。区域 XLSX 输出保留，图片输出选项已移除。")
    if version < 15:
        notes.append("配置已升级到 v15：新增独立源文件合并模式，可不选模板地将首份 CSV 或 XLSX 与后续 XLSX 合并为自定义命名的 XLSX 文件。")
    if version < 16:
        notes.append("配置已升级到 v16：模板模式可为每个目标表设置合并源工作簿文件名；留空时使用模板子表名。可选择保存合并源副本，不保存时会在模板处理完成后清理临时文件。旧保留选项将自动兼容。")
    if version < 17:
        notes.append("配置已升级到 v17：区域规则的工作表名可留空以使用第一张表，名称匹配不区分大小写；区域筛选范围可留空以筛选整张工作表。源文件规则先处理合并后的源副本，再将处理结果用于模板。")
    if version < 18:
        notes.append("配置已升级到 v18：新增“模板处理全部成功后删除已另存的源副本”选项；默认关闭并保留副本，启用后仅在模板处理成功时删除，失败时保留供排查。")
    if "open_region_exports" in raw:
        notes.append("旧版自动打开区域 XLSX 的选项已停用；区域 XLSX 将输出到所选目录，不再自动打开。")
    legacy_workflow = raw.get("workflow", [])
    if not isinstance(legacy_workflow, list):
        legacy_workflow = []
    for index, job in enumerate(jobs, 1):
        if not isinstance(job, dict):
            continue
        legacy_values = [(key, job[key]) for key in LEGACY_FORMULA_ROW_KEYS if key in job]
        if not legacy_values:
            continue
        try:
            template_row = _formula_template_row(job)
        except ProcessingError:
            template_row = "无法计算（粘贴起始行必须至少为 2）"
        old_fields = "、".join(f"{key}={value!r}" for key, value in legacy_values)
        notes.append(
            f"第 {index} 个目标表的旧公式行字段（{old_fields}）不会用于选取公式；"
            f"实际公式模板行按 start_row - 1 自动确定为 {template_row}。"
        )
    return notes


def _validate_sheet_name(value: Any, label: str, allow_blank: bool = False) -> str:
    name = str(value or "").strip()
    if not name:
        if allow_blank:
            return ""
        raise ProcessingError(f"{label}工作表名不能为空。")
    if len(name) > 31 or any(char in name for char in "[]:*?/\\"):
        raise ProcessingError(f"{label}工作表名“{name}”不符合 Excel 命名规则。")
    return name


def _resolve_sheet_name(sheet_names, requested: Any, label: str = "工作簿", allow_blank: bool = True) -> str:
    names = [str(name) for name in sheet_names]
    if not names:
        raise ProcessingError(f"{label}没有可读取的工作表。")
    requested_name = str(requested or "").strip()
    if not requested_name:
        if allow_blank:
            return names[0]
        raise ProcessingError(f"{label}工作表名不能为空。")
    requested_key = requested_name.casefold()
    for name in names:
        if name.casefold() == requested_key:
            return name
    raise ProcessingError(f"{label}中找不到配置的工作表“{requested_name}”；工作表名匹配不区分大小写。")


def _com_worksheet_names(workbook) -> list[str]:
    collection = workbook.Worksheets
    count = int(collection.Count)
    return [str(collection.Item(index).Name) for index in range(1, count + 1)]


def _resolve_com_worksheet(workbook, requested: Any, label: str, log=None):
    requested_name = str(requested or "").strip()
    actual_name = _resolve_sheet_name(_com_worksheet_names(workbook), requested_name, label)
    if not requested_name and log is not None:
        log(f"{label}未填写工作表名，默认使用第一个工作表“{actual_name}”。")
    return workbook.Worksheets.Item(actual_name)


def _parse_area(value: Any, label: str) -> tuple[int, int, int, int]:
    """Parse one worksheet-local A1 rectangle as (min_col, min_row, max_col, max_row)."""
    text = str(value or "").strip().upper()
    column_range = re.fullmatch(r"\$?([A-Z]{1,3})(?::\$?([A-Z]{1,3}))?", text)
    if column_range:
        min_col = column_number(column_range.group(1))
        max_col = column_number(column_range.group(2) or column_range.group(1))
        if min_col > max_col:
            raise ProcessingError(f"{label}区域“{text}”的起点不能在终点之后。")
        return min_col, 1, max_col, 1_048_576
    if not re.fullmatch(r"\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})?", text):
        raise ProcessingError(f"{label}区域“{text or '(空)'}”无效；请填写 A1 或 A1:C10 这样的单一单元格区域。")
    try:
        bounds = range_boundaries(text)
    except ValueError as exc:
        raise ProcessingError(f"{label}区域“{text}”无效。") from exc
    if any(value is None for value in bounds):
        raise ProcessingError(f"{label}区域“{text}”必须包含起止行和列。")
    min_col, min_row, max_col, max_row = bounds
    if min_col < 1 or min_row < 1 or max_col > 16_384 or max_row > 1_048_576:
        raise ProcessingError(f"{label}区域“{text}”超出 Excel 工作表边界。")
    if min_col > max_col or min_row > max_row:
        raise ProcessingError(f"{label}区域“{text}”的起点不能在终点之后。")
    return min_col, min_row, max_col, max_row


def _parse_selection_details(value: Any, label: str = "区域") -> tuple[tuple[int, int, int, int], str]:
    """Parse the shared UI/JSON selector, normalizing reversed endpoint rectangles."""
    text = str(value or "").strip().upper()
    coordinate = r"\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}"

    def point(token: str) -> tuple[int, int]:
        token = token.strip().upper()
        if not re.fullmatch(coordinate, token):
            raise ProcessingError(f"{label}端点“{token or '(空)'}”无效；请使用如 A5、D1 或 B3,B3 的 Excel 坐标。")
        try:
            col, row = range_boundaries(token)[0:2]
            column_number(get_column_letter(col))
            _positive_row(row, f"{label}行号")
        except (ValueError, TypeError) as exc:
            raise ProcessingError(f"{label}端点“{token}”超出 Excel 范围（XFD 列、1048576 行）。") from exc
        return col, row

    if "," in text:
        endpoints = text.split(",")
        if len(endpoints) != 2:
            raise ProcessingError(f"{label}“{text}”只允许两个逗号分隔的端点，例如 B3,D8 或 B3,B3。")
        first, second = point(endpoints[0]), point(endpoints[1])
        if first == second:
            return (first[0], first[1], first[0], first[1]), "cell"
        return (min(first[0], second[0]), min(first[1], second[1]),
                max(first[0], second[0]), max(first[1], second[1])), "rectangle"

    if ":" in text:
        endpoints = text.split(":")
        if len(endpoints) != 2:
            raise ProcessingError(f"{label}“{text}”格式无效；请使用两个端点。")
        first, second = point(endpoints[0]), point(endpoints[1])
        if first == second:
            return (first[0], first[1], first[0], first[1]), "cell"
        return (min(first[0], second[0]), min(first[1], second[1]),
                max(first[0], second[0]), max(first[1], second[1])), "rectangle"

    col, row = point(text)
    if col == 1 and row == 1:
        return (1, 1, 16_384, 1_048_576), "sheet"
    if col == 1:
        return (1, row, 16_384, row), "row"
    if row == 1:
        return (col, 1, col, 1_048_576), "column"
    return (col, row, col, row), "cell"


def _parse_selected_area(value: Any, label: str = "区域") -> tuple[int, int, int, int]:
    return _parse_selection_details(value, label)[0]


def _parse_region_filter_area(value: Any, label: str = "区域筛选") -> tuple[tuple[int, int, int, int], str]:
    """Parse filter-only shorthand while preserving other selector semantics."""
    text = str(value or "").strip()
    if not text:
        return (1, 1, 16_384, 1_048_576), "sheet"
    if re.fullmatch(r"[A-Za-z]{1,3}", text):
        column = column_number(text)
        return (column, 1, column, 1_048_576), "column"
    if re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}", text):
        return _parse_area(text, label), "cell"
    return _parse_selection_details(text, label)


def _last_used_row_from_com(worksheet, label: str) -> int:
    """Return the final row of a COM worksheet's used range, failing safely."""
    try:
        used_range = worksheet.UsedRange
        first_row = int(used_range.Row)
        row_count = int(used_range.Rows.Count)
        return max(1, first_row + row_count - 1)
    except Exception as exc:
        raise ProcessingError(f"无法读取{label}的已用行范围；整列筛选已停止。原因：{exc}") from exc


def _used_range_bounds_from_com(worksheet, label: str) -> tuple[int, int, int, int]:
    """Read the bounded used rectangle of a WPS worksheet for whole-sheet selectors."""
    try:
        used_range = worksheet.UsedRange
        min_row = max(1, int(used_range.Row))
        min_col = max(1, int(used_range.Column))
        row_count = max(1, int(used_range.Rows.Count))
        col_count = max(1, int(used_range.Columns.Count))
        return (min_col, min_row, min_col + col_count - 1, min_row + row_count - 1)
    except Exception as exc:
        raise ProcessingError(f"无法读取{label}的 WPS UsedRange；整表/整行/整列处理已停止。原因：{exc}") from exc


def _resolve_com_selection_bounds(worksheet, bounds: tuple[int, int, int, int],
                                  selection_kind: str, label: str) -> tuple[int, int, int, int]:
    if selection_kind not in ("sheet", "row", "column"):
        return bounds
    used_min_col, used_min_row, used_max_col, used_max_row = _used_range_bounds_from_com(worksheet, label)
    if selection_kind == "sheet":
        return used_min_col, used_min_row, used_max_col, used_max_row
    if selection_kind == "row":
        return used_min_col, bounds[1], used_max_col, bounds[1]
    return bounds[0], 1, bounds[0], used_max_row


def _area_cell_count(bounds: tuple[int, int, int, int]) -> int:
    return (bounds[2] - bounds[0] + 1) * (bounds[3] - bounds[1] + 1)


def _check_processing_area_size(bounds: tuple[int, int, int, int], label: str) -> None:
    count = _area_cell_count(bounds)
    if count > MAX_PROCESSING_CELLS:
        raise ProcessingError(f"{label}区域包含 {count:,} 个单元格，超过处理上限 {MAX_PROCESSING_CELLS:,}；请缩小范围。")


def _check_text_area_size(bounds: tuple[int, int, int, int], label: str) -> None:
    count = _area_cell_count(bounds)
    if count > MAX_TEXT_CELLS:
        raise ProcessingError(f"{label}包含 {count:,} 个单元格，超过文本预览/复制上限 {MAX_TEXT_CELLS:,}；请缩小范围。")


def _check_region_xlsx_area_size(bounds: tuple[int, int, int, int], label: str) -> None:
    cells = _area_cell_count(bounds)
    if cells > MAX_REGION_XLSX_CELLS:
        raise ProcessingError(
            f"{label}区域包含 {cells:,} 个单元格，超过独立 XLSX 区域输出上限 "
            f"{MAX_REGION_XLSX_CELLS:,}；请缩小范围。"
        )


def _area_address(bounds: tuple[int, int, int, int]) -> str:
    min_col, min_row, max_col, max_row = bounds
    if min_row == 1 and max_row == 1_048_576:
        first_col = get_column_letter(min_col)
        last_col = get_column_letter(max_col)
        return f"{first_col}:{last_col}"
    first = f"{get_column_letter(min_col)}{min_row}"
    last = f"{get_column_letter(max_col)}{max_row}"
    return first if first == last else f"{first}:{last}"


def _is_full_column_area(bounds: tuple[int, int, int, int]) -> bool:
    return bounds[1] == 1 and bounds[3] == 1_048_576


def _expand_bounds_for_merged_cells(worksheet, bounds: tuple[int, int, int, int],
                                    label: str) -> tuple[int, int, int, int]:
    """Expand a selected area until it completely contains every intersecting merged cell."""
    expanded = bounds
    while True:
        previous = expanded
        for merged in worksheet.merged_cells.ranges:
            merged_bounds = (merged.min_col, merged.min_row, merged.max_col, merged.max_row)
            if not _areas_overlap(expanded, merged_bounds):
                continue
            expanded = (min(expanded[0], merged_bounds[0]), min(expanded[1], merged_bounds[1]),
                        max(expanded[2], merged_bounds[2]), max(expanded[3], merged_bounds[3]))
            _check_region_xlsx_area_size(
                expanded, f"{label}完整包含合并单元格后的 {_area_address(expanded)}"
            )
        if expanded == previous:
            return expanded


def _areas_overlap(first: tuple[int, int, int, int], second: tuple[int, int, int, int]) -> bool:
    return (first[0] <= second[2] and second[0] <= first[2]
            and first[1] <= second[3] and second[1] <= first[3])


def _template_output_filename(template_path: Any) -> str:
    """Return the template basename in a platform-independent way, including its suffix."""
    raw = str(template_path or "").strip().replace("\\", "/")
    return raw.rsplit("/", 1)[-1] if raw else ""


def _filename_is_custom(raw_settings: dict[str, Any], filename: str) -> bool:
    marker = raw_settings.get("output_filename_custom")
    if marker is not None:
        if isinstance(marker, str):
            return marker.strip().lower() in ("1", "true", "yes", "on")
        return bool(marker)
    # Older configurations did not record whether the field was auto-generated.
    # Treat the former generic default and the current template basename as automatic.
    template_name = _template_output_filename(raw_settings.get("template_path"))
    return bool(filename and filename not in ("处理结果.xlsx", template_name))


def normalize_settings(raw: dict[str, Any]) -> dict[str, Any]:
    """Return the current schema; migrate the fixed three-source v1 schema in memory."""
    if not isinstance(raw, dict):
        raise ProcessingError("配置根节点必须是 JSON 对象。")
    settings = deepcopy(DEFAULT_SETTINGS)
    settings.update({k: deepcopy(v) for k, v in raw.items() if k not in ("jobs", "_migration_notices")})
    try:
        source_config_version = int(raw.get("config_version", 1))
    except (TypeError, ValueError):
        source_config_version = 1
    if "region_xlsx_outputs" not in raw:
        settings["region_xlsx_outputs"] = deepcopy(raw.get("image_outputs", []))
    settings.pop("image_outputs", None)
    settings.pop("export_region_pngs", None)
    raw_jobs = raw.get("jobs", settings["jobs"])
    if not isinstance(raw_jobs, list):
        raise ProcessingError("配置中的 jobs 必须是数组。")
    old_sources = raw.get("sources")
    is_legacy = any(isinstance(job, dict) and "sources" not in job for job in raw_jobs)
    if is_legacy and isinstance(old_sources, list):
        migrated = []
        global_start = raw.get("source_start_row", 2)
        for idx, old in enumerate(raw_jobs):
            if not isinstance(old, dict):
                migrated.append(old)
                continue
            job = deepcopy(old)
            src_path = old_sources[idx] if idx < len(old_sources) else ""
            job["sources"] = [{"path": src_path, "source_sheet": job.pop("source_sheet", ""),
                               "copy_start_row": global_start, "copy_end_row": None}]
            migrated.append(job)
        raw_jobs = migrated
    normalized_jobs = []
    for original in raw_jobs:
        if not isinstance(original, dict):
            normalized_jobs.append(original)
            continue
        job = deepcopy(original)
        for legacy_key in LEGACY_FORMULA_ROW_KEYS:
            job.pop(legacy_key, None)
        legacy_result_col = job.pop("result_col", "")
        legacy_target_col = job.pop("target_col", "")
        mappings = job.get("mappings")
        if mappings is None:
            mappings = []
            if str(legacy_result_col or "").strip() or str(legacy_target_col or "").strip():
                mappings.append({"result_col": legacy_result_col, "target_col": legacy_target_col,
                                 "write_as": "value"})
        elif isinstance(mappings, list):
            mappings = [dict(mapping) if isinstance(mapping, dict) else mapping for mapping in mappings]
        job["mappings"] = mappings
        job["merged_source_filename"] = str(job.get("merged_source_filename") or "").strip()
        source_items = job.get("sources", [])
        if isinstance(source_items, list):
            converted = []
            for source in source_items:
                if isinstance(source, str):
                    if not source.strip():
                        continue
                    converted.append({"path": source, "source_sheet": "", "copy_start_row": raw.get("source_start_row", 2), "copy_end_row": None,
                                     "keep_first_row": True})
                elif isinstance(source, dict):
                    item = {"path": "", "source_sheet": "", "copy_start_row": raw.get("source_start_row", 2), "copy_end_row": None,
                            "keep_first_row": True}
                    item.update(source)
                    if not str(item.get("path") or "").strip():
                        continue
                    item.pop("_optional_placeholder", None)
                    item["keep_first_row"] = bool(item.get("keep_first_row", True))
                    converted.append(item)
                else:
                    converted.append(source)
            job["sources"] = converted
        normalized_jobs.append(job)
    settings["jobs"] = normalized_jobs
    settings["merge_only"] = raw.get("merge_only", settings.get("merge_only", False))
    settings["merge_output_filename"] = str(raw.get("merge_output_filename", settings.get("merge_output_filename", "合并结果.xlsx")) or "合并结果.xlsx").strip()
    if settings["merge_output_filename"] and not Path(settings["merge_output_filename"]).suffix:
        settings["merge_output_filename"] += ".xlsx"
    raw_merge_sources = raw.get("merge_sources", settings.get("merge_sources", []))
    if not isinstance(raw_merge_sources, list):
        raise ProcessingError("配置中的 merge_sources 必须是数组。")
    normalized_merge_sources = []
    for source_index, original in enumerate(raw_merge_sources):
        if not isinstance(original, dict):
            normalized_merge_sources.append(original)
            continue
        source = {"path": "", "source_sheet": "", "copy_start_row": 1,
                  "copy_end_row": None, "keep_first_row": True}
        source.update(deepcopy(original))
        if not str(source.get("path") or "").strip():
            continue
        source["keep_first_row"] = True if not normalized_merge_sources else bool(source.get("keep_first_row", True))
        normalized_merge_sources.append(source)
    settings["merge_sources"] = normalized_merge_sources
    # The old auto-open switch is not supported; configured XLSX outputs remain.
    settings.pop("open_region_exports", None)
    for key in ("pre_mappings", "post_mappings", "region_xlsx_outputs", "text_outputs"):
        if settings.get(key) is None:
            settings[key] = []
    retain_merged_sources = bool(raw.get(
        "retain_merged_sources", raw.get("retain_edited_sources", settings.get("retain_merged_sources", False))
    ))
    settings["retain_merged_sources"] = retain_merged_sources
    settings["retain_edited_sources"] = retain_merged_sources
    settings["delete_processed_sources_after_success"] = bool(raw.get(
        "delete_processed_sources_after_success",
        settings.get("delete_processed_sources_after_success", False),
    ))
    for key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
        operation = settings.get(key)
        if operation is None:
            settings[key] = {"enabled": False, "rules": []}
        elif isinstance(operation, dict):
            settings[key] = {"enabled": operation.get("enabled", False),
                             "rules": deepcopy(operation.get("rules", []))}
    if settings.get("workflow") is None:
        settings["workflow"] = list(DEFAULT_WORKFLOW)
    if not isinstance(settings.get("workflow"), list):
        raise ProcessingError("配置中的 workflow 必须是模块 ID 数组。")
    settings["workflow"] = ["region_xlsx_output" if item == "image_output" else item
                             for item in settings["workflow"]]
    if "source_merge" not in settings["workflow"]:
        settings["workflow"].insert(0, "source_merge")
    if source_config_version < 13 and "result_mapping" not in settings["workflow"]:
        insert_at = settings["workflow"].index("formula_recalc") + 1 if "formula_recalc" in settings["workflow"] else len(settings["workflow"])
        settings["workflow"].insert(insert_at, "result_mapping")
    if "source_preprocess" not in settings["workflow"]:
        settings["workflow"].insert(settings["workflow"].index("source_merge") + 1, "source_preprocess")
    if "CustomFormula" not in settings["workflow"]:
        insertion_candidates = [settings["workflow"].index(module_id) for module_id in
                                ("TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert",
                                "final_recalc_save", "region_xlsx_output", "text_output")
                                if module_id in settings["workflow"]]
        settings["workflow"].insert(min(insertion_candidates, default=len(settings["workflow"])), "CustomFormula")
    if settings["region_xlsx_outputs"] and "region_xlsx_output" not in settings["workflow"]:
        settings["workflow"].append("region_xlsx_output")
    try:
        prefix_length = int(settings.get("template_prefix_length", 5))
    except (TypeError, ValueError) as exc:
        raise ProcessingError("模板匹配前缀长度 N 必须是整数。") from exc
    if not 1 <= prefix_length <= 120:
        raise ProcessingError("模板匹配前缀长度 N 必须介于 1 和 120。")
    settings["template_prefix_length"] = prefix_length
    settings["config_name"] = str(settings.get("config_name") or "").strip()
    filename = str(settings.get("output_filename") or "").strip()
    is_custom_filename = _filename_is_custom(raw, filename)
    settings["output_filename_custom"] = is_custom_filename
    if not is_custom_filename:
        template_filename = _template_output_filename(settings.get("template_path"))
        if template_filename:
            settings["output_filename"] = template_filename
    settings.pop("sources", None)
    settings["config_version"] = CONFIG_VERSION
    return settings


def _validate_merge_only_static_settings(settings: dict[str, Any]) -> None:
    if not isinstance(settings.get("merge_only"), bool):
        raise ProcessingError("merge_only 必须是布尔值。")
    if not str(settings.get("output_dir") or "").strip():
        raise ProcessingError("独立合并模式需要选择输出目录。")
    sources = settings.get("merge_sources")
    if not isinstance(sources, list) or not sources:
        raise ProcessingError("请至少添加一份独立合并源文件。")
    output_name = str(settings.get("merge_output_filename") or "").strip()
    if not output_name:
        raise ProcessingError("合并输出文件名不能为空。")
    if Path(output_name).name != output_name or "/" in output_name or "\\" in output_name:
        raise ProcessingError("合并输出文件名只能填写文件名，不能包含目录路径。")
    if Path(output_name).suffix and Path(output_name).suffix.lower() != ".xlsx":
        raise ProcessingError("合并结果必须保存为 .xlsx 文件。")
    if any(character in output_name for character in '<>:"|?*') or output_name.endswith((" ", ".")):
        raise ProcessingError("合并输出文件名包含 Windows 不支持的字符，或以空格/句点结尾。")
    reserved_stem = Path(output_name).stem.rstrip(" .").upper()
    if reserved_stem in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                         *(f"LPT{i}" for i in range(1, 10))}:
        raise ProcessingError("合并输出文件名不能使用 Windows 保留设备名。")
    for index, source in enumerate(sources):
        if not isinstance(source, dict):
            raise ProcessingError(f"第 {index + 1} 份独立合并源文件配置无效。")
        path_text = str(source.get("path") or "").strip()
        if not path_text:
            raise ProcessingError(f"第 {index + 1} 份独立合并源文件尚未选择。")
        suffix = Path(path_text).suffix.lower()
        if index == 0:
            if suffix not in (".csv", ".xlsx"):
                raise ProcessingError("独立合并的第一份文件必须是 .csv 或 .xlsx。")
            if suffix == ".csv" and str(source.get("source_sheet") or "").strip():
                raise ProcessingError("CSV 没有工作表；第一份文件为 CSV 时请将工作表名留空。")
        elif suffix != ".xlsx":
            raise ProcessingError(f"独立合并的第 {index + 1} 份及后续文件必须是 .xlsx。")
        _positive_row(source.get("copy_start_row", 1), f"第 {index + 1} 份独立合并源文件的复制起始行")
        end_row = source.get("copy_end_row")
        if end_row not in (None, ""):
            end_row = _positive_row(end_row, f"第 {index + 1} 份独立合并源文件的复制结束行")
            if end_row < _positive_row(source.get("copy_start_row", 1), "复制起始行"):
                raise ProcessingError(f"第 {index + 1} 份独立合并源文件结束行不得早于起始行。")
        if not isinstance(source.get("keep_first_row", True), bool):
            raise ProcessingError(f"第 {index + 1} 份独立合并源文件的首行保留设置必须为布尔值。")


def validate_static_settings(raw_settings: dict[str, Any]) -> None:
    settings = normalize_settings(raw_settings)
    if settings.get("merge_only"):
        _validate_merge_only_static_settings(settings)
        return
    try:
        validate_workflow(settings)
    except ValueError as exc:
        raise ProcessingError(str(exc)) from exc
    jobs = settings.get("jobs")
    if not isinstance(jobs, list) or not jobs:
        raise ProcessingError("至少需要配置一个目标工作表。")
    has_sources = any(isinstance(job, dict) and job.get("sources") for job in jobs)
    copy_start = column_number(str(settings.get("copy_start_col", "A")).strip() or "A") if has_sources else 1
    copy_end_value = str(settings.get("copy_end_col", "")).strip()
    copy_range = None
    if has_sources and copy_end_value:
        copy_end = column_number(copy_end_value)
        if copy_start > copy_end:
            raise ProcessingError("数据复制列范围的起始列不能在末列之后。")
        copy_range = (copy_start, copy_end)
    names: list[str] = []
    for idx, job in enumerate(jobs, start=1):
        if not isinstance(job, dict):
            raise ProcessingError(f"第 {idx} 个目标工作表配置必须是对象。")
        for key in ("template_sheet", "header_rows", "start_row", "formula_range", "sources"):
            if key not in job:
                raise ProcessingError(f"第 {idx} 个目标工作表配置缺少“{key}”。")
        sheet_name = str(job["template_sheet"]).strip()
        if not sheet_name:
            raise ProcessingError(f"第 {idx} 个目标工作表名称不能为空。")
        if len(sheet_name) > 31 or any(char in sheet_name for char in "[]:*?/\\"):
            raise ProcessingError(f"第 {idx} 个目标工作表名称不符合 Excel 工作表命名规则。")
        names.append(sheet_name)
        _validate_merged_source_filename(job.get("merged_source_filename", ""), idx)
        header_rows = _positive_row(job["header_rows"], f"第 {idx} 项保留标题行数", allow_zero=True)
        start_row = _positive_row(job["start_row"], f"第 {idx} 项目标粘贴起始行")
        if start_row <= 1:
            raise ProcessingError(f"第 {idx} 项目标粘贴起始行必须至少为第 2 行；公式模板行自动取其上一行。")
        if start_row <= header_rows:
            raise ProcessingError(f"第 {idx} 项目标起始行必须大于保留标题行数。")
        formula_range = parse_column_range(job["formula_range"])
        mappings = job.get("mappings", [])
        if not isinstance(mappings, list):
            raise ProcessingError(f"第 {idx} 个目标表的 mappings 必须是结果映射规则数组。")
        mapped_targets: dict[int, int] = {}
        for map_idx, mapping in enumerate(mappings, 1):
            if not isinstance(mapping, dict):
                raise ProcessingError(f"第 {idx} 个目标表的第 {map_idx} 条结果映射必须是对象。")
            try:
                result_col = column_number(mapping.get("result_col", ""))
                target_col = column_number(mapping.get("target_col", ""))
            except ProcessingError as exc:
                raise ProcessingError(f"第 {idx} 个目标表第 {map_idx} 条结果映射：{exc}") from exc
            if not formula_range[0] <= result_col <= formula_range[1]:
                raise ProcessingError(f"第 {idx} 个目标表结果列 {mapping['result_col']} 不在公式范围 {job['formula_range']} 内。")
            if formula_range[0] <= target_col <= formula_range[1]:
                raise ProcessingError(f"第 {idx} 个目标表第 {map_idx} 条结果映射的目标列 {mapping['target_col']} 位于公式范围 {job['formula_range']} 内；写回会覆盖必须保留的公式列。")
            if target_col in mapped_targets:
                raise ProcessingError(f"第 {idx} 个目标表第 {mapped_targets[target_col]} 条和第 {map_idx} 条结果映射都写入目标列 {mapping['target_col']}，会互相覆盖。")
            mapped_targets[target_col] = map_idx
            mode = str(mapping.get("write_as", "value")).strip().lower()
            if mode not in ("value", "text"):
                raise ProcessingError(f"第 {idx} 个目标表第 {map_idx} 条结果映射的写入类型只能是 value（值）或 text（文本）。")
        if copy_range is not None and ranges_overlap(copy_range, formula_range):
            raise ProcessingError(f"第 {idx} 项数据复制列范围与公式区域 {job['formula_range']} 重叠。")
        sources = job["sources"]
        if not isinstance(sources, list):
            raise ProcessingError(f"第 {idx} 个目标表的源文件配置必须是列表；可留空以单独编辑模板。")
        for source_idx, source in enumerate(sources, 1):
            if not isinstance(source, dict):
                raise ProcessingError(f"第 {idx} 个目标表的第 {source_idx} 份源文件配置无效。")
            if "path" not in source or "source_sheet" not in source:
                raise ProcessingError(f"第 {idx} 个目标表的第 {source_idx} 份源文件需要 path 和 source_sheet。")
            src_start = _positive_row(source.get("copy_start_row"), f"第 {idx} 个目标表第 {source_idx} 份源文件的复制起始行")
            src_end = source.get("copy_end_row")
            if src_end not in (None, ""):
                src_end = _positive_row(src_end, f"第 {idx} 个目标表第 {source_idx} 份源文件的复制结束行")
                if src_end < src_start:
                    raise ProcessingError(f"第 {idx} 个目标表第 {source_idx} 份源文件结束行不得早于起始行。")
    if len(set(names)) != len(names):
        raise ProcessingError("目标工作表名称必须互不相同。")

    for key, label in (("pre_mappings", "前映射"), ("post_mappings", "后映射")):
        rules = settings.get(key, [])
        if not isinstance(rules, list):
            raise ProcessingError(f"配置中的 {key} 必须是数组。")
        destinations: dict[str, list[tuple[int, tuple[int, int, int, int]]]] = {}
        for rule_idx, rule in enumerate(rules, 1):
            if not isinstance(rule, dict):
                raise ProcessingError(f"第 {rule_idx} 条{label}规则必须是对象。")
            for field in ("source_sheet", "source_range", "target_sheet", "target_range"):
                if field not in rule:
                    raise ProcessingError(f"第 {rule_idx} 条{label}规则缺少“{field}”。")
            _validate_sheet_name(rule["source_sheet"], f"第 {rule_idx} 条{label}源")
            target_sheet = _validate_sheet_name(rule["target_sheet"], f"第 {rule_idx} 条{label}目标")
            source_bounds = _parse_area(rule["source_range"], f"第 {rule_idx} 条{label}源")
            target_bounds = _parse_area(rule["target_range"], f"第 {rule_idx} 条{label}目标")
            source_shape = (source_bounds[2] - source_bounds[0], source_bounds[3] - source_bounds[1])
            target_shape = (target_bounds[2] - target_bounds[0], target_bounds[3] - target_bounds[1])
            if source_shape != target_shape:
                raise ProcessingError(f"第 {rule_idx} 条{label}规则的源/目标区域行列数必须相同。")
            for previous_idx, previous in destinations.setdefault(target_sheet, []):
                if _areas_overlap(previous, target_bounds):
                    raise ProcessingError(f"第 {previous_idx} 条和第 {rule_idx} 条{label}规则的目标区域重叠：{target_sheet}!{_area_address(previous)} 与 {_area_address(target_bounds)}。")
            destinations[target_sheet].append((rule_idx, target_bounds))

    outputs = settings.get("region_xlsx_outputs", [])
    if not isinstance(outputs, list):
        raise ProcessingError("配置中的 region_xlsx_outputs 必须是数组。")
    for output_idx, output in enumerate(outputs, 1):
        if not isinstance(output, dict):
            raise ProcessingError(f"第 {output_idx} 项区域 XLSX 输出配置必须是对象。")
        if "sheet" not in output or "range" not in output:
            raise ProcessingError(f"第 {output_idx} 项区域 XLSX 输出配置需要 sheet 和 range。")
        _validate_sheet_name(output["sheet"], f"第 {output_idx} 项区域 XLSX 输出")
        bounds, selection_kind = _parse_selection_details(output["range"], f"第 {output_idx} 项区域 XLSX 输出")
        if selection_kind not in ("sheet", "row", "column"):
            _check_region_xlsx_area_size(bounds, f"第 {output_idx} 项区域 XLSX 输出")
    outputs = settings.get("text_outputs", [])
    if not isinstance(outputs, list):
        raise ProcessingError("配置中的 text_outputs 必须是数组。")
    for output_idx, output in enumerate(outputs, 1):
        if not isinstance(output, dict):
            raise ProcessingError(f"第 {output_idx} 项文本输出配置必须是对象。")
        if "sheet" not in output or "range" not in output:
            raise ProcessingError(f"第 {output_idx} 项文本输出配置需要 sheet 和 range。")
        _validate_sheet_name(output["sheet"], f"第 {output_idx} 项文本输出")
        bounds, selection_kind = _parse_selection_details(output["range"], f"第 {output_idx} 项文本输出")
        if selection_kind not in ("sheet", "row", "column"):
            _check_text_area_size(bounds, f"第 {output_idx} 项文本输出")

    operation_labels = {"CustomFormula": "公式填充", "TextToColumns": "文本分列", "RegionFilter": "区域筛选",
                        "Deduplicate": "去重", "FormatConvert": "格式转换"}
    for key, label in operation_labels.items():
        operation = settings.get(key)
        if not isinstance(operation, dict) or not isinstance(operation.get("enabled", False), bool):
            raise ProcessingError(f"{label}必须包含布尔 enabled 开关和 rules 数组。")
        rules = operation.get("rules", [])
        if not isinstance(rules, list):
            raise ProcessingError(f"{label}的 rules 必须是数组。")
        for idx, rule in enumerate(rules, 1):
            if not isinstance(rule, dict):
                raise ProcessingError(f"第 {idx} 条{label}规则必须是对象。")
            _validate_sheet_name(rule.get("sheet"), f"第 {idx} 条{label}", allow_blank=True)
            scope = str(rule.get("scope", "target")).strip().lower()
            if scope not in ("target", "source"):
                raise ProcessingError(f"第 {idx} 条{label}规则的 scope 只能是 target（目标工作簿）或 source（源文件）。")
            if key == "RegionFilter":
                bounds, selection_kind = _parse_region_filter_area(rule.get("range"), f"第 {idx} 条{label}")
            else:
                bounds, selection_kind = _parse_selection_details(rule.get("range"), f"第 {idx} 条{label}")
            if selection_kind not in ("column", "row", "sheet"):
                _check_processing_area_size(bounds, f"第 {idx} 条{label}")
            if key == "TextToColumns":
                if bounds[0] != bounds[2]:
                    raise ProcessingError(f"第 {idx} 条文本分列规则的源范围必须是单列。")
                if not str(rule.get("delimiter", "")):
                    raise ProcessingError(f"第 {idx} 条文本分列规则必须指定非空分隔符。")
                column_number(rule.get("output_start_col", ""))
                formats = rule.get("column_formats", ["text"])
                if not isinstance(formats, list) or not formats:
                    raise ProcessingError(f"第 {idx} 条文本分列规则须配置每个结果列的格式。")
                if any(str(fmt).lower() not in ("text", "number", "date") for fmt in formats):
                    raise ProcessingError(f"第 {idx} 条文本分列规则的列格式只能是 text、number 或 date。")
            elif key == "RegionFilter":
                if not isinstance(rule.get("has_header", False), bool):
                    raise ProcessingError(f"第 {idx} 条区域筛选规则的 has_header 必须是布尔值。")
                filter_col = column_number(rule.get("filter_column", ""))
                if not bounds[0] <= filter_col <= bounds[2]:
                    raise ProcessingError(f"第 {idx} 条区域筛选规则的条件列必须位于筛选区域内。")
                operator = str(rule.get("operator", "equals")).strip().lower()
                if operator not in FILTER_OPERATORS:
                    raise ProcessingError(f"第 {idx} 条区域筛选规则的条件类型无效。")
                if operator not in ("blank", "not_blank") and not str(rule.get("criteria", "")):
                    raise ProcessingError(f"第 {idx} 条区域筛选规则必须填写条件值。")
            elif key == "Deduplicate":
                if not isinstance(rule.get("has_header", False), bool):
                    raise ProcessingError(f"第 {idx} 条去重规则的 has_header 必须是布尔值。")
                key_mode = str(rule.get("key_mode", "row")).lower()
                if key_mode not in ("row", "columns"):
                    raise ProcessingError(f"第 {idx} 条去重规则的 key_mode 只能是 row 或 columns。")
                keys = rule.get("key_columns", [])
                if key_mode == "columns":
                    if not isinstance(keys, list) or not keys:
                        raise ProcessingError(f"第 {idx} 条去重规则须指定至少一个关键列。")
                    seen = set()
                    for key_col in keys:
                        col = column_number(key_col)
                        if not bounds[0] <= col <= bounds[2]:
                            raise ProcessingError(f"第 {idx} 条去重规则关键列 {key_col} 不在所选区域内。")
                        if col in seen:
                            raise ProcessingError(f"第 {idx} 条去重规则重复指定关键列 {key_col}。")
                        seen.add(col)
            elif key == "FormatConvert":
                if str(rule.get("mode", "")).lower() not in ("text", "number", "date"):
                    raise ProcessingError(f"第 {idx} 条格式转换规则 mode 只能是 text、number 或 date。")
            elif key == "CustomFormula":
                formula_mode = str(rule.get("formula_mode", "custom")).strip().lower()
                if formula_mode == "custom":
                    formula = str(rule.get("formula", "")).strip()
                    if not formula.startswith("="):
                        raise ProcessingError(f"第 {idx} 条公式填充规则的自定义公式必须以 = 开头。")
                    if len(formula) > 8_192:
                        raise ProcessingError(f"第 {idx} 条公式填充规则超过 WPS/Excel 公式长度上限 8192 字符。")
                    if _find_external_formula(formula):
                        raise ProcessingError(f"第 {idx} 条公式填充规则包含外部工作簿引用或外部访问函数，已拒绝。")
                elif formula_mode == "cell":
                    seed_bounds = _parse_area(rule.get("seed_cell", ""), f"第 {idx} 条公式来源单元格")
                    if seed_bounds[0] != seed_bounds[2] or seed_bounds[1] != seed_bounds[3]:
                        raise ProcessingError(f"第 {idx} 条公式填充规则的公式来源必须是单个单元格，例如 B2。")
                else:
                    raise ProcessingError(f"第 {idx} 条公式填充规则的 formula_mode 只能是 custom 或 cell。")

    _validate_region_module_dependencies(settings)


def _validate_region_module_dependencies(settings: dict[str, Any]) -> None:
    workflow = settings.get("workflow", DEFAULT_WORKFLOW)
    positions = {module_id: index for index, module_id in enumerate(workflow)}
    split_operation = settings.get("TextToColumns", {})
    if not split_operation.get("enabled", False) or "TextToColumns" not in positions:
        return
    for later_key in ("RegionFilter", "Deduplicate", "FormatConvert"):
        operation = settings.get(later_key, {})
        if not operation.get("enabled", False) or later_key not in positions:
            continue
        for split_idx, split_rule in enumerate(split_operation.get("rules", []), 1):
            source = _parse_selected_area(split_rule["range"], f"第 {split_idx} 条文本分列")
            formats = split_rule.get("column_formats", ["text"])
            start_col = column_number(split_rule["output_start_col"])
            target = (start_col, source[1], start_col + len(formats) - 1, source[3])
            for rule_idx, rule in enumerate(operation.get("rules", []), 1):
                if str(rule["sheet"]).strip().casefold() != str(split_rule["sheet"]).strip().casefold():
                    continue
                if str(rule.get("scope", "target")) != str(split_rule.get("scope", "target")):
                    continue
                if later_key == "RegionFilter":
                    other = _parse_region_filter_area(rule["range"], f"第 {rule_idx} 条{MODULE_LABELS[later_key]}")[0]
                else:
                    other = _parse_selected_area(rule["range"], f"第 {rule_idx} 条{MODULE_LABELS[later_key]}")
                if _areas_overlap(source, other) or _areas_overlap(target, other):
                    if positions["TextToColumns"] >= positions[later_key]:
                        raise ProcessingError(
                            f"工作流顺序无效：“文本分列”必须位于“{MODULE_LABELS[later_key]}”之前；"
                            f"第 {split_idx} 条分列规则的输入/输出区域与第 {rule_idx} 条{MODULE_LABELS[later_key]}规则重叠，"
                            "否则分列输入会先被修改，或分列目标会先被占用。"
                        )


def _resource_path(relative: str) -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / relative


def load_settings(path: Path | str | None = None) -> dict[str, Any]:
    settings = deepcopy(DEFAULT_SETTINGS)
    source_path = Path(path) if path is not None else CONFIG_PATH
    if source_path.exists():
        try:
            saved = json.loads(source_path.read_text(encoding="utf-8"))
            settings = normalize_settings(saved)
            notices = configuration_migration_notes(saved)
            if notices:
                settings["_migration_notices"] = notices
        except Exception as exc:
            raise ProcessingError(f"读取配置失败：{source_path}\n{exc}") from exc
    return settings


def save_settings(settings: dict[str, Any], path: Path | str | None = None) -> None:
    normalized = normalize_settings(settings)
    validate_static_settings(normalized)
    target = Path(path) if path is not None else CONFIG_PATH
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")


def _check_xlsx_file(path: Path, label: str) -> None:
    if not path.is_file():
        raise ProcessingError(f"{label}不存在或不是文件：\n{path}")
    if path.suffix.lower() != ".xlsx":
        raise ProcessingError(f"{label}必须是 .xlsx 文件（不接受含宏工作簿或其他格式）：\n{path}")
    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = archive.namelist()
            if any(n.startswith("xl/externalLinks/") for n in names):
                raise ProcessingError(f"{label}包含外部工作簿链接；为避免访问外部内容，已拒绝处理：\n{path}")
            if "xl/vbaProject.bin" in names:
                raise ProcessingError(f"{label}包含 VBA 宏；程序只处理无宏 .xlsx 文件：\n{path}")
    except ProcessingError:
        raise
    except (OSError, zipfile.BadZipFile) as exc:
        raise ProcessingError(f"{label}不是可读取的有效 XLSX 工作簿：\n{path}\n{exc}") from exc


def _zip_member_name_is_safe(name: str) -> bool:
    if not name or "\x00" in name or "\\" in name or name.startswith("/"):
        return False
    relative = name[:-1] if name.endswith("/") else name
    if not relative:
        return False
    path = PurePosixPath(relative)
    if path.is_absolute() or any(part in (".", "..") for part in path.parts):
        return False
    if path.parts and ":" in path.parts[0]:
        return False
    return True


def _xml_root(data: bytes, description: str):
    compact = data.replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in compact or b"<!ENTITY" in compact:
        raise ValueError(f"{description}包含不支持的 DTD 或实体声明。")
    return ET.fromstring(data)


def _local_xml_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _read_bounded_zip_xml(archive: zipfile.ZipFile, name: str, limit: int) -> bytes:
    info = archive.getinfo(name)
    if info.file_size > limit:
        raise ValueError(f"工作簿 XML 部件 {name} 超过安全读取上限。")
    with archive.open(info, "r") as source:
        data = source.read(limit + 1)
    if len(data) > limit or len(data) != info.file_size:
        raise ValueError(f"工作簿 XML 部件 {name} 的实际大小异常。")
    return data


def _resolve_internal_relationship(base_part: str, target: str) -> str:
    split = urlsplit(target)
    if split.scheme or split.netloc or split.query or split.fragment:
        raise ValueError("工作表关系目标不是安全的包内路径。")
    decoded = unquote(split.path)
    if not decoded or "\\" in decoded or "\x00" in decoded:
        raise ValueError("工作表关系目标路径无效。")
    if decoded.startswith("/"):
        resolved = posixpath.normpath(decoded.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(base_part), decoded))
    if resolved in ("", ".", "..") or resolved.startswith("../") or not _zip_member_name_is_safe(resolved):
        raise ValueError("工作表关系目标越出了 XLSX 包目录。")
    return resolved


def _worksheet_parts(archive: zipfile.ZipFile, member_names: set[str]) -> dict[str, str]:
    workbook_part = "xl/workbook.xml"
    relationships_part = "xl/_rels/workbook.xml.rels"
    if workbook_part not in member_names or relationships_part not in member_names:
        raise ValueError("XLSX 缺少工作簿清单或工作表关系部件。")
    workbook_root = _xml_root(_read_bounded_zip_xml(archive, workbook_part, MAX_XLSX_METADATA_XML_BYTES), workbook_part)
    relationships_root = _xml_root(
        _read_bounded_zip_xml(archive, relationships_part, MAX_XLSX_METADATA_XML_BYTES), relationships_part
    )
    relationship_by_id = {}
    for relation in relationships_root.iter():
        if _local_xml_name(relation.tag) != "Relationship":
            continue
        relation_id = relation.attrib.get("Id")
        if relation_id:
            relationship_by_id[relation_id] = relation.attrib

    result: dict[str, str] = {}
    relationship_id_attributes = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
        "{http://purl.oclc.org/ooxml/officeDocument/relationships}id",
    )
    for sheet in workbook_root.iter():
        if _local_xml_name(sheet.tag) != "sheet":
            continue
        relation_id = next((sheet.attrib[key] for key in relationship_id_attributes if key in sheet.attrib), None)
        if not relation_id:
            continue
        relation = relationship_by_id.get(relation_id)
        if relation is None or not relation.get("Type", "").endswith("/worksheet"):
            continue
        if relation.get("TargetMode", "").lower() == "external":
            raise ValueError("工作表关系指向外部目标，无法安全检查。")
        part = _resolve_internal_relationship(workbook_part, relation.get("Target", ""))
        if part not in member_names:
            raise ValueError(f"工作表 XML 部件不存在：{part}")
        sheet_name = sheet.attrib.get("name", "(未命名工作表)")
        result[part] = sheet_name
    return result


def _valid_excel_a1_area(reference: str | None) -> bool:
    if not isinstance(reference, str) or not reference:
        return False
    cell = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
    column = r"\$?[A-Za-z]{1,3}"
    row = r"\$?[1-9][0-9]{0,6}"
    patterns = (rf"{cell}(?::{cell})?", rf"{column}:{column}", rf"{row}:{row}")
    if not any(re.fullmatch(pattern, reference) for pattern in patterns):
        return False
    try:
        min_col, min_row, max_col, max_row = range_boundaries(reference)
    except (TypeError, ValueError):
        return False
    return (
        (min_col is None or 1 <= min_col <= 16_384)
        and (max_col is None or 1 <= max_col <= 16_384)
        and (min_row is None or 1 <= min_row <= 1_048_576)
        and (max_row is None or 1 <= max_row <= 1_048_576)
    )


def _xml_start_tag_end(data: bytes, start: int) -> int:
    quote = None
    for index in range(start, len(data)):
        byte = data[index]
        if quote is not None:
            if byte == quote:
                quote = None
        elif byte in (ord("\""), ord("'")):
            quote = byte
        elif byte == ord(">"):
            return index + 1
    raise ValueError("无法定位 sortState XML 开始标签的结束位置。")


def _normalized_sort_state_ref(reference: str | None, sheet_name: str) -> str | None:
    if reference is None:
        raise ValueError(f"工作表“{sheet_name}”的 sortState 缺少 ref；为保护排序信息，未删除该排序定义。")
    row_range = re.fullmatch(r"([1-9][0-9]*):([1-9][0-9]*)", reference)
    if row_range:
        start_text, end_text = row_range.groups()
        if len(start_text) > 7 or len(end_text) > 7:
            raise ValueError(
                f"工作表“{sheet_name}”的 sortState ref {reference!r} 超出 Excel 行号范围；"
                "为保护排序信息，未删除该排序定义。"
            )
        start_row, end_row = int(start_text), int(end_text)
        if start_row > 1_048_576 or end_row > 1_048_576:
            raise ValueError(
                f"工作表“{sheet_name}”的 sortState ref {reference!r} 超出 Excel 行号范围；"
                "为保护排序信息，未删除该排序定义。"
            )
        if start_row > end_row:
            raise ValueError(
                f"工作表“{sheet_name}”的 sortState ref {reference!r} 起始行大于结束行；"
                "为保护排序信息，未删除该排序定义。"
            )
        return f"A{start_row}:XFD{end_row}"
    if _valid_excel_a1_area(reference):
        return None
    raise ValueError(
        f"工作表“{sheet_name}”的 sortState ref {reference!r} 不是有效 Excel 排序区域；"
        "无法安全规范化，为保护排序信息，未删除该排序定义。"
    )


def _preprocess_worksheet_xml(data: bytes, sheet_name: str) -> tuple[bytes, list[str], list[tuple[str, str]]]:
    if b"\x00" in data[:256]:
        raise ValueError("工作表 XML 不是可安全定位修改的 UTF-8 编码。")
    compact = data.replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in compact or b"<!ENTITY" in compact:
        raise ValueError("工作表 XML 包含不支持的 DTD 或实体声明。")

    parser = expat.ParserCreate(namespace_separator="}")
    stack: list[dict[str, Any]] = []
    invalid_filter_spans: list[tuple[int, int, str]] = []
    sort_ref_edits: list[tuple[int, int, bytes, str, str]] = []
    normalized_refs: list[tuple[str, str]] = []

    def start_element(name: str, attributes: dict[str, str]) -> None:
        depth = len(stack)
        node = {"name": name, "start": parser.CurrentByteIndex, "invalid_ref": None}
        if depth == 0:
            root_namespace = name.rsplit("}", 1)[0] if "}" in name else ""
            node["is_worksheet"] = _local_xml_name(name) == "worksheet"
            node["namespace"] = root_namespace
        elif stack[0].get("is_worksheet"):
            namespace = name.rsplit("}", 1)[0] if "}" in name else ""
            if namespace == stack[0].get("namespace") and _local_xml_name(name) == "autoFilter":
                reference = attributes.get("ref")
                if not _valid_excel_a1_area(reference):
                    node["invalid_ref"] = "<缺失>" if reference is None else reference
            if namespace == stack[0].get("namespace") and _local_xml_name(name) == "sortState":
                reference = attributes.get("ref")
                replacement = _normalized_sort_state_ref(reference, sheet_name)
                if replacement is not None:
                    start = int(node["start"])
                    tag_end = _xml_start_tag_end(data, start)
                    start_tag = data[start:tag_end]
                    matches = list(re.finditer(
                        rb"(?:^|[ \t\r\n])ref[ \t\r\n]*=[ \t\r\n]*(?:\"([^\"]*)\"|'([^']*)')",
                        start_tag,
                    ))
                    if len(matches) != 1:
                        raise ValueError(
                            f"工作表“{sheet_name}”的 sortState ref {reference!r} 无法在 XML 中安全定位；"
                            "为保护排序信息，未删除该排序定义。"
                        )
                    match = matches[0]
                    ref_group = 1 if match.group(1) is not None else 2
                    raw_ref = match.group(ref_group)
                    if raw_ref != reference.encode("ascii"):
                        raise ValueError(
                            f"工作表“{sheet_name}”的 sortState ref {reference!r} 使用了无法安全改写的 XML 编码；"
                            "为保护排序信息，未删除该排序定义。"
                        )
                    sort_ref_edits.append((
                        start + match.start(ref_group), start + match.end(ref_group),
                        replacement.encode("ascii"), reference, replacement,
                    ))
                    normalized_refs.append((reference, replacement))
        stack.append(node)

    def end_element(name: str) -> None:
        if not stack:
            raise ValueError("工作表 XML 元素结构异常。")
        node = stack.pop()
        if node["name"] != name or node["invalid_ref"] is None:
            return
        start = int(node["start"])
        event_position = parser.CurrentByteIndex
        if data[max(start, event_position - 2):event_position] == b"/>":
            end = event_position
        else:
            closing = data.find(b">", event_position)
            if closing < 0:
                raise ValueError("无法定位无效自动筛选定义的 XML 结束位置。")
            end = closing + 1
        invalid_filter_spans.append((start, end, str(node["invalid_ref"])))

    parser.StartElementHandler = start_element
    parser.EndElementHandler = end_element
    parser.ExternalEntityRefHandler = lambda *_args: 0
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.Parse(data, True)
    for edit_start, edit_end, _replacement, old_ref, _new_ref in sort_ref_edits:
        for filter_start, filter_end, _invalid_ref in invalid_filter_spans:
            if filter_start <= edit_start < edit_end <= filter_end:
                raise ValueError(
                    f"工作表“{sheet_name}”的 sortState ref {old_ref!r} 位于将被临时移除的无效 autoFilter 内；"
                    "为保护排序信息，已拒绝该文件且未删除排序定义。"
                )
    edits: list[tuple[int, int, bytes, str | None]] = [
        (start, end, b"", reference) for start, end, reference in invalid_filter_spans
    ]
    edits.extend((start, end, replacement, None) for start, end, replacement, _old, _new in sort_ref_edits)
    if not edits:
        return data, [], []
    edits.sort(key=lambda item: item[0])
    output: list[bytes] = []
    invalid_filter_refs: list[str] = []
    position = 0
    for start, end, replacement, reference in edits:
        if start < position or end < start or end > len(data):
            raise ValueError("工作表 XML 修改区域重叠或越界；未生成修改副本。")
        output.append(data[position:start])
        output.append(replacement)
        position = end
        if reference is not None:
            shown_reference = reference if len(reference) <= 200 else reference[:197] + "..."
            invalid_filter_refs.append(shown_reference)
    output.append(data[position:])
    return b"".join(output), invalid_filter_refs, normalized_refs


class _TemporaryWorkbookCopy:
    """Path-like temporary copy that owns every openpyxl workbook loaded from it."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.workbooks = []

    def __fspath__(self) -> str:
        return os.fspath(self.path)

    def load_workbook(self, *args, **kwargs):
        workbook = load_workbook(self.path, *args, **kwargs)
        self.workbooks.append(workbook)
        return workbook

    def close_workbooks(self) -> list[str]:
        errors = []
        for workbook in reversed(self.workbooks):
            try:
                workbook.close()
            except Exception as exc:
                errors.append(str(exc))
        self.workbooks.clear()
        return errors


_TEMP_COPY_CLEANUP_RETRY_DELAYS = (0.05, 0.1, 0.2)


def _remove_temporary_workbook_copy(path: Path) -> OSError | None:
    for attempt in range(len(_TEMP_COPY_CLEANUP_RETRY_DELAYS) + 1):
        try:
            path.unlink(missing_ok=True)
            return None
        except OSError as exc:
            transient_lock = isinstance(exc, PermissionError) or getattr(exc, "winerror", None) in (32, 33)
            if not transient_lock or attempt == len(_TEMP_COPY_CLEANUP_RETRY_DELAYS):
                return exc
            time.sleep(_TEMP_COPY_CLEANUP_RETRY_DELAYS[attempt])
    return None


def _log_temporary_copy_warning(log: Callable[[str], None], message: str) -> None:
    try:
        log(message)
    except Exception:
        print(message, file=sys.stderr)


@contextmanager
def _temporary_workbook_copy(path: Path, label: str, log: Callable[[str], None]):
    """Prepare a unique XLSX copy and close all tracked workbooks before deleting it."""
    try:
        fd, temp_name = tempfile.mkstemp(prefix="wps_sheet_autofilter_", suffix=".xlsx")
    except OSError as exc:
        raise ProcessingError(
            f"无法为{label}“{path.name}”创建安全的临时读取副本。请在 WPS 中打开该文件的副本，"
            f"移除或重建筛选后另存为 XLSX 再试；原文件未修改。\n原因：{exc}"
        ) from exc
    os.close(fd)
    temp_path = Path(temp_name)
    temporary_copy = _TemporaryWorkbookCopy(temp_path)
    operation_error = None
    try:
        try:
            with zipfile.ZipFile(path, "r") as source:
                infos = source.infolist()
                if len(infos) > MAX_XLSX_MEMBERS:
                    raise ValueError(f"ZIP 成员数超过 {MAX_XLSX_MEMBERS} 项安全上限。")
                total_declared = sum(info.file_size for info in infos)
                if total_declared > MAX_XLSX_UNCOMPRESSED_BYTES:
                    raise ValueError("ZIP 解压总大小超过 512 MiB 安全上限。")
                names: set[str] = set()
                for info in infos:
                    if not _zip_member_name_is_safe(info.filename):
                        raise ValueError(f"ZIP 成员路径不安全：{info.filename!r}")
                    if info.filename in names:
                        raise ValueError(f"ZIP 中存在重名成员：{info.filename}")
                    names.add(info.filename)
                    if info.flag_bits & 0x1:
                        raise ValueError(f"ZIP 成员已加密，无法安全读取：{info.filename}")
                sheet_part_names = _worksheet_parts(source, names)
                changed_filters: list[tuple[str, str]] = []
                changed_sort_states: list[tuple[str, str, str]] = []
                actual_total = 0
                with zipfile.ZipFile(temp_path, "w") as destination:
                    destination.comment = source.comment
                    for info in infos:
                        output_info = copy(info)
                        if info.is_dir():
                            destination.writestr(output_info, b"")
                            continue
                        if info.filename in sheet_part_names:
                            if info.file_size > MAX_XLSX_WORKSHEET_XML_BYTES:
                                raise ValueError(f"工作表 XML 部件 {info.filename} 超过 64 MiB 安全上限。")
                            raw = _read_bounded_zip_xml(source, info.filename, MAX_XLSX_WORKSHEET_XML_BYTES)
                            actual_total += len(raw)
                            cleaned, invalid_refs, normalized_refs = _preprocess_worksheet_xml(
                                raw, sheet_part_names[info.filename]
                            )
                            destination.writestr(output_info, cleaned)
                            for invalid_ref in invalid_refs:
                                changed_filters.append((sheet_part_names[info.filename], invalid_ref))
                            for old_ref, new_ref in normalized_refs:
                                changed_sort_states.append((sheet_part_names[info.filename], old_ref, new_ref))
                        else:
                            member_total = 0
                            with source.open(info, "r") as member_in, destination.open(output_info, "w") as member_out:
                                while True:
                                    chunk = member_in.read(1024 * 1024)
                                    if not chunk:
                                        break
                                    member_total += len(chunk)
                                    actual_total += len(chunk)
                                    if actual_total > MAX_XLSX_UNCOMPRESSED_BYTES or member_total > info.file_size:
                                        raise ValueError("ZIP 实际解压数据超过安全上限或与目录大小不符。")
                                    member_out.write(chunk)
                            if member_total != info.file_size:
                                raise ValueError(f"ZIP 成员 {info.filename} 的实际大小与目录记录不符。")
                        if actual_total > MAX_XLSX_UNCOMPRESSED_BYTES:
                            raise ValueError("ZIP 实际解压总大小超过 512 MiB 安全上限。")
            for sheet_name, invalid_ref in changed_filters:
                log(
                    f"{label}“{path.name}”工作表“{sheet_name}”的自动筛选区域 {invalid_ref!r} 无效（文件：{path}）；"
                    "仅在临时读取副本中移除此筛选定义，原文件未更改。"
                )
            for sheet_name, old_ref, new_ref in changed_sort_states:
                log(
                    f"{label}“{path.name}”工作表“{sheet_name}”的 sortState ref {old_ref!r} 已在临时读取副本中"
                    f"规范为 {new_ref!r}（文件：{path}）；原文件未更改。"
                )
        except ProcessingError:
            raise
        except Exception as exc:
            raise ProcessingError(
                f"无法安全检查或准备{label}“{path.name}”的临时副本。请在 WPS 中打开该文件的副本，"
                f"检查并修正工作表排序/筛选区域引用后另存为 XLSX 再试；不要直接删除排序定义，原文件未修改。\n原因：{exc}"
            ) from exc
        yield temporary_copy
    except BaseException as exc:
        operation_error = exc
        raise
    finally:
        close_errors = temporary_copy.close_workbooks()
        if close_errors:
            _log_temporary_copy_warning(
                log,
                f"警告：{label}“{path.name}”临时读取副本中的 openpyxl 工作簿关闭时发生错误："
                f"{temp_path}；请确认该文件未被其他程序占用。原因：{'；'.join(close_errors)}",
            )
        cleanup_error = _remove_temporary_workbook_copy(temp_path)
        if cleanup_error is not None:
            state = "读取/处理已成功，但临时读取副本无法清理" if operation_error is None else "读取/处理本身已失败，且临时读取副本也无法清理"
            _log_temporary_copy_warning(
                log,
                f"警告：{state}：{temp_path}。请关闭占用该文件的程序后手动删除；原工作簿未修改。"
                f"清理原因：{cleanup_error}",
            )


def _find_external_formula(formula: str) -> bool:
    if re.search(r"\[[^\]]+\.(?:xlsx?|xlsm|xlsb)\]", formula, flags=re.IGNORECASE):
        return True
    return bool(re.search(r"\b(?:WEBSERVICE|FILTERXML|RTD|CALL|REGISTER\.ID|EXEC)\s*\(", formula, flags=re.IGNORECASE))


def _normalize_path(value: str, label: str) -> Path:
    raw = str(value or "").strip()
    if not raw:
        raise ProcessingError(f"请先选择{label}。")
    return Path(raw).expanduser().resolve()


def _first_missing_source_location(settings: dict[str, Any]) -> tuple[int, int] | None:
    """Return the zero-based job/source position of the first blank source path."""
    for job_idx, job in enumerate(settings.get("jobs", [])):
        for source_idx, source in enumerate(job.get("sources", [])):
            if not str(source.get("path") or "").strip() and not source.get("_optional_placeholder", False):
                return job_idx, source_idx
    return None


def _missing_source_path_message(settings: dict[str, Any], location: tuple[int, int]) -> str:
    job_idx, source_idx = location
    sheet_name = settings["jobs"][job_idx].get("template_sheet", f"第 {job_idx + 1} 个目标表")
    return (f"目标表“{sheet_name}”的第 {source_idx + 1} 份源工作簿尚未配置文件路径。"
            "请在源文件列表中编辑该项并选择或填写 .xlsx 路径。")


def _get_formula_templates(workbook, job: dict[str, Any], index: int, start_row: int) -> tuple[int, int, dict[int, str]]:
    sheet_name = str(job["template_sheet"]).strip()
    if sheet_name not in workbook.sheetnames:
        raise ProcessingError(f"模板中找不到第 {index} 个目标工作表“{sheet_name}”。")
    ws = workbook[sheet_name]
    formula_row = _formula_template_row({**job, "start_row": start_row})
    col_start, col_end = parse_column_range(job["formula_range"])
    formulas: dict[int, str] = {}
    missing: list[str] = []
    for col in range(col_start, col_end + 1):
        value = ws.cell(row=formula_row, column=col).value
        if not (isinstance(value, str) and value.startswith("=")):
            missing.append(get_column_letter(col))
            continue
        if _find_external_formula(value):
            raise ProcessingError(f"模板公式 {sheet_name}!{get_column_letter(col)}{formula_row} 包含外部引用或外部访问函数，已拒绝重算。")
        formulas[col] = value
    if missing:
        columns = "、".join(missing)
        raise ProcessingError(
            f"模板“{sheet_name}”的公式模板行是第 {formula_row} 行（由粘贴起始行 {start_row} 自动确定），"
            f"以下公式列缺少公式：{columns}。请补齐该模板行中的公式。"
        )
    return col_start, col_end, formulas


def _cell_has_data(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or value.strip() != "")


def _last_nonempty_source_row(ws) -> tuple[int, int]:
    """Find the last row with content and the sheet's used-column boundary."""
    max_row = max(ws.max_row or 0, 1)
    max_col = max(ws.max_column or 0, 1)
    last_nonempty = 0
    for row_number, cells in enumerate(
        ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col), start=1
    ):
        if any(_cell_has_data(cell.value) for cell in cells):
            last_nonempty = row_number
    return last_nonempty, max_col


def _read_source_records(path: Path, sheet_name: str, start_row: int, col_start: int, col_end: int,
                         log: Callable[[str], None], end_row: int | None = None,
                         source_processing: dict[str, Any] | None = None,
                         source_rules_used: set[tuple[str, int]] | None = None) -> list[list[Any]]:
    _check_xlsx_file(path, f"源工作簿“{path.name}”")
    with _temporary_workbook_copy(path, "源工作簿", log) as load_path:
        return _read_source_records_from_copy(load_path, path, sheet_name, start_row, col_start, col_end,
                                              log, end_row, source_processing, source_rules_used)


def _read_source_records_from_copy(load_path: _TemporaryWorkbookCopy, path: Path, sheet_name: str, start_row: int,
                                   col_start: int, col_end: int, log: Callable[[str], None],
                                   end_row: int | None = None,
                                   source_processing: dict[str, Any] | None = None,
                                   source_rules_used: set[tuple[str, int]] | None = None) -> list[list[Any]]:
    try:
        values_book = load_path.load_workbook(read_only=True, data_only=True, keep_links=False)
        formula_book = load_path.load_workbook(read_only=True, data_only=False, keep_links=False)
        requested_sheet = str(sheet_name or "").strip()
        if not values_book.sheetnames or not formula_book.sheetnames:
            raise ProcessingError(f"源工作簿“{path.name}”没有可读取的工作表。")
        resolved_sheet = _resolve_sheet_name(formula_book.sheetnames, requested_sheet, f"源工作簿“{path.name}”")
        if not requested_sheet:
            log(f"{path.name} 未填写源工作表，默认使用第一个工作表“{resolved_sheet}”。")
        if resolved_sheet not in values_book.sheetnames:
            raise ProcessingError(f"源工作簿“{path.name}”中找不到工作表“{resolved_sheet}”。")
        val_ws, raw_ws = values_book[resolved_sheet], formula_book[resolved_sheet]
        actual_last_row, used_max_col = _last_nonempty_source_row(raw_ws)
        if end_row is not None and end_row < start_row:
            raise ProcessingError(f"源工作簿“{path.name}”的结束行早于复制起始行。")
        if actual_last_row < start_row:
            raise ProcessingError(f"源工作簿“{path.name}”工作表“{resolved_sheet}”在第 {start_row} 行及之后没有可复制数据。")
        if end_row is not None:
            last_row = min(end_row, actual_last_row)
            log(f"{path.name} 使用显式复制结束行 {end_row}；本次读取上限为第 {last_row} 行。")
        else:
            last_row = actual_last_row
            boundary_row = None
            for row_number, cells in enumerate(
                raw_ws.iter_rows(min_row=start_row, max_row=actual_last_row, min_col=1, max_col=used_max_col),
                start=start_row,
            ):
                if not any(_cell_has_data(cell.value) for cell in cells):
                    boundary_row = row_number
                    last_row = row_number - 1
                    break
            if boundary_row is not None:
                log(f"{path.name} 在第 {boundary_row} 行遇到整行空白边界（检查全部 {used_max_col} 个已使用列）；采用第 {last_row} 行为末行，边界后的数据不纳入此源。")
            else:
                log(f"{path.name} 从第 {start_row} 行至最后非空行 {actual_last_row} 未遇到整行空白边界；采用最后非空行作为末行。")
        if last_row < start_row:
            raise ProcessingError(f"源工作簿“{path.name}”工作表“{resolved_sheet}”从第 {start_row} 行起遇到空白边界，没有可复制数据。")

        excluded_rows: set[int] = set()
        overlay: dict[tuple[int, int], Any] = {}
        if source_processing is not None:
            overlay, excluded_rows, used = _apply_source_region_processing(
                raw_ws, val_ws, resolved_sheet, source_processing, path, log, actual_last_row
            )
            if source_rules_used is not None:
                source_rules_used.update(used)

        records: list[list[Any]] = []
        blank_formula_cache_count = 0
        raw_rows = raw_ws.iter_rows(min_row=start_row, max_row=last_row, min_col=col_start, max_col=col_end)
        value_rows = val_ws.iter_rows(min_row=start_row, max_row=last_row, min_col=col_start, max_col=col_end)
        for row_number, (raw_cells, value_cells) in enumerate(zip(raw_rows, value_rows), start=start_row):
            if row_number in excluded_rows:
                continue
            raw = [cell.value for cell in raw_cells]
            vals = [cell.value for cell in value_cells]
            for col_offset in range(col_end - col_start + 1):
                key = (row_number, col_start + col_offset)
                if key in overlay:
                    vals[col_offset] = overlay[key]
            if not any(_cell_has_data(value) for value in raw) and not any(_cell_has_data(value) for value in vals):
                continue
            if not any(_cell_has_data(value) for value in raw):
                # Newly split output values still make this a copyable row.
                raw = vals
            for source_value, cached_value in zip(raw, vals):
                if isinstance(source_value, str) and source_value.startswith("=") and cached_value is None:
                    blank_formula_cache_count += 1
            records.append(vals)
        if not records:
            interval = f"第 {start_row} 至 {last_row} 行" if end_row is not None else f"第 {start_row} 行起至第 {last_row} 行"
            raise ProcessingError(f"源工作簿“{path.name}”工作表“{resolved_sheet}”的{interval}在复制列范围内没有数据。")
        if blank_formula_cache_count:
            log(f"提示：{path.name} 有 {blank_formula_cache_count} 个源公式单元格没有已保存的计算缓存；这些单元格将按空值复制，不会执行源公式。")
        return records
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"读取源工作簿“{path.name}”失败：{exc}") from exc


def _apply_source_region_processing(raw_ws, val_ws, sheet_name: str, settings: dict[str, Any],
                                    path: Path, log: Callable[[str], None], used_last_row: int | None = None):
    """Apply source-scoped rules to an in-memory overlay; the original XLSX is never written."""
    overlay: dict[tuple[int, int], Any] = {}
    excluded_rows: set[int] = set()
    used_rules: set[tuple[str, int]] = set()

    def matrices(bounds):
        min_col, min_row, max_col, max_row = bounds
        raw = [[cell.value for cell in row] for row in raw_ws.iter_rows(
            min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)]
        values = [[cell.value for cell in row] for row in val_ws.iter_rows(
            min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)]
        for ri, row in enumerate(values):
            row_number = min_row + ri
            if row_number in excluded_rows:
                values[ri] = [None] * len(row)
                raw[ri] = [None] * len(raw[ri])
                continue
            for ci, value in enumerate(row):
                key = (row_number, min_col + ci)
                if key in overlay:
                    row[ci] = overlay[key]
        return raw, values

    for operation_key in settings.get("workflow", DEFAULT_WORKFLOW):
        if operation_key not in ("TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
            continue
        operation = settings.get(operation_key, {})
        if not operation.get("enabled", False):
            continue
        label = MODULE_LABELS[operation_key]
        for rule_idx, rule in enumerate(operation.get("rules", []), 1):
            if str(rule.get("scope", "target")).strip().lower() != "source":
                continue
            try:
                rule_sheet = _resolve_sheet_name(
                    raw_ws.parent.sheetnames, rule.get("sheet", ""), f"源工作簿“{path.name}”"
                )
            except ProcessingError:
                continue
            if rule_sheet.casefold() != sheet_name.casefold():
                continue
            used_rules.add((operation_key, rule_idx - 1))
            if operation_key == "RegionFilter":
                bounds, selection_kind = _parse_region_filter_area(rule["range"], f"第 {rule_idx} 条{label}")
                if selection_kind == "sheet":
                    last_row, max_col = _last_nonempty_source_row(raw_ws)
                    bounds = (1, 1, max(1, max_col), max(1, last_row))
                elif selection_kind == "column":
                    last_row = used_last_row or _last_nonempty_source_row(raw_ws)[0] or 1
                    bounds = (bounds[0], 1, bounds[2], last_row)
            else:
                bounds = _parse_selected_area(rule["range"], f"第 {rule_idx} 条{label}")
            _check_processing_area_size(bounds, f"第 {rule_idx} 条{label}")
            raw_matrix, matrix = matrices(bounds)
            if operation_key == "TextToColumns":
                values = [row[0] for row in matrix]
                formats = [str(value).lower() for value in rule.get("column_formats", ["text"])]
                split_matrix, result_columns = split_text_column(values, str(rule["delimiter"]), formats)
                start_col = column_number(rule["output_start_col"])
                target_bounds = (start_col, bounds[1], start_col + result_columns - 1, bounds[3])
                if _areas_overlap(bounds, target_bounds) or target_bounds[2] > 16_384:
                    raise ProcessingError(f"第 {rule_idx} 条源文件文本分列规则的输出区域无效或与输入重叠。")
                _check_processing_area_size(target_bounds, f"第 {rule_idx} 条{label}目标")
                target_raw, target_values = matrices(target_bounds)
                for ri, row in enumerate(target_values):
                    source_row = target_bounds[1] + ri
                    if source_row in excluded_rows:
                        continue
                    for ci, value in enumerate(row):
                        if _cell_has_data(value) or (isinstance(target_raw[ri][ci], str) and target_raw[ri][ci].startswith("=")):
                            raise ProcessingError(f"第 {rule_idx} 条源文件文本分列规则的目标单元格 {get_column_letter(start_col + ci)}{source_row} 已有内容；不会覆盖。")
                for ri, row in enumerate(split_matrix):
                    source_row = bounds[1] + ri
                    if source_row not in excluded_rows:
                        for ci, value in enumerate(row):
                            overlay[(source_row, start_col + ci)] = value
                log(f"已在源文件 {path.name} 的内存数据中执行文本分列：{sheet_name}!{_area_address(bounds)} → {_area_address(target_bounds)}。原文件未修改。")
            elif operation_key == "RegionFilter":
                filter_col = column_number(rule["filter_column"])
                if any(isinstance(value, str) and value.startswith("=") for row in raw_matrix for value in row):
                    raise ProcessingError(f"第 {rule_idx} 条源文件区域筛选规则含公式单元格；为避免公式缓存不一致，已停止。")
                _filtered, removed, keep_indexes = filter_matrix(
                    matrix, bool(rule.get("has_header", False)), filter_col - bounds[0],
                    str(rule.get("operator", "equals")), str(rule.get("criteria", "")))
                for offset in range(len(matrix)):
                    row_number = bounds[1] + offset
                    if row_number in excluded_rows or offset not in keep_indexes:
                        if not (rule.get("has_header") and offset == 0):
                            excluded_rows.add(row_number)
                log(f"已在源文件 {path.name} 的内存数据中执行区域筛选：{sheet_name}!{_area_address(bounds)}，移除不匹配记录 {removed} 行；原文件未修改。")
            elif operation_key == "Deduplicate":
                if any(isinstance(value, str) and value.startswith("=") for row in raw_matrix for value in row):
                    raise ProcessingError(f"第 {rule_idx} 条源文件去重规则含公式单元格；已停止。")
                mode = str(rule.get("key_mode", "row")).lower()
                key_indexes = None if mode == "row" else [column_number(value) - bounds[0] for value in rule["key_columns"]]
                keep_indexes, removed = _deduplicate_keep_indexes(matrix, bool(rule.get("has_header", False)), mode, key_indexes)
                for offset in range(len(matrix)):
                    row_number = bounds[1] + offset
                    if offset not in keep_indexes and not (rule.get("has_header") and offset == 0):
                        excluded_rows.add(row_number)
                log(f"已在源文件 {path.name} 的内存数据中执行去重：{sheet_name}!{_area_address(bounds)}，移除 {removed} 行；原文件未修改。")
            else:
                if any(isinstance(value, str) and value.startswith("=") for row in raw_matrix for value in row):
                    raise ProcessingError(f"第 {rule_idx} 条源文件格式转换规则含公式单元格；已停止。")
                mode = str(rule["mode"]).lower()
                for ri, row in enumerate(matrix):
                    row_number = bounds[1] + ri
                    if row_number in excluded_rows:
                        continue
                    try:
                        converted = [convert_cell_value(value, mode) for value in row]
                    except ProcessingError as exc:
                        raise ProcessingError(f"源文件 {path.name} 第 {rule_idx} 条格式转换规则，第 {ri + 1} 行：{exc}") from exc
                    for ci, value in enumerate(converted):
                        overlay[(row_number, bounds[0] + ci)] = value
                log(f"已在源文件 {path.name} 的内存数据中执行格式转换：{sheet_name}!{_area_address(bounds)} → {mode}；原文件未修改。")
    return overlay, excluded_rows, used_rules

def _validate_inputs(raw_settings: dict[str, Any]) -> tuple[Path, Path]:
    settings = normalize_settings(raw_settings)
    validate_static_settings(settings)
    missing_source = _first_missing_source_location(settings)
    if missing_source is not None:
        raise ProcessingError(_missing_source_path_message(settings, missing_source))
    template_path = _normalize_path(settings.get("template_path", ""), "模板工作簿")
    _check_xlsx_file(template_path, "模板工作簿")
    all_inputs = {template_path}
    for job_idx, job in enumerate(settings["jobs"], 1):
        for source_idx, source in enumerate(job["sources"], 1):
            path = _normalize_path(source.get("path", ""), f"目标表“{job['template_sheet']}”第 {source_idx} 份源工作簿")
            # Blank source_sheet is intentionally resolved by the reader to the first sheet.
            if path == template_path:
                raise ProcessingError("模板工作簿不能同时作为源工作簿。")
            _check_xlsx_file(path, f"目标表“{job['template_sheet']}”第 {source_idx} 份源工作簿")
            source["path"] = str(path)
            all_inputs.add(path)
    output_dir = _normalize_path(settings.get("output_dir", ""), "输出目录")
    if not output_dir.is_dir():
        raise ProcessingError(f"输出目录不存在：\n{output_dir}")
    filename = str(settings.get("output_filename", "")).strip()
    if not filename:
        raise ProcessingError("输出文件名不能为空。")
    if Path(filename).name != filename or "/" in filename or "\\" in filename:
        raise ProcessingError("输出文件名只能填写文件名，不能包含目录路径。")
    if Path(filename).suffix and Path(filename).suffix.lower() != ".xlsx":
        raise ProcessingError("输出文件只支持 .xlsx 扩展名；请勿使用 .xlsm 或其他格式。")
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    final_path = (output_dir / filename).resolve()
    if os.path.normcase(str(final_path)) == os.path.normcase(str(template_path)):
        raise ProcessingError(
            f"输出文件名与模板同名且位于同一目录，不能覆盖模板：\n{template_path}\n"
            "请改名或选择其他输出目录。"
        )
    if final_path.exists():
        raise ProcessingError(f"输出文件已存在，为防止覆盖任何文件已停止：\n{final_path}\n请更改输出文件名或选择其他输出目录。")
    if final_path in all_inputs:
        raise ProcessingError("输出路径不能与模板或任何源工作簿相同。")
    if any(job.get("sources") for job in settings["jobs"]):
        column_number(str(settings.get("copy_start_col", "A")).strip() or "A")
        if str(settings.get("copy_end_col", "")).strip():
            column_number(settings["copy_end_col"])
        else:
            raise ProcessingError("请填写数据复制末列；程序不会猜测需要复制的源列范围。仅模板编辑时无需设置复制列范围。")
    return template_path, final_path


def _validate_template_area_sheets(workbook, settings: dict[str, Any]) -> None:
    for key, label in (("pre_mappings", "前映射"), ("post_mappings", "后映射")):
        for idx, rule in enumerate(settings.get(key, []), 1):
            for field in ("source_sheet", "target_sheet"):
                sheet_name = str(rule[field]).strip()
                if sheet_name not in workbook.sheetnames:
                    raise ProcessingError(f"第 {idx} 条{label}规则引用了模板中不存在的工作表“{sheet_name}”。")
    for idx, output in enumerate(settings.get("region_xlsx_outputs", []), 1):
        sheet_name = str(output["sheet"]).strip()
        if sheet_name not in workbook.sheetnames:
            raise ProcessingError(f"第 {idx} 项区域 XLSX 输出引用了模板中不存在的工作表“{sheet_name}”。")
    for idx, output in enumerate(settings.get("text_outputs", []), 1):
        sheet_name = str(output["sheet"]).strip()
        if sheet_name not in workbook.sheetnames:
            raise ProcessingError(f"第 {idx} 项文本输出配置引用了模板中不存在的工作表“{sheet_name}”。")
    for key, label in (("CustomFormula", "公式填充"), ("TextToColumns", "文本分列"), ("RegionFilter", "区域筛选"),
                       ("Deduplicate", "去重"), ("FormatConvert", "格式转换")):
        for idx, rule in enumerate(settings.get(key, {}).get("rules", []), 1):
            if str(rule.get("scope", "target")).strip().lower() != "target":
                continue
            sheet_name = str(rule.get("sheet", "") or "").strip()
            _resolve_sheet_name(workbook.sheetnames, sheet_name, f"第 {idx} 条{label}模板规则")


def _capture_pre_mapping_values(workbook, settings: dict[str, Any]):
    """Read all pre-mapping sources from the untouched template, before any target cleanup."""
    _validate_template_area_sheets(workbook, settings)
    captured = []
    for idx, rule in enumerate(settings.get("pre_mappings", []), 1):
        bounds = _parse_area(rule["source_range"], f"第 {idx} 条前映射源")
        min_col, min_row, max_col, max_row = bounds
        worksheet = workbook[str(rule["source_sheet"]).strip()]
        values = [[worksheet.cell(row=row, column=col).value
                   for col in range(min_col, max_col + 1)]
                  for row in range(min_row, max_row + 1)]
        captured.append((deepcopy(rule), bounds, values))
    return captured


def _apply_pre_mapping_values(workbook, captured, log: Callable[[str], None]) -> None:
    for idx, (rule, source_bounds, values) in enumerate(captured, 1):
        target_bounds = _parse_area(rule["target_range"], f"第 {idx} 条前映射目标")
        source_min_col, source_min_row, _source_max_col, _source_max_row = source_bounds
        target_min_col, target_min_row, _target_max_col, _target_max_row = target_bounds
        source_sheet = workbook[str(rule["source_sheet"]).strip()]
        target_sheet = workbook[str(rule["target_sheet"]).strip()]
        for row_offset, row_values in enumerate(values):
            for col_offset, value in enumerate(row_values):
                target = target_sheet.cell(row=target_min_row + row_offset, column=target_min_col + col_offset)
                if isinstance(target, MergedCell):
                    raise ProcessingError(f"前映射目标 {rule['target_sheet']}!{target.coordinate} 是合并单元格的非左上角，不能写入。")
                if isinstance(value, str) and value.startswith("="):
                    source_coord = f"{get_column_letter(source_min_col + col_offset)}{source_min_row + row_offset}"
                    try:
                        value = Translator(value, origin=source_coord).translate_formula(target.coordinate)
                    except Exception as exc:
                        raise ProcessingError(f"无法将前映射公式 {rule['source_sheet']}!{source_coord} 平移到 {rule['target_sheet']}!{target.coordinate}：{exc}") from exc
                target.value = value
        log(f"已应用前映射 {rule['source_sheet']}!{_area_address(source_bounds)} → {rule['target_sheet']}!{_area_address(target_bounds)}。")


def _validate_pre_mapping_conflicts(settings: dict[str, Any], source_records: list[list[list[Any]]]) -> None:
    """Reject pre-map targets that overlap main-data, formula, result-map or post-map writes."""
    copy_start = column_number(str(settings.get("copy_start_col", "A")).strip() or "A")
    copy_end_value = str(settings.get("copy_end_col", "")).strip()
    copy_end = column_number(copy_end_value) if copy_end_value else copy_start
    protected: dict[str, list[tuple[str, tuple[int, int, int, int]]]] = {}
    for job, records in zip(settings["jobs"], source_records):
        sheet = str(job["template_sheet"]).strip()
        template_row = _formula_template_row(job)
        protected.setdefault(sheet, []).append(("公式模板", (1, template_row, 16_384, template_row)))
        if records:
            start_row = int(job["start_row"])
            end_row = start_row + len(records) - 1
            protected[sheet].append(("主数据复制", (copy_start, start_row, copy_end, end_row)))
            formula_start, formula_end = parse_column_range(job["formula_range"])
            protected[sheet].append(("公式填充", (formula_start, start_row, formula_end, end_row)))
            for mapping in job.get("mappings", []):
                col = column_number(mapping["target_col"])
                protected[sheet].append(("结果映射", (col, start_row, col, end_row)))
    for rule in settings.get("post_mappings", []):
        bounds = _parse_area(rule["target_range"], "后映射目标")
        protected.setdefault(str(rule["target_sheet"]).strip(), []).append(("后映射", bounds))

    for idx, rule in enumerate(settings.get("pre_mappings", []), 1):
        target_sheet = str(rule["target_sheet"]).strip()
        target_bounds = _parse_area(rule["target_range"], f"第 {idx} 条前映射目标")
        for write_kind, write_bounds in protected.get(target_sheet, []):
            if _areas_overlap(target_bounds, write_bounds):
                raise ProcessingError(
                    f"第 {idx} 条前映射目标 {target_sheet}!{_area_address(target_bounds)} 与{write_kind}区域 "
                    f"{_area_address(write_bounds)} 冲突；已在修改临时工作簿前停止，避免覆盖。请调整前映射目标区域。"
                )


def _create_staging_workbook(template_path: Path, raw_settings: dict[str, Any], source_records: list[list[list[Any]]],
                             staging_path: Path, log: Callable[[str], None]) -> list[tuple[str, int, int, list[dict[str, str]]]]:
    _check_xlsx_file(template_path, "模板工作簿")
    with _temporary_workbook_copy(template_path, "模板工作簿", log) as load_path:
        return _create_staging_workbook_from_copy(template_path, load_path, raw_settings, source_records, staging_path, log)


def _create_staging_workbook_from_copy(template_path: Path, load_path: _TemporaryWorkbookCopy, raw_settings: dict[str, Any],
                                       source_records: list[list[list[Any]]], staging_path: Path,
                                       log: Callable[[str], None]) -> list[tuple[str, int, int, list[dict[str, str]]]]:
    settings = normalize_settings(raw_settings)
    try:
        wb = load_path.load_workbook(keep_links=False, data_only=False)
        captured_pre_mappings = _capture_pre_mapping_values(wb, settings)
        for ws in wb.worksheets:
            for cell in ws._cells.values():
                if isinstance(cell.value, str) and cell.value.startswith("=") and _find_external_formula(cell.value):
                    raise ProcessingError(f"模板“{ws.title}”单元格 {cell.coordinate} 含外部引用或外部访问函数；为避免外部访问已停止。")
        copy_start = column_number(str(settings.get("copy_start_col", "A")).strip() or "A")
        if len(source_records) != len(settings["jobs"]):
            raise ProcessingError("内部数据与目标工作表数量不一致。")
        job_ranges = []
        for idx, (job, records) in enumerate(zip(settings["jobs"], source_records), 1):
            sheet_name = str(job["template_sheet"]).strip()
            if sheet_name not in wb.sheetnames:
                raise ProcessingError(f"模板中找不到第 {idx} 个目标工作表“{sheet_name}”。")
            ws = wb[sheet_name]
            start_row = int(job["start_row"])
            header_rows = int(job["header_rows"])
            if not records:
                log(f"目标表“{sheet_name}”未关联源文件；保留模板现有内容，仅执行配置的模板规则。")
                continue
            _c1, _c2, formulas = _get_formula_templates(wb, job, idx, start_row)
            end_row = start_row + len(records) - 1
            if end_row > 1_048_576:
                raise ProcessingError(f"第 {idx} 个工作表目标末行超出 Excel 最大行数。")
            job_ranges.append((job, ws, start_row, end_row, header_rows, formulas, records))

        # Cleanup is deliberately a separate pass so every pre-map source was captured first.
        for job, ws, start_row, _end_row, header_rows, formulas, _records in job_ranges:
            template_row = _formula_template_row(job)
            cleared = 0
            for cell in list(ws._cells.values()):
                if isinstance(cell, MergedCell) or cell.row <= header_rows:
                    continue
                if cell.row == template_row:
                    continue
                if cell.value is not None or cell.hyperlink is not None or cell.comment is not None:
                    cell.value = None
                    cell.hyperlink = None
                    cell.comment = None
                    cleared += 1
            log(f"已清理“{job['template_sheet']}”第 {header_rows + 1} 行起的旧内容（{cleared} 个单元格），保留第 {template_row} 行公式模板及其行内容。")

        _apply_pre_mapping_values(wb, captured_pre_mappings, log)

        formula_jobs: list[tuple[str, int, int, list[dict[str, str]]]] = []
        for job, ws, start_row, end_row, _header_rows, formulas, records in job_ranges:
            for offset, values in enumerate(records):
                target_row = start_row + offset
                for col_offset, value in enumerate(values):
                    ws.cell(row=target_row, column=copy_start + col_offset).value = value
            for col, formula in formulas.items():
                origin = f"{get_column_letter(col)}{_formula_template_row(job)}"
                for target_row in range(start_row, end_row + 1):
                    target = f"{get_column_letter(col)}{target_row}"
                    try:
                        ws[target] = Translator(formula, origin=origin).translate_formula(target)
                    except Exception as exc:
                        raise ProcessingError(f"无法将模板公式 {job['template_sheet']}!{origin} 填充到 {target}：{exc}") from exc
            mappings = [{"result_col": str(mapping["result_col"]).strip().upper(),
                         "target_col": str(mapping["target_col"]).strip().upper(),
                         "write_as": str(mapping.get("write_as", "value")).strip().lower()}
                        for mapping in job.get("mappings", [])]
            formula_jobs.append((str(job["template_sheet"]).strip(), start_row, end_row, mappings))
            log(f"已准备“{job['template_sheet']}”：按配置顺序追加 {len(records)} 行，填充第 {start_row} 至 {end_row} 行公式。")
        wb.save(staging_path)
        return formula_jobs
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"创建中间工作簿失败：{exc}") from exc


def _com_value_matrix(value, rows: int, cols: int) -> list[list[Any]]:
    if rows == 1 and cols == 1:
        if isinstance(value, (tuple, list)):
            first = value[0]
            return [list(first) if isinstance(first, (tuple, list)) else [first]]
        return [[value]]
    if rows == 1:
        if value and isinstance(value[0], (tuple, list)):
            return [list(value[0])]
        return [list(value)]
    if cols == 1:
        return [[row[0]] if isinstance(row, (tuple, list)) else [row] for row in value]
    return [list(row) for row in value]


def _write_mapped_values(destination, values: list[Any], write_as: str) -> None:
    """Write one mapped result column; text mode preserves values as text."""
    if str(write_as).strip().lower() == "text":
        values = [None if value is None else str(value) for value in values]
        try:
            destination.NumberFormat = "@"
        except Exception:
            pass
    destination.Value2 = tuple((value,) for value in values)


def _recalculate_and_wait(app, log: Callable[[str], None], phase: str) -> None:
    calculate_errors = []
    for method_name in ("CalculateFullRebuild", "CalculateFull", "Calculate"):
        try:
            getattr(app, method_name)()
            calculate_errors = []
            log(f"WPS 已执行 {method_name}（{phase}）。")
            break
        except Exception as exc:
            calculate_errors.append(f"{method_name}: {exc}")
    if calculate_errors:
        raise ProcessingError("WPS 无法触发公式重算：\n" + "\n".join(calculate_errors))
    deadline = time.time() + 180
    while time.time() < deadline:
        try:
            if int(app.CalculationState) == 0:
                return
        except Exception:
            return
        time.sleep(0.25)
    raise ProcessingError("等待 WPS 完成公式重算超过 180 秒，已中止以避免输出不完整文件。")


def _read_com_area(com_range, bounds: tuple[int, int, int, int]) -> list[list[Any]]:
    rows = bounds[3] - bounds[1] + 1
    cols = bounds[2] - bounds[0] + 1
    return _com_value_matrix(com_range.Value2, rows, cols)


def _write_com_area(com_range, matrix: list[list[Any]]) -> None:
    com_range.Value2 = tuple(tuple(row) for row in matrix)


def _flatten_com_values(value):
    if isinstance(value, (tuple, list)):
        for item in value:
            yield from _flatten_com_values(item)
    else:
        yield value


def _assert_no_formulas(com_range, label: str) -> None:
    try:
        formulas = com_range.Formula
    except Exception as exc:
        raise ProcessingError(f"无法检查{label}是否包含公式；为避免覆盖公式，已拒绝此区域处理。") from exc
    if any(isinstance(value, str) and value.startswith("=") for value in _flatten_com_values(formulas)):
        raise ProcessingError(f"{label}包含公式单元格；为避免删除公式，已拒绝此区域处理。")


def _assert_target_area_empty(com_range, bounds: tuple[int, int, int, int], label: str) -> None:
    _assert_no_formulas(com_range, label)
    matrix = _read_com_area(com_range, bounds)
    for row_offset, row in enumerate(matrix):
        for col_offset, value in enumerate(row):
            if _cell_has_data(value):
                cell = f"{get_column_letter(bounds[0] + col_offset)}{bounds[1] + row_offset}"
                raise ProcessingError(f"{label}将写入的目标单元格 {cell} 已有数据；不会静默覆盖，请调整目标列或先清空该区域。")


def _excel_number_format(mode: str) -> str:
    return {"text": "@", "number": "0.###############", "date": "yyyy-mm-dd"}[mode]


def _apply_region_processing(workbook, settings: dict[str, Any], formula_jobs, log: Callable[[str], None],
                             application=None) -> bool:
    formula_areas: dict[str, list[tuple[int, int, int, int]]] = {}
    result_mapping_areas: dict[str, list[tuple[int, int, int, int]]] = {}
    custom_formula_areas: dict[str, list[tuple[int, int, int, int]]] = {}
    formula_jobs_by_sheet = {item[0]: item[1:3] for item in formula_jobs}
    for job in settings.get("jobs", []):
        sheet_name = str(job["template_sheet"]).strip()
        if sheet_name in formula_jobs_by_sheet:
            start_row, end_row = formula_jobs_by_sheet[sheet_name]
            if end_row >= start_row:
                col_start, col_end = parse_column_range(job["formula_range"])
                formula_areas.setdefault(sheet_name, []).append((col_start, start_row, col_end, end_row))
        template_row = _formula_template_row(job)
        formula_areas.setdefault(sheet_name, []).append((1, template_row, 16_384, template_row))
    for sheet_name, start_row, end_row, mappings in formula_jobs:
        if end_row < start_row:
            continue
        for mapping in mappings:
            col = column_number(mapping["target_col"])
            result_mapping_areas.setdefault(sheet_name, []).append((col, start_row, col, end_row))
    applied_any = False
    operation_labels = {"CustomFormula": "公式填充", "TextToColumns": "文本分列", "RegionFilter": "区域筛选",
                        "Deduplicate": "去重", "FormatConvert": "格式转换"}
    configured_order = settings.get("workflow", DEFAULT_WORKFLOW)
    for operation_key in (module_id for module_id in configured_order if module_id in operation_labels):
        label = operation_labels[operation_key]
        operation = settings.get(operation_key, {"enabled": False, "rules": []})
        if not operation.get("enabled", False):
            continue
        formula_fill_applied = False
        for rule_idx, rule in enumerate(operation.get("rules", []), 1):
            if str(rule.get("scope", "target")).strip().lower() != "target":
                continue
            requested_sheet_name = str(rule.get("sheet", "") or "").strip()
            worksheet = _resolve_com_worksheet(
                workbook, requested_sheet_name, f"第 {rule_idx} 条{label}", log
            )
            sheet_name = str(worksheet.Name)
            if operation_key == "RegionFilter":
                bounds, selection_kind = _parse_region_filter_area(rule["range"], f"第 {rule_idx} 条{label}")
            else:
                bounds, selection_kind = _parse_selection_details(rule["range"], f"第 {rule_idx} 条{label}")
            bounds = _resolve_com_selection_bounds(worksheet, bounds, selection_kind, f"第 {rule_idx} 条{label}")
            _check_processing_area_size(bounds, f"第 {rule_idx} 条{label}")
            if any(_areas_overlap(bounds, protected) for protected in result_mapping_areas.get(sheet_name, [])):
                raise ProcessingError(f"第 {rule_idx} 条{label}规则区域 {sheet_name}!{_area_address(bounds)} 与结果映射目标区域重叠；为避免覆盖映射结果，已停止。")
            protected_areas = custom_formula_areas if operation_key == "CustomFormula" else formula_areas
            if any(_areas_overlap(bounds, protected) for protected in protected_areas.get(sheet_name, [])):
                raise ProcessingError(f"第 {rule_idx} 条{label}规则区域 {sheet_name}!{_area_address(bounds)} 与已配置的公式区域重叠；为避免覆盖其他公式规则，已停止。")
            source_range = worksheet.Range(_area_address(bounds))
            if operation_key == "CustomFormula":
                try:
                    merge_state = source_range.MergeCells
                except Exception as exc:
                    raise ProcessingError(f"无法检查第 {rule_idx} 条公式填充区域是否含合并单元格；为安全起见已停止。") from exc
                if merge_state not in (False, 0):
                    raise ProcessingError(f"第 {rule_idx} 条公式填充区域 {sheet_name}!{_area_address(bounds)} 含合并单元格或混合合并区域，不支持填充。")
                formula_mode = str(rule.get("formula_mode", "custom")).strip().lower()
                if formula_mode == "custom":
                    formula = str(rule.get("formula", "")).strip()
                    if not formula.startswith("=") or _find_external_formula(formula):
                        raise ProcessingError(f"第 {rule_idx} 条自定义公式无效或包含外部访问；已拒绝填充。")
                    top_left_address = f"{get_column_letter(bounds[0])}{bounds[1]}"
                    top_left = worksheet.Range(top_left_address)
                    try:
                        top_left.Formula = formula
                        formula_r1c1 = top_left.FormulaR1C1
                    except Exception as exc:
                        raise ProcessingError(f"WPS 无法将第 {rule_idx} 条自定义公式写入 {sheet_name}!{top_left_address}：{exc}") from exc
                else:
                    seed_address = str(rule.get("seed_cell", "")).strip().upper()
                    seed_bounds = _parse_area(seed_address, f"第 {rule_idx} 条公式来源单元格")
                    seed_sheet_name = str(rule.get("_seed_sheet") or sheet_name).strip()
                    seed_worksheet = _resolve_com_worksheet(workbook, seed_sheet_name,
                                                            f"第 {rule_idx} 条公式来源")
                    seed_range = seed_worksheet.Range(_area_address(seed_bounds))
                    try:
                        seed_formula = seed_range.Formula
                        formula_r1c1 = seed_range.FormulaR1C1
                    except Exception as exc:
                        raise ProcessingError(f"无法读取第 {rule_idx} 条公式来源单元格 {seed_sheet_name}!{seed_address}：{exc}") from exc
                    if not isinstance(seed_formula, str) or not seed_formula.startswith("="):
                        raise ProcessingError(f"公式来源单元格 {seed_sheet_name}!{seed_address} 不含公式。")
                    if _find_external_formula(seed_formula):
                        raise ProcessingError(f"公式来源单元格 {seed_sheet_name}!{seed_address} 包含外部工作簿引用或外部访问函数，已拒绝。")
                if not isinstance(formula_r1c1, str) or not formula_r1c1.startswith("="):
                    raise ProcessingError(f"第 {rule_idx} 条公式填充规则无法取得有效的 WPS R1C1 公式。")
                try:
                    source_range.FormulaR1C1 = formula_r1c1
                except Exception as exc:
                    raise ProcessingError(f"WPS 无法将第 {rule_idx} 条公式按相对引用填充到 {sheet_name}!{_area_address(bounds)}：{exc}") from exc
                custom_formula_areas.setdefault(sheet_name, []).append(bounds)
                formula_areas.setdefault(sheet_name, []).append(bounds)
                mode_label = "自定义公式" if formula_mode == "custom" else f"来源单元格 {rule['seed_cell']}"
                log(f"已执行公式填充规则 {rule_idx}：{sheet_name}!{_area_address(bounds)}，使用{mode_label}并按相对引用向下/向右填充。")
                formula_fill_applied = True
            elif operation_key == "TextToColumns":
                values_matrix = _read_com_area(source_range, bounds)
                values = [row[0] for row in values_matrix]
                formats = [str(value).lower() for value in rule.get("column_formats", ["text"])]
                split_matrix, result_columns = split_text_column(values, str(rule["delimiter"]), formats)
                start_col = column_number(rule["output_start_col"])
                target_bounds = (start_col, bounds[1], start_col + result_columns - 1, bounds[3])
                if _areas_overlap(bounds, target_bounds):
                    raise ProcessingError(f"第 {rule_idx} 条文本分列规则的输出区域 {_area_address(target_bounds)} 与源区域 {_area_address(bounds)} 重叠；请更改输出起始列。")
                if target_bounds[2] > 16_384:
                    raise ProcessingError(f"第 {rule_idx} 条文本分列规则输出超出 XFD 列。")
                _check_processing_area_size(target_bounds, f"第 {rule_idx} 条文本分列目标")
                if any(_areas_overlap(target_bounds, protected) for protected in formula_areas.get(sheet_name, [])):
                    raise ProcessingError(f"第 {rule_idx} 条文本分列输出区域 {sheet_name}!{_area_address(target_bounds)} 会覆盖公式区域；已停止。")
                target_range = worksheet.Range(_area_address(target_bounds))
                _assert_target_area_empty(target_range, target_bounds, f"第 {rule_idx} 条文本分列")
                _write_com_area(target_range, split_matrix)
                for col_offset in range(result_columns):
                    col_letter = get_column_letter(start_col + col_offset)
                    worksheet.Range(f"{col_letter}{bounds[1]}:{col_letter}{bounds[3]}").NumberFormat = _excel_number_format(formats[col_offset])
                log(f"已执行文本分列规则 {rule_idx}：{sheet_name}!{_area_address(bounds)} → {sheet_name}!{_area_address(target_bounds)}，分隔符={rule['delimiter']!r}。")
            elif operation_key == "RegionFilter":
                filter_col = column_number(rule["filter_column"])
                if not bounds[0] <= filter_col <= bounds[2]:
                    raise ProcessingError(f"第 {rule_idx} 条区域筛选规则的条件列 {get_column_letter(filter_col)} 不在实际已用范围 {sheet_name}!{_area_address(bounds)} 内。")
                _assert_no_formulas(source_range, f"第 {rule_idx} 条区域筛选规则区域")
                matrix = _read_com_area(source_range, bounds)
                filtered, removed, _keep_indexes = filter_matrix(
                    matrix, bool(rule.get("has_header", False)),
                    filter_col - bounds[0],
                    str(rule.get("operator", "equals")), str(rule.get("criteria", "")))
                _write_com_area(source_range, filtered)
                log(f"已执行区域筛选规则 {rule_idx}：{sheet_name}!{_area_address(bounds)}，移除不匹配记录 {removed} 行；仅压缩所选区域。")
            elif operation_key == "Deduplicate":
                _assert_no_formulas(source_range, f"第 {rule_idx} 条去重规则区域")
                matrix = _read_com_area(source_range, bounds)
                key_mode = str(rule.get("key_mode", "row")).lower()
                key_indexes = None if key_mode == "row" else [column_number(value) - bounds[0] for value in rule["key_columns"]]
                if key_indexes is not None and any(index < 0 or index >= bounds[2] - bounds[0] + 1 for index in key_indexes):
                    raise ProcessingError(f"第 {rule_idx} 条去重规则的关键列不在实际已用范围 {sheet_name}!{_area_address(bounds)} 内。")
                result, removed = deduplicate_matrix(matrix, bool(rule.get("has_header", False)), key_mode, key_indexes)
                _write_com_area(source_range, result)
                log(f"已执行去重规则 {rule_idx}：{sheet_name}!{_area_address(bounds)}，按{'整行' if key_mode == 'row' else '关键列 ' + ','.join(rule['key_columns'])}判重，移除 {removed} 行，保留首次出现。")
            else:
                _assert_no_formulas(source_range, f"第 {rule_idx} 条格式转换规则区域")
                matrix = _read_com_area(source_range, bounds)
                mode = str(rule["mode"]).lower()
                converted = convert_matrix_values(matrix, mode)
                _write_com_area(source_range, converted)
                source_range.NumberFormat = _excel_number_format(mode)
                log(f"已执行格式转换规则 {rule_idx}：{sheet_name}!{_area_address(bounds)} → {mode}。")
            applied_any = True
        if formula_fill_applied:
            if application is None:
                raise ProcessingError("区域公式填充需要 WPS 应用程序执行完整重算；已停止。")
            _recalculate_and_wait(application, log, "区域公式填充后")
    return applied_any


def _pywin32_import_error_message(exc: ImportError) -> str:
    return (
        "无法导入 Windows COM 依赖 pywin32（pythoncom/win32com）。\n"
        f"当前 Python 解释器：{sys.executable}\n"
        "请从本程序目录双击 run_app.bat，它会使用本项目 .venv 中的 Python。"
        "不要直接双击 main.py，也不要用系统 Python 运行。\n"
        "如果 run_app.bat 仍提示缺少依赖，请先运行 install_and_build.bat；"
        "安装失败诊断保存在 install_and_build.log。\n"
        f"导入错误详情：{type(exc).__name__}: {exc}"
    )


def _connect_wps(log: Callable[[str], None]):
    if sys.platform != "win32":
        raise ProcessingError("WPS 自动重算只能在安装了 Windows 版 WPS 表格的 Windows 环境运行；当前环境不是 Windows，未生成输出文件。")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ProcessingError(_pywin32_import_error_message(exc)) from exc
    try:
        pythoncom.CoInitialize()
    except Exception as exc:
        raise ProcessingError(f"无法初始化 Windows COM，因而不能连接 WPS 表格：{type(exc).__name__}: {exc}") from exc
    progids = ("ket.Application", "ET.Application", "et.Application")
    failures = []
    for progid in progids:
        try:
            log(f"正在连接 WPS 表格 COM：{progid}")
            return win32com.client.DispatchEx(progid), pythoncom, progid
        except Exception as exc:
            failures.append(f"{progid}: {exc}")
    pythoncom.CoUninitialize()
    raise ProcessingError("无法连接 Windows 版 WPS 表格 COM/ProgID。已尝试 ket.Application、ET.Application、et.Application。\n"
                          "请确认 Windows 11 已安装并能正常启动 WPS 表格；修复/重新注册 WPS 表格自动化组件后重试。\n"
                          "本程序不会改用 Microsoft Excel，也不会在连接失败后留下最终输出文件。\n\n连接错误详情：\n" + "\n".join(failures))


def _configured_job_data_ranges(jobs: list[dict[str, Any]], source_records=None, formula_jobs=None):
    """Map configured template sheets to the rows containing this run's data."""
    formula_ranges = {
        str(sheet).strip(): (int(start_row), int(end_row))
        for sheet, start_row, end_row, _formulas in (formula_jobs or [])
    }
    data_ranges = {}
    for index, job in enumerate(jobs):
        sheet = str(job["template_sheet"]).strip()
        if source_records is not None:
            records = source_records[index] if index < len(source_records) else []
            start_row = int(job["start_row"])
            data_ranges[sheet] = (start_row, start_row + len(records) - 1) if records else None
        else:
            data_ranges[sheet] = formula_ranges.get(sheet)
    return data_ranges


def _resolve_post_mapping_areas(rule: dict[str, Any], job_data_ranges: dict[str, tuple[int, int] | None],
                                index: int) -> tuple[tuple[int, int, int, int], tuple[int, int, int, int]] | None:
    """Resolve whole-column post mappings to the applicable data rows when jobs define them."""
    source_bounds = _parse_area(rule["source_range"], f"第 {index} 条后映射源")
    target_bounds = _parse_area(rule["target_range"], f"第 {index} 条后映射目标")
    source_sheet = str(rule["source_sheet"]).strip()
    target_sheet = str(rule["target_sheet"]).strip()
    source_is_column = _is_full_column_area(source_bounds)
    target_is_column = _is_full_column_area(target_bounds)
    if not source_is_column and not target_is_column:
        return source_bounds, target_bounds

    source_known = source_sheet in job_data_ranges
    target_known = target_sheet in job_data_ranges
    source_rows = job_data_ranges.get(source_sheet)
    target_rows = job_data_ranges.get(target_sheet)

    # A configured target with no imported records has no applicable data rows.
    if target_is_column and target_known and target_rows is None:
        return None
    if source_is_column and source_known and source_rows is None:
        if target_is_column and target_known and target_rows is not None:
            raise ProcessingError(f"第 {index} 条后映射源工作表“{source_sheet}”没有可映射数据行，但目标“{target_sheet}”有数据。")
        return None

    known_ranges = [rows for rows in (source_rows if source_is_column else None,
                                      target_rows if target_is_column else None)
                    if rows is not None]
    if len(known_ranges) == 2:
        source_count = source_rows[1] - source_rows[0] + 1
        target_count = target_rows[1] - target_rows[0] + 1
        if source_count != target_count:
            raise ProcessingError(f"第 {index} 条后映射源/目标工作表的数据行数不一致，无法安全应用整列映射。")
    if not known_ranges:
        return source_bounds, target_bounds

    fallback_rows = known_ranges[0]
    if source_is_column:
        rows = source_rows if source_rows is not None else fallback_rows
        source_bounds = (source_bounds[0], rows[0], source_bounds[2], rows[1])
    if target_is_column:
        rows = target_rows if target_rows is not None else fallback_rows
        target_bounds = (target_bounds[0], rows[0], target_bounds[2], rows[1])

    source_shape = (source_bounds[2] - source_bounds[0], source_bounds[3] - source_bounds[1])
    target_shape = (target_bounds[2] - target_bounds[0], target_bounds[3] - target_bounds[1])
    if source_shape != target_shape:
        raise ProcessingError(f"第 {index} 条后映射源/目标区域解析为实际数据行后行列数不一致。")
    return source_bounds, target_bounds


def _calculate_in_wps(staging_path: Path, formula_jobs: list[tuple[str, int, int, list[dict[str, str]]]],
                      log: Callable[[str], None], post_mappings: list[dict[str, Any]] | None = None,
                      region_processing: dict[str, Any] | None = None) -> None:
    app = workbook = pythoncom = None
    com_initialized = False
    try:
        app, pythoncom, progid = _connect_wps(log)
        com_initialized = True
        for prop, value in (("Visible", False), ("DisplayAlerts", False), ("AskToUpdateLinks", False), ("EnableEvents", False), ("AutomationSecurity", 3)):
            try:
                setattr(app, prop, value)
            except Exception:
                pass
        try:
            app.Calculation = -4105
        except Exception:
            pass
        try:
            workbook = app.Workbooks.Open(str(staging_path), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        except Exception as exc:
            raise ProcessingError(f"已连接 WPS（{progid}），但无法在禁用外部链接更新模式下打开中间工作簿：{exc}") from exc
        log("已在 WPS 中打开中间工作簿，开始完整重算公式。")
        _recalculate_and_wait(app, log, "映射前")
        log("首次完整重算完成；开始应用可选结果映射及配置的后映射区域。")
        for sheet_name, start_row, end_row, mappings in formula_jobs:
            if end_row < start_row:
                continue
            worksheet = workbook.Worksheets(sheet_name)
            count = end_row - start_row + 1
            for mapping in mappings:
                result_col = str(mapping["result_col"]).strip().upper()
                target_col = str(mapping["target_col"]).strip().upper()
                source = worksheet.Range(f"{result_col}{start_row}:{result_col}{end_row}")
                values = [row[0] for row in _com_value_matrix(source.Value2, count, 1)]
                destination = worksheet.Range(f"{target_col}{start_row}:{target_col}{end_row}")
                _write_mapped_values(destination, values, mapping.get("write_as", "value"))
                mode = "文本" if str(mapping.get("write_as", "value")).lower() == "text" else "值"
                log(f"已将 {sheet_name}!{result_col}{start_row}:{result_col}{end_row} 按{mode}写入 {target_col} 列。")
        job_data_ranges = _configured_job_data_ranges((region_processing or {}).get("jobs", []),
                                                      formula_jobs=formula_jobs)
        for idx, rule in enumerate(post_mappings or [], 1):
            resolved = _resolve_post_mapping_areas(rule, job_data_ranges, idx)
            if resolved is None:
                log(f"第 {idx} 条后映射没有对应的数据行，已跳过整列映射。")
                continue
            source_bounds, target_bounds = resolved
            rows = source_bounds[3] - source_bounds[1] + 1
            cols = source_bounds[2] - source_bounds[0] + 1
            source_address = _area_address(source_bounds)
            target_address = _area_address(target_bounds)
            source_sheet = workbook.Worksheets(str(rule["source_sheet"]).strip())
            target_sheet = workbook.Worksheets(str(rule["target_sheet"]).strip())
            if _is_full_column_area(source_bounds):
                chunk_rows = max(1, 50_000 // cols)
                for first_row in range(source_bounds[1], source_bounds[3] + 1, chunk_rows):
                    last_row = min(source_bounds[3], first_row + chunk_rows - 1)
                    chunk_bounds = (source_bounds[0], first_row, source_bounds[2], last_row)
                    chunk_rows_count = last_row - first_row + 1
                    source_range = source_sheet.Range(_area_address(chunk_bounds))
                    source_values = _com_value_matrix(source_range.Value2, chunk_rows_count, cols)
                    target_range = target_sheet.Range(_area_address(
                        (target_bounds[0], first_row, target_bounds[2], last_row)
                    ))
                    target_range.Value2 = tuple(tuple(row) for row in source_values)
            else:
                source_range = source_sheet.Range(source_address)
                source_values = _com_value_matrix(source_range.Value2, rows, cols)
                target_range = target_sheet.Range(target_address)
                target_range.Value2 = tuple(tuple(row) for row in source_values)
            log(f"已应用后映射 {rule['source_sheet']}!{source_address} → {rule['target_sheet']}!{target_address}。")

        step_settings = dict(region_processing or {})
        step_settings.setdefault("CustomFormula", {"enabled": False, "rules": []})
        step_settings.setdefault("TextToColumns", {"enabled": False, "rules": []})
        step_settings.setdefault("Deduplicate", {"enabled": False, "rules": []})
        step_settings.setdefault("FormatConvert", {"enabled": False, "rules": []})
        step_settings.setdefault("RegionFilter", {"enabled": False, "rules": []})
        step_settings.setdefault("workflow", list(DEFAULT_WORKFLOW))
        applied_steps = _apply_region_processing(workbook, step_settings, formula_jobs, log, application=app)

        # Rebuild once after all result/post-mapping and region-processing writes so dependent formulas are current.
        recalc_phase = "结果映射、后映射及区域处理后" if applied_steps else "结果映射/后映射写入后"
        _recalculate_and_wait(app, log, recalc_phase)
        workbook.Save()
        log("WPS 已完成处理、第二次完整重算并保存工作簿；接下来只读取已配置的文本预览，不再改动工作簿。")
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"WPS 自动化处理中失败：{type(exc).__name__}: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _validate_post_mapping_conflicts(settings: dict[str, Any], source_records: list[list[list[Any]]]) -> None:
    formula_areas: dict[str, list[tuple[int, int, int, int]]] = {}
    result_mapping_areas: dict[str, list[tuple[int, int, int, int]]] = {}
    job_data_ranges = _configured_job_data_ranges(settings["jobs"], source_records=source_records)
    for job, records in zip(settings["jobs"], source_records):
        sheet = str(job["template_sheet"]).strip()
        start_row = int(job["start_row"])
        template_row = _formula_template_row(job)
        formula_areas.setdefault(sheet, []).append((1, template_row, 16_384, template_row))
        if records:
            end_row = start_row + len(records) - 1
            c1, c2 = parse_column_range(job["formula_range"])
            formula_areas[sheet].append((c1, start_row, c2, end_row))
            for mapping in job.get("mappings", []):
                col = column_number(mapping["target_col"])
                result_mapping_areas.setdefault(sheet, []).append((col, start_row, col, end_row))

    for idx, rule in enumerate(settings.get("post_mappings", []), 1):
        sheet = str(rule["target_sheet"]).strip()
        resolved = _resolve_post_mapping_areas(rule, job_data_ranges, idx)
        if resolved is None:
            continue
        _source, target = resolved
        for protected in formula_areas.get(sheet, []):
            if _areas_overlap(target, protected):
                raise ProcessingError(f"第 {idx} 条后映射目标 {sheet}!{_area_address(target)} 会覆盖保留公式区域 {_area_address(protected)}；已停止处理。")
        for protected in result_mapping_areas.get(sheet, []):
            if _areas_overlap(target, protected):
                raise ProcessingError(f"第 {idx} 条后映射目标 {sheet}!{_area_address(target)} 与结果映射目标 {_area_address(protected)} 重复；已停止处理。")


def _sanitize_filename_fragment(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value).strip())
    sanitized = re.sub(r"\s+", "_", sanitized).strip("._ ")
    return sanitized[:80] or "区域"


def _sanitize_merged_source_filename_stem(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value).strip()).strip(" .")
    sanitized = sanitized[:80].rstrip(" .") or "合并源"
    reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                *(f"LPT{i}" for i in range(1, 10))}
    return f"_{sanitized}" if sanitized.upper() in reserved else sanitized


def _validate_merged_source_filename(value: Any, job_index: int) -> None:
    filename = str(value or "").strip()
    if not filename:
        return
    if Path(filename).name != filename or "/" in filename or "\\" in filename:
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名只能填写文件名，不能包含目录路径。")
    if filename.lower() == ".xlsx":
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名不能只有 .xlsx 扩展名。")
    suffix = Path(filename).suffix
    if suffix and suffix.lower() != ".xlsx":
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名只能使用 .xlsx 扩展名。")
    if any(character in filename for character in '<>:"|?*') or filename.endswith((" ", ".")):
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名包含 Windows 不支持的字符。")
    stem = filename[:-5] if filename.lower().endswith(".xlsx") else filename
    reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                *(f"LPT{i}" for i in range(1, 10))}
    if stem.rstrip(" .").upper() in reserved:
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名不能使用 Windows 保留设备名。")
    if not stem.strip(" ."):
        raise ProcessingError(f"第 {job_index} 个目标表的合并源文件名不能为空或仅包含点和空格。")


def _region_xlsx_filename(output: dict[str, Any], index: int, collision: int = 1,
                          multiple_exports: bool = False) -> str:
    sheet = _sanitize_filename_fragment(str(output["sheet"]))
    export_suffix = str(index) if multiple_exports else ""
    collision_suffix = "" if collision == 1 else f"_{collision}"
    return f"{sheet}{export_suffix}{collision_suffix}.xlsx"


def _copy_file_exclusive(source: Path, destination: Path) -> None:
    """Copy a completed file without replacing any pre-existing destination."""
    descriptor = os.open(destination, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
    try:
        with os.fdopen(descriptor, "wb") as target, source.open("rb") as origin:
            while chunk := origin.read(1024 * 1024):
                target.write(chunk)
    except Exception:
        destination.unlink(missing_ok=True)
        raise


def _publish_processed_source_copy(source_copy: Path, original_path: Path, output_dir: Path,
                                   reserved_paths: set[Path] | None = None,
                                   filename_stem: str | None = None) -> Path:
    """Publish a merged source workbook beside the main output without replacing files."""
    output_dir.mkdir(parents=True, exist_ok=True)
    reserved = {Path(path).resolve() for path in (reserved_paths or set())}
    requested_name = str(filename_stem or "").strip()
    if requested_name.lower().endswith(".xlsx"):
        requested_name = requested_name[:-5]
    base = _sanitize_merged_source_filename_stem(requested_name or f"{original_path.stem}_已编辑")
    collision = 1
    while True:
        suffix = "" if collision == 1 else f"_{collision}"
        destination = output_dir / f"{base}{suffix}.xlsx"
        if destination.resolve() in reserved:
            collision += 1
            continue
        try:
            _copy_file_exclusive(source_copy, destination)
            return destination
        except FileExistsError:
            collision += 1


def _source_rules_for_sheets(settings: dict[str, Any], sheet_names):
    ordered_sheet_names = list(sheet_names)
    source_settings = deepcopy(settings)
    source_settings["jobs"] = []
    used_rules: set[tuple[str, int]] = set()
    any_rules = False
    for operation_key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
        operation = settings.get(operation_key, {})
        selected = []
        if operation.get("enabled", False):
            for index, rule in enumerate(operation.get("rules", [])):
                if str(rule.get("scope", "target")).strip().lower() != "source":
                    continue
                try:
                    _resolve_sheet_name(ordered_sheet_names, rule.get("sheet", ""), "源工作簿")
                except ProcessingError:
                    continue
                selected_rule = deepcopy(rule)
                selected_rule["scope"] = "target"
                selected.append(selected_rule)
                used_rules.add((operation_key, index))
        source_settings[operation_key] = {"enabled": bool(selected), "rules": selected}
        any_rules = any_rules or bool(selected)
    return source_settings, used_rules, any_rules


def _translated_source_rule_bounds(rule: dict[str, Any], operation_key: str,
                                   segment: dict[str, Any]) -> tuple[int, int, int, int] | None:
    """Map one original-sheet selector onto that source's rows in the merged sheet."""
    if operation_key == "RegionFilter":
        bounds, kind = _parse_region_filter_area(rule.get("range", ""), "源文件合并规则")
    else:
        bounds, kind = _parse_selection_details(rule.get("range", ""), "源文件合并规则")
    source_start, source_end = segment["source_start"], segment["source_end"]
    target_start, target_end = segment["target_start"], segment["target_end"]
    width = max(1, segment["max_col"])
    if kind == "sheet":
        return (1, target_start, width, target_end)
    if kind == "column":
        return (bounds[0], target_start, bounds[2], target_end)
    if kind == "row":
        source_row = bounds[1]
        if not source_start <= source_row <= source_end:
            return None
        target_row = target_start + source_row - source_start
        return (1, target_row, width, target_row)
    first = max(bounds[1], source_start)
    last = min(bounds[3], source_end)
    if first > last:
        return None
    return (bounds[0], target_start + first - source_start,
            bounds[2], target_start + last - source_start)


def _merge_source_workbooks(sources: list[dict[str, Any]], settings: dict[str, Any],
                            col_start: int, col_end: int,
                            log: Callable[[str], None]) -> dict[str, Any]:
    """Create a safe, row-concatenated workbook for one target sheet before source rules run."""
    try:
        fd, temp_name = tempfile.mkstemp(prefix="wps_source_merge_", suffix=".xlsx")
    except OSError as exc:
        raise ProcessingError(f"无法创建源文件合并临时工作簿：{exc}") from exc
    os.close(fd)
    merged_path = Path(temp_name)
    merged_sheet_name = "合并源数据"
    merged_book = Workbook()
    merged_sheet = merged_book.active
    merged_sheet.title = merged_sheet_name
    next_target_row = 1
    segments: list[dict[str, Any]] = []
    formula_cells: set[tuple[int, int]] = set()
    blank_formula_cache_count = 0
    try:
        for source_index, source in enumerate(sources):
            source_path = Path(source["path"]).expanduser().resolve()
            start_row = _positive_row(source.get("copy_start_row"), f"第 {source_index + 1} 份源文件复制起始行")
            requested_end = source.get("copy_end_row")
            end_row = _positive_row(requested_end, f"第 {source_index + 1} 份源文件复制结束行") if requested_end not in (None, "") else None
            keep_header = source_index == 0 or bool(source.get("keep_first_row", True))
            with _temporary_workbook_copy(source_path, "源工作簿", log) as safe_copy:
                raw_book = safe_copy.load_workbook(read_only=True, data_only=False, keep_links=False)
                values_book = safe_copy.load_workbook(read_only=True, data_only=True, keep_links=False)
                try:
                    requested_sheet = str(source.get("source_sheet") or "").strip()
                    if not raw_book.sheetnames or not values_book.sheetnames:
                        raise ProcessingError(f"源工作簿“{source_path.name}”没有可读取的工作表。")
                    resolved_sheet = _resolve_sheet_name(raw_book.sheetnames, requested_sheet,
                                                          f"源工作簿“{source_path.name}”")
                    if not requested_sheet:
                        log(f"{source_path.name} 未填写源工作表，默认使用第一个工作表“{resolved_sheet}”。")
                    if resolved_sheet not in values_book.sheetnames:
                        raise ProcessingError(f"源工作簿“{source_path.name}”中找不到工作表“{resolved_sheet}”。")
                    raw_ws, values_ws = raw_book[resolved_sheet], values_book[resolved_sheet]
                    actual_last, used_max_col = _last_nonempty_source_row(raw_ws)
                    if actual_last < start_row:
                        raise ProcessingError(f"源工作簿“{source_path.name}”工作表“{resolved_sheet}”在第 {start_row} 行及之后没有可复制数据，无法合并。")
                    if end_row is not None:
                        last_row = min(end_row, actual_last)
                    else:
                        last_row = actual_last
                        for row_number, cells in enumerate(
                            raw_ws.iter_rows(min_row=start_row, max_row=actual_last,
                                             min_col=1, max_col=used_max_col), start=start_row
                        ):
                            if not any(_cell_has_data(cell.value) for cell in cells):
                                last_row = row_number - 1
                                break
                    if last_row < start_row:
                        raise ProcessingError(f"源工作簿“{source_path.name}”从第 {start_row} 行起遇到空白边界，没有可合并内容。")
                    included_start = start_row if keep_header else start_row + 1
                    if not keep_header:
                        log(f"源文件 {source_path.name} 按设置跳过复制范围首行（源第 {start_row} 行）；其余内容接续合并。")
                    if included_start > last_row:
                        log(f"源文件 {source_path.name} 仅有首行，按设置跳过后没有可合并数据。")
                        continue
                    target_start = next_target_row
                    row_count = last_row - included_start + 1
                    target_end = target_start + row_count - 1
                    segment = {
                        "path": source_path, "source_sheet": resolved_sheet,
                        "source_start": included_start, "source_end": last_row,
                        "target_start": target_start, "target_end": target_end,
                        "max_col": used_max_col, "seed_formulas": {},
                        "is_first_sheet": resolved_sheet.casefold() == raw_book.sheetnames[0].casefold(),
                    }
                    source_rules = [
                        (operation_key, rule_index, rule)
                        for operation_key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert")
                        for rule_index, rule in enumerate(settings.get(operation_key, {}).get("rules", []))
                        if settings.get(operation_key, {}).get("enabled", False)
                        and str(rule.get("scope", "target")).strip().lower() == "source"
                        and ((str(rule.get("sheet", "")).strip().casefold() == resolved_sheet.casefold())
                             if str(rule.get("sheet", "")).strip()
                             else resolved_sheet.casefold() == raw_book.sheetnames[0].casefold())
                    ]
                    for operation_key, rule_index, rule in source_rules:
                        if operation_key != "CustomFormula" or str(rule.get("formula_mode", "custom")).lower() != "cell":
                            continue
                        seed_address = str(rule.get("seed_cell", "")).strip().upper()
                        try:
                            seed_col, seed_row, _max_col, _max_row = range_boundaries(seed_address)
                        except ValueError:
                            raise ProcessingError(f"源工作表“{resolved_sheet}”的公式种子单元格“{seed_address}”无效。")
                        if seed_col is None or seed_row is None:
                            raise ProcessingError(f"源工作表“{resolved_sheet}”的公式种子单元格“{seed_address}”无效。")
                        formula = raw_ws.cell(row=seed_row, column=seed_col).value
                        if not isinstance(formula, str) or not formula.startswith("="):
                            raise ProcessingError(f"源工作表“{resolved_sheet}”的公式种子单元格 {seed_address} 不含公式。")
                        if _find_external_formula(formula):
                            raise ProcessingError(f"源工作表“{resolved_sheet}”的公式种子单元格 {seed_address} 含外部引用或访问函数，已拒绝。")
                        segment["seed_formulas"][rule_index] = {"formula": formula, "address": seed_address}
                    raw_rows = raw_ws.iter_rows(min_row=included_start, max_row=last_row,
                                                min_col=1, max_col=used_max_col)
                    value_rows = values_ws.iter_rows(min_row=included_start, max_row=last_row,
                                                     min_col=1, max_col=used_max_col)
                    for row_offset, (raw_cells, value_cells) in enumerate(zip(raw_rows, value_rows)):
                        target_row = target_start + row_offset
                        for col_index, (raw_cell, value_cell) in enumerate(zip(raw_cells, value_cells), start=1):
                            value = value_cell.value
                            merged_sheet.cell(row=target_row, column=col_index, value=value)
                            if isinstance(raw_cell.value, str) and raw_cell.value.startswith("="):
                                formula_cells.add((target_row, col_index))
                                if value is None:
                                    blank_formula_cache_count += 1
                    segments.append(segment)
                    next_target_row = target_end + 1
                    log(f"已将源文件 {source_path.name}（{resolved_sheet}!第 {included_start}–{last_row} 行）合并到“{merged_sheet_name}”第 {target_start}–{target_end} 行。")
                finally:
                    raw_book.close()
                    values_book.close()
        if not segments:
            raise ProcessingError("源文件合并后没有可供模板使用的行。")
        used_rules: set[tuple[str, int]] = set()
        processing_settings = deepcopy(settings)
        processing_settings["jobs"] = []
        seed_formula_row = next_target_row
        for operation_key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
            operation = settings.get(operation_key, {})
            mapped_rules = []
            if operation.get("enabled", False):
                for rule_index, rule in enumerate(operation.get("rules", [])):
                    if str(rule.get("scope", "target")).strip().lower() != "source":
                        continue
                    for segment in segments:
                        requested_rule_sheet = str(rule.get("sheet", "")).strip()
                        if requested_rule_sheet:
                            if requested_rule_sheet.casefold() != segment["source_sheet"].casefold():
                                continue
                        elif not segment.get("is_first_sheet", False):
                            continue
                        mapped_bounds = _translated_source_rule_bounds(rule, operation_key, segment)
                        if mapped_bounds is None:
                            continue
                        mapped_rule = deepcopy(rule)
                        mapped_rule["scope"] = "target"
                        mapped_rule["sheet"] = merged_sheet_name
                        mapped_rule["range"] = _area_address(mapped_bounds)
                        if operation_key == "CustomFormula":
                            if str(rule.get("formula_mode", "custom")).lower() == "custom":
                                try:
                                    origin_bounds = _parse_selection_details(rule["range"], "源公式范围")[0]
                                    origin = f"{get_column_letter(origin_bounds[0])}{origin_bounds[1]}"
                                    destination = f"{get_column_letter(mapped_bounds[0])}{mapped_bounds[1]}"
                                    mapped_rule["formula"] = Translator(rule.get("formula", ""), origin=origin).translate_formula(destination)
                                except Exception as exc:
                                    raise ProcessingError(f"无法将源公式按合并行偏移：{exc}") from exc
                            else:
                                seed_info = segment["seed_formulas"].get(rule_index)
                                if seed_info is None:
                                    raise ProcessingError(f"源工作表“{segment['source_sheet']}”的公式种子单元格“{rule.get('seed_cell', '')}”不可读取。")
                                seed_col, _seed_row, _max_seed_col, _max_seed_row = range_boundaries(seed_info["address"])
                                seed_cell_address = f"{get_column_letter(seed_col)}{seed_formula_row}"
                                try:
                                    seed_formula = Translator(
                                        seed_info["formula"], origin=seed_info["address"]
                                    ).translate_formula(seed_cell_address)
                                except Exception as exc:
                                    raise ProcessingError(f"无法将源公式种子平移到合并工作表：{exc}") from exc
                                merged_sheet[seed_cell_address] = seed_formula
                                merged_sheet.row_dimensions[seed_formula_row].hidden = True
                                seed_formula_row += 1
                                mapped_rule["seed_cell"] = seed_cell_address
                        mapped_rules.append(mapped_rule)
                        used_rules.add((operation_key, rule_index))
            processing_settings[operation_key] = {"enabled": bool(mapped_rules), "rules": mapped_rules}
        for operation_key in ("RegionFilter", "Deduplicate", "FormatConvert"):
            for rule in processing_settings.get(operation_key, {}).get("rules", []):
                bounds = _parse_region_filter_area(rule["range"], "源文件合并规则")[0] if operation_key == "RegionFilter" else _parse_selected_area(rule["range"], "源文件合并规则")
                if any(bounds[0] <= col <= bounds[2] and bounds[1] <= row <= bounds[3]
                       for row, col in formula_cells):
                    raise ProcessingError(f"合并后的源文件区域 {merged_sheet_name}!{_area_address(bounds)} 含原始公式单元格；为避免缓存不一致，{MODULE_LABELS[operation_key]}已停止。")
        merged_book.save(merged_path)
        if blank_formula_cache_count:
            log(f"提示：合并源文件中有 {blank_formula_cache_count} 个源公式单元格没有已保存的计算缓存；这些单元格将按空值合并。")
        log(f"源文件合并完成：共 {len(segments)} 段、{next_target_row - 1} 行；后续源文件原件不会单独写入模板。")
        return {"path": merged_path, "sheet": merged_sheet_name, "last_row": next_target_row - 1,
                "first_source": Path(sources[0]["path"]).expanduser().resolve(),
                "settings": processing_settings, "used_rules": used_rules,
                "has_rules": any(operation.get("enabled", False) and operation.get("rules")
                                 for key, operation in processing_settings.items()
                                 if key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"))}
    except Exception:
        merged_book.close()
        merged_path.unlink(missing_ok=True)
        raise
    finally:
        try:
            merged_book.close()
        except Exception:
            pass


def _process_merged_source_workbook(merged: dict[str, Any], settings: dict[str, Any],
                                   col_start: int, col_end: int, log: Callable[[str], None],
                                   save_merged_source: bool, output_dir: Path,
                                   final_path: Path,
                                   retained_filename: str | None = None) -> tuple[list[list[Any]], Path | None]:
    """Process the merged copy, publish processed source rules, and read the published copy for template input."""
    if merged["has_rules"]:
        app = workbook = pythoncom = None
        com_initialized = False
        try:
            app, pythoncom, progid = _connect_wps(log)
            com_initialized = True
            for prop, value in (("Visible", False), ("DisplayAlerts", False), ("AskToUpdateLinks", False),
                                ("EnableEvents", False), ("AutomationSecurity", 3)):
                try:
                    setattr(app, prop, value)
                except Exception:
                    pass
            try:
                app.Calculation = -4105
            except Exception:
                pass
            try:
                workbook = app.Workbooks.Open(str(merged["path"]), UpdateLinks=0, ReadOnly=False, AddToMru=False)
            except Exception as exc:
                raise ProcessingError(f"已连接 WPS（{progid}），但无法打开合并源文件处理副本：{exc}") from exc
            log("源文件已先合并；现在仅对合并后的安全临时工作簿执行源范围规则，原始文件不修改。")
            _apply_region_processing(workbook, merged["settings"], [], log, application=app)
            workbook.Save()
        except ProcessingError:
            raise
        except Exception as exc:
            raise ProcessingError(f"WPS 处理合并源文件失败：{type(exc).__name__}: {exc}") from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            if com_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    retained_path = None
    # A source-rule result must be materialized as a separate workbook before
    # it becomes template input. The optional setting still controls saving a
    # merged workbook when no source rules were applied.
    if save_merged_source or merged["has_rules"]:
        retained_path = _publish_processed_source_copy(
            merged["path"], merged["first_source"], output_dir, reserved_paths={final_path},
            filename_stem=retained_filename,
        )
        if merged["has_rules"]:
            log(f"源侧规则处理完成，已另存独立副本；模板数据将从该副本读取（原始源文件均未修改）：{retained_path}")
        else:
            log(f"已保存合并源工作簿（原始源文件均未修改）：{retained_path}")
    source_input_path = retained_path or merged["path"]
    records = _read_source_records(source_input_path, merged["sheet"], 1, col_start, col_end,
                                   log, merged["last_row"], None, merged["used_rules"])
    return records, retained_path


def _read_source_records_with_rules(path: Path, sheet_name: str, start_row: int,
                                    col_start: int, col_end: int, log: Callable[[str], None],
                                    end_row: int | None, settings: dict[str, Any],
                                    source_rules_used: set[tuple[str, int]],
                                    retain_edited_sources: bool, output_dir: Path,
                                    final_path: Path,
                                    on_source_output: Callable[[Path], None] | None = None) -> list[list[Any]]:
    """Run source-scoped rules on a temporary WPS-edited copy before reading rows for the template."""
    any_source_rules = any(
        operation.get("enabled", False) and any(
            str(rule.get("scope", "target")).strip().lower() == "source"
            for rule in operation.get("rules", []))
        for operation in (settings.get(key, {}) for key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"))
    )
    if not any_source_rules:
        return _read_source_records(path, sheet_name, start_row, col_start, col_end, log,
                                   end_row, None, source_rules_used)

    with _temporary_workbook_copy(path, "源工作簿", log) as safe_copy:
        preview = safe_copy.load_workbook(read_only=True, data_only=False, keep_links=False)
        try:
            sheet_names = list(preview.sheetnames)
        finally:
            preview.close()
        source_settings, used_rules, has_matching_rules = _source_rules_for_sheets(settings, sheet_names)
        source_rules_used.update(used_rules)
        if not has_matching_rules:
            return _read_source_records(path, sheet_name, start_row, col_start, col_end, log,
                                        end_row, None, source_rules_used)

        app = workbook = pythoncom = None
        com_initialized = False
        try:
            app, pythoncom, progid = _connect_wps(log)
            com_initialized = True
            for prop, value in (("Visible", False), ("DisplayAlerts", False), ("AskToUpdateLinks", False),
                                ("EnableEvents", False), ("AutomationSecurity", 3)):
                try:
                    setattr(app, prop, value)
                except Exception:
                    pass
            try:
                app.Calculation = -4105
            except Exception:
                pass
            try:
                workbook = app.Workbooks.Open(str(safe_copy.path), UpdateLinks=0, ReadOnly=False, AddToMru=False)
            except Exception as exc:
                raise ProcessingError(f"已连接 WPS（{progid}），但无法打开源工作簿处理副本“{path.name}”：{exc}") from exc
            log(f"先在 WPS 临时副本中按规则处理源工作簿 {path.name}；原文件不修改。")
            _apply_region_processing(workbook, source_settings, [], log, application=app)
            workbook.Save()
        except ProcessingError:
            raise
        except Exception as exc:
            raise ProcessingError(f"WPS 处理源工作簿“{path.name}”失败：{type(exc).__name__}: {exc}") from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            if com_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        if retain_edited_sources:
            retained_path = _publish_processed_source_copy(
                safe_copy.path, path, output_dir, reserved_paths={final_path}
            )
            log(f"已保留源文件处理副本（原件未修改）：{retained_path}")
            if on_source_output is not None:
                on_source_output(retained_path)
        return _read_source_records(safe_copy.path, sheet_name, start_row, col_start, col_end,
                                    log, end_row, None, source_rules_used)


def _publish_staging_workbook(staging_path: Path, final_path: Path) -> None:
    """Publish a completed workbook without replacing a pre-existing file."""
    try:
        os.link(staging_path, final_path)
    except FileExistsError as exc:
        raise ProcessingError(f"输出文件在处理过程中被创建；为避免覆盖已停止：\n{final_path}") from exc
    except OSError:
        try:
            _copy_file_exclusive(staging_path, final_path)
        except FileExistsError as exc:
            raise ProcessingError(f"输出文件在处理过程中被创建；为避免覆盖已停止：\n{final_path}") from exc


def _resolve_selected_output_bounds(worksheet, selected_bounds: tuple[int, int, int, int],
                                    selection_kind: str) -> tuple[int, int, int, int]:
    used_rows = max(1, min(1_048_576, int(worksheet.max_row or 1)))
    used_cols = max(1, min(16_384, int(worksheet.max_column or 1)))
    min_col, min_row, max_col, max_row = selected_bounds
    if selection_kind == "sheet":
        return 1, 1, used_cols, used_rows
    if selection_kind == "row":
        return 1, min_row, used_cols, min_row
    if selection_kind == "column":
        return min_col, 1, min_col, used_rows
    return selected_bounds


def _column_dimension_for_index(worksheet, column: int):
    letter = get_column_letter(column)
    exact = worksheet.column_dimensions.get(letter)
    if exact is not None and (exact.min or column) <= column <= (exact.max or exact.min or column):
        return exact
    for dimension in worksheet.column_dimensions.values():
        first = dimension.min
        if first is None:
            try:
                first = column_index_from_string(str(dimension.index))
            except ValueError:
                continue
        last = dimension.max or first
        if first <= column <= last:
            return dimension
    return None


def _column_is_hidden(worksheet, column: int) -> bool:
    dimension = _column_dimension_for_index(worksheet, column)
    return bool(dimension.hidden) if dimension is not None else False


def _split_excel_format_sections(format_code: str) -> list[str]:
    sections: list[str] = []
    current: list[str] = []
    quoted = False
    escaped = False
    bracket_depth = 0
    for char in str(format_code or "General"):
        if escaped:
            current.append(char)
            escaped = False
        elif char == "\\":
            current.append(char)
            escaped = True
        elif char == '"' and bracket_depth == 0:
            current.append(char)
            quoted = not quoted
        elif not quoted and char == "[":
            bracket_depth += 1
            current.append(char)
        elif not quoted and char == "]" and bracket_depth:
            bracket_depth -= 1
            current.append(char)
        elif char == ";" and not quoted and bracket_depth == 0:
            sections.append("".join(current))
            current = []
        else:
            current.append(char)
    sections.append("".join(current))
    return sections


def _excel_format_condition(section: str) -> tuple[str, Decimal] | None:
    match = re.search(r"\[(<=|>=|<>|=|<|>)\s*(-?\d+(?:\.\d+)?)\]", section)
    if not match:
        return None
    try:
        return match.group(1), Decimal(match.group(2))
    except InvalidOperation:
        return None


def _format_condition_matches(value: Decimal, condition: tuple[str, Decimal]) -> bool:
    operator, threshold = condition
    return {
        "<": value < threshold,
        "<=": value <= threshold,
        ">": value > threshold,
        ">=": value >= threshold,
        "=": value == threshold,
        "<>": value != threshold,
    }[operator]


def _select_excel_format_section(value: Any, format_code: str) -> tuple[str, int, bool]:
    sections = _split_excel_format_sections(format_code)
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return sections[-1] if len(sections) >= 4 else sections[0], 0, False
    conditional = [_excel_format_condition(section) for section in sections]
    if any(condition is not None for condition in conditional):
        for index, condition in enumerate(conditional):
            if condition is not None and _format_condition_matches(number, condition):
                return sections[index], index, True
        for index, condition in enumerate(conditional):
            if condition is None:
                return sections[index], index, True
        return sections[0], 0, True
    if number < 0 and len(sections) >= 2:
        return sections[1], 1, False
    if number == 0 and len(sections) >= 3:
        return sections[2], 2, False
    return sections[0], 0, False


def _format_literal_affix(text: str) -> str:
    result: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            end = text.find('"', index + 1)
            if end < 0:
                result.append(text[index + 1:])
                break
            result.append(text[index + 1:end])
            index = end + 1
        elif char == "\\" and index + 1 < len(text):
            result.append(text[index + 1])
            index += 2
        elif char == "_" and index + 1 < len(text):
            result.append(" ")
            index += 2
        elif char == "*" and index + 1 < len(text):
            index += 2
        elif char == "[":
            end = text.find("]", index + 1)
            if end < 0:
                index += 1
                continue
            directive = text[index + 1:end]
            currency = re.match(r"\$([^\]-]+)(?:-[^]]*)?$", directive)
            if currency and currency.group(1):
                result.append(currency.group(1))
            index = end + 1
        else:
            result.append(char)
            index += 1
    return "".join(result)


def _numeric_pattern_span(section: str) -> tuple[int, int, str] | None:
    visible: list[tuple[str, int, int]] = []
    index = 0
    while index < len(section):
        char = section[index]
        if char == '"':
            end = section.find('"', index + 1)
            index = len(section) if end < 0 else end + 1
        elif char in ("\\", "_", "*") and index + 1 < len(section):
            index += 2
        elif char == "[":
            end = section.find("]", index + 1)
            index = len(section) if end < 0 else end + 1
        else:
            visible.append((char, index, index + 1))
            index += 1
    for offset, (char, start, end) in enumerate(visible):
        if char not in "0#?":
            continue
        stop = offset + 1
        while stop < len(visible):
            next_char = visible[stop][0]
            if next_char in "0#?,.":
                stop += 1
                continue
            if next_char in "Ee" and stop + 1 < len(visible) and visible[stop + 1][0] in "+-":
                stop += 2
                while stop < len(visible) and visible[stop][0] in "0#?":
                    stop += 1
                continue
            break
        start_pos = start
        end_pos = visible[stop - 1][2]
        return start_pos, end_pos, section[start_pos:end_pos]
    return None


def _format_number_with_excel_mask(value: Any, section: str, section_index: int,
                                   conditional_format: bool) -> str:
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return str(value)
    if not number.is_finite():
        return str(value)
    selected_negative_section = number < 0 and section_index == 1 and not conditional_format
    number_for_display = abs(number) if selected_negative_section else number
    pattern_info = _numeric_pattern_span(section)
    if pattern_info is None:
        literal = _format_literal_affix(section)
        return literal if literal else str(value) if section.strip().lower() == "general" else ""
    start, end, pattern = pattern_info
    if re.search(r"\d\s*/\s*[0#?]", pattern):
        return str(value)
    exponent_match = re.search(r"[Ee]([+-]?)([0#?]+)$", pattern)
    percent_count = section.count("%")
    scaled = number_for_display * (Decimal(100) ** percent_count)
    mantissa_pattern = pattern
    exponent_pattern = None
    if exponent_match:
        exponent_pattern = exponent_match.group(0)
        mantissa_pattern = pattern[:exponent_match.start()]
    integer_pattern, dot, decimal_pattern = mantissa_pattern.partition(".")
    trailing_scale_commas = len(integer_pattern) - len(integer_pattern.rstrip(","))
    integer_pattern = integer_pattern.rstrip(",")
    scaled /= Decimal(1000) ** trailing_scale_commas

    if exponent_pattern:
        decimals = len(decimal_pattern)
        try:
            rendered = f"{scaled:.{decimals}E}"
            mantissa, exponent_value = rendered.split("E")
            exponent_value_int = int(exponent_value)
            sign_char, exponent_digits = exponent_pattern[1], exponent_pattern[2:]
            exponent_sign = "+" if exponent_value_int >= 0 else "-"
            if sign_char == "-" and exponent_value_int >= 0:
                exponent_sign = ""
            formatted = (f"{mantissa}E{exponent_sign}"
                         f"{abs(exponent_value_int):0{max(1, len(exponent_digits))}d}")
        except (ValueError, InvalidOperation):
            return str(value)
    else:
        min_decimals = decimal_pattern.count("0")
        max_decimals = len(decimal_pattern)
        quantizer = Decimal(1).scaleb(-max_decimals)
        try:
            rounded = scaled.quantize(quantizer, rounding=ROUND_HALF_UP) if max_decimals else scaled.quantize(Decimal(1), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            return str(value)
        fixed = f"{abs(rounded):.{max_decimals}f}" if rounded < 0 else f"{rounded:.{max_decimals}f}"
        integer_text, _, decimal_text = fixed.partition(".")
        required_integer_zeros = integer_pattern.count("0")
        if not integer_text and required_integer_zeros:
            integer_text = "0"
        integer_text = integer_text.zfill(required_integer_zeros)
        if "," in integer_pattern:
            integer_text = f"{int(integer_text or '0'):,}"
        if max_decimals > min_decimals:
            while len(decimal_text) > min_decimals and decimal_text.endswith("0"):
                decimal_text = decimal_text[:-1]
        decimal_text = decimal_text.ljust(min_decimals, "0")
        if decimal_text:
            formatted = f"{integer_text or '0'}.{decimal_text}"
        else:
            formatted = integer_text
        if not formatted and scaled == 0 and not required_integer_zeros and not decimal_text:
            formatted = ""

    prefix = _format_literal_affix(section[:start])
    suffix = _format_literal_affix(section[end:])
    if number < 0 and not selected_negative_section and not conditional_format:
        prefix = "-" + prefix
    return prefix + formatted + suffix


def _format_temporal_value(value: Any, format_code: str) -> str:
    if isinstance(value, timedelta):
        seconds = int(value.total_seconds())
        hours, remainder = divmod(abs(seconds), 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    if isinstance(value, date) and not isinstance(value, datetime):
        value = datetime(value.year, value.month, value.day)
    if isinstance(value, datetime_time):
        value = datetime.combine(date(1899, 12, 30), value)
    if not isinstance(value, datetime):
        return str(value)
    section = _split_excel_format_sections(format_code)[0]
    tokens: list[tuple[str, str]] = []
    index = 0
    token_re = re.compile(r"AM/PM|A/P|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s", re.IGNORECASE)
    while index < len(section):
        char = section[index]
        if char == '"':
            end = section.find('"', index + 1)
            if end < 0:
                tokens.append(("literal", section[index + 1:]))
                break
            tokens.append(("literal", section[index + 1:end]))
            index = end + 1
        elif char == "\\" and index + 1 < len(section):
            tokens.append(("literal", section[index + 1]))
            index += 2
        elif char in "_*" and index + 1 < len(section):
            if char == "_":
                tokens.append(("literal", " "))
            index += 2
        elif char == "[":
            end = section.find("]", index + 1)
            if end < 0:
                index += 1
                continue
            directive = section[index + 1:end]
            if directive.lower() in ("h", "hh", "m", "mm", "s", "ss"):
                tokens.append(("elapsed", directive))
            else:
                currency = re.match(r"\$([^\]-]+)(?:-[^]]*)?$", directive)
                if currency and currency.group(1):
                    tokens.append(("literal", currency.group(1)))
            index = end + 1
        else:
            match = token_re.match(section, index)
            if match:
                tokens.append(("date", match.group(0)))
                index = match.end()
            else:
                tokens.append(("literal", char))
                index += 1

    date_tokens = [text.lower() for kind, text in tokens if kind == "date"]
    has_ampm = any(token in ("am/pm", "a/p") for token in date_tokens)
    rendered: list[str] = []
    prior_date_token = ""
    for token_index, (kind, token) in enumerate(tokens):
        if kind == "literal":
            rendered.append(token)
            continue
        if kind == "elapsed":
            rendered.append(token)
            continue
        lower = token.lower()
        if lower == "am/pm":
            rendered.append("AM" if value.hour < 12 else "PM")
        elif lower == "a/p":
            rendered.append("A" if value.hour < 12 else "P")
        elif lower.startswith("y"):
            rendered.append(f"{value.year % 100:02d}" if len(lower) <= 2 else f"{value.year:04d}")
        elif lower.startswith("d"):
            if len(lower) >= 4:
                rendered.append(("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")[value.weekday()])
            elif len(lower) == 3:
                rendered.append(("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")[value.weekday()])
            elif len(lower) == 2:
                rendered.append(f"{value.day:02d}")
            else:
                rendered.append(str(value.day))
        elif lower.startswith("h"):
            hour = value.hour % 12 or 12 if has_ampm else value.hour
            rendered.append(f"{hour:02d}" if len(lower) == 2 else str(hour))
        elif lower.startswith("s"):
            rendered.append(f"{value.second:02d}" if len(lower) == 2 else str(value.second))
        elif lower.startswith("m"):
            next_date_token = ""
            for later_kind, later_text in tokens[token_index + 1:]:
                if later_kind == "date":
                    next_date_token = later_text.lower()
                    break
                if later_kind == "literal" and later_text not in (":", ".", " "):
                    break
            minute = prior_date_token.startswith("h") or next_date_token.startswith("s")
            if minute:
                rendered.append(f"{value.minute:02d}" if len(lower) >= 2 else str(value.minute))
            elif len(lower) >= 4:
                rendered.append(("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")[value.month - 1])
            elif len(lower) == 3:
                rendered.append(("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")[value.month - 1])
            elif len(lower) == 2:
                rendered.append(f"{value.month:02d}")
            else:
                rendered.append(str(value.month))
        prior_date_token = lower
    return "".join(rendered)


def _value_as_export_text(value: Any, number_format: str = "General") -> str | None:
    """Return a non-formula text value approximating Excel's displayed cell text."""
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, datetime_time, timedelta)):
        if number_format and number_format.lower() not in ("general", "@") and is_date_format(number_format):
            return _format_temporal_value(value, number_format)
        if isinstance(value, datetime):
            return value.isoformat(sep=" ")
        if isinstance(value, date):
            return value.isoformat()
        if isinstance(value, datetime_time):
            return value.isoformat()
        return str(value)
    if isinstance(value, (int, float, Decimal)):
        if not number_format or number_format.lower() in ("general", "@"):
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
        section, section_index, conditional = _select_excel_format_section(value, number_format)
        return _format_number_with_excel_mask(value, section, section_index, conditional)
    return str(value)


def _read_wps_display_text(safe_copy_path: Path, requested_cells: list[tuple[str, int, int]],
                           log: Callable[[str], None]) -> dict[tuple[str, int, int], str]:
    """Read numeric/date Range.Text from a sanitized temporary copy when WPS COM is available."""
    if sys.platform != "win32" or not requested_cells:
        return {}
    app = workbook = pythoncom = None
    com_initialized = False
    result: dict[tuple[str, int, int], str] = {}
    try:
        app, pythoncom, progid = _connect_wps(log)
        com_initialized = True
        for prop, value in (("Visible", False), ("DisplayAlerts", False), ("AskToUpdateLinks", False),
                            ("EnableEvents", False), ("AutomationSecurity", 3)):
            try:
                setattr(app, prop, value)
            except Exception:
                pass
        workbook = app.Workbooks.Open(str(safe_copy_path), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        sheets: dict[str, Any] = {}
        for sheet_name, row, col in requested_cells:
            try:
                if sheet_name not in sheets:
                    sheets[sheet_name] = workbook.Worksheets(sheet_name)
                worksheet = sheets[sheet_name]
                text = worksheet.Cells(row, col).Text
                if text is None:
                    continue
                text = str(text)
                # Range.Text may return hashes solely because the source column is too narrow.
                if text and re.fullmatch(r"#+", text.strip()):
                    continue
                result[(sheet_name, row, col)] = text
            except Exception:
                continue
        if result:
            log(f"区域 XLSX 导出已通过 WPS {progid} 的单元格 Text 读取 {len(result):,} 个数值/日期显示文本。")
    except Exception as exc:
        log(f"提示：无法从 WPS COM 读取单元格显示文本，将按 number_format 转换；导出继续且原工作簿不变。原因：{type(exc).__name__}: {exc}")
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return result


def _copy_dimension_properties(source_dimension, target_dimension, names: tuple[str, ...]) -> None:
    for name in names:
        try:
            setattr(target_dimension, name, getattr(source_dimension, name))
        except (AttributeError, TypeError, ValueError):
            pass
    if getattr(source_dimension, "has_style", False):
        for name in ("font", "fill", "border", "alignment", "protection", "number_format"):
            try:
                setattr(target_dimension, name, copy(getattr(source_dimension, name)))
            except (AttributeError, TypeError, ValueError):
                pass


def _export_table_areas(workbook_path: Path, area_outputs: list[dict[str, Any]], output_dir: Path,
                        log: Callable[[str], None], reserved_paths: set[Path] | None = None) -> list[Path]:
    """Export areas from a sanitized read-only copy, never directly from WPS's saved file."""
    if not area_outputs:
        return []
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        with _temporary_workbook_copy(workbook_path, "已保存工作簿", log) as safe_copy:
            source_workbook = safe_copy.load_workbook(data_only=True, keep_links=False)
            reserved = {Path(path).resolve() for path in (reserved_paths or set())}
            return _export_table_areas_loaded(
                source_workbook, area_outputs, output_dir, log,
                multiple_exports=len(area_outputs) > 1, reserved_paths=reserved,
                source_copy_path=safe_copy.path,
            )
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"独立区域 XLSX 导出失败：{type(exc).__name__}: {exc}") from exc


def _export_table_areas_loaded(source_workbook, area_outputs: list[dict[str, Any]], output_dir: Path,
                               log: Callable[[str], None], multiple_exports: bool,
                               reserved_paths: set[Path], source_copy_path: Path | None = None) -> list[Path]:
    """Create one styled, text-valued XLSX workbook for each configured source area."""
    generated: list[Path] = []
    try:
        export_specs = []
        requested_display_cells = set()
        for index, output in enumerate(area_outputs, 1):
            sheet_name = str(output["sheet"]).strip()
            if sheet_name not in source_workbook.sheetnames:
                raise ProcessingError(f"第 {index} 项区域 XLSX 输出工作表“{sheet_name}”在已保存工作簿中不存在。")
            worksheet = source_workbook[sheet_name]
            selected_bounds, selection_kind = _parse_selection_details(
                output["range"], f"第 {index} 项区域 XLSX 输出"
            )
            configured_bounds = _resolve_selected_output_bounds(worksheet, selected_bounds, selection_kind)
            _check_region_xlsx_area_size(configured_bounds, f"第 {index} 项区域 XLSX 输出")
            bounds = _expand_bounds_for_merged_cells(
                worksheet, configured_bounds, f"第 {index} 项区域 XLSX 输出"
            )
            _check_region_xlsx_area_size(bounds, f"第 {index} 项区域 XLSX 输出（合并区域闭包）")
            min_col, min_row, max_col, max_row = bounds
            source_rows = list(range(min_row, max_row + 1))
            source_cols = list(range(min_col, max_col + 1))
            export_specs.append({
                "index": index, "output": output, "sheet_name": sheet_name, "worksheet": worksheet,
                "bounds": bounds, "configured_bounds": configured_bounds,
                "min_col": min_col, "min_row": min_row,
                "source_rows": source_rows, "source_cols": source_cols,
            })
            for source_row in source_rows:
                for source_col in source_cols:
                    value = worksheet.cell(source_row, source_col).value
                    if value is not None and not isinstance(value, bool) and isinstance(
                            value, (int, float, Decimal, datetime, date, datetime_time, timedelta)):
                        requested_display_cells.add((sheet_name, source_row, source_col))

        display_texts = _read_wps_display_text(
            source_copy_path, sorted(requested_display_cells), log
        ) if source_copy_path is not None else {}

        for spec in export_specs:
            index = spec["index"]
            output = spec["output"]
            sheet_name = spec["sheet_name"]
            worksheet = spec["worksheet"]
            bounds = spec["bounds"]
            configured_bounds = spec["configured_bounds"]
            min_col, min_row, max_col, max_row = bounds
            source_rows = spec["source_rows"]
            source_cols = spec["source_cols"]
            row_map = {source_row: source_row - min_row + 1 for source_row in source_rows}
            col_map = {source_col: source_col - min_col + 1 for source_col in source_cols}

            export_workbook = Workbook()
            export_sheet = export_workbook.active
            export_sheet.title = worksheet.title
            export_sheet.sheet_view.showGridLines = worksheet.sheet_view.showGridLines
            export_sheet.sheet_view.zoomScale = worksheet.sheet_view.zoomScale
            for name in ("defaultRowHeight", "defaultColWidth", "baseColWidth", "zeroHeight"):
                try:
                    setattr(export_sheet.sheet_format, name, getattr(worksheet.sheet_format, name))
                except (AttributeError, TypeError, ValueError):
                    pass
            export_sheet.sheet_format.zeroHeight = worksheet.sheet_format.zeroHeight

            total_chars = 0
            try:
                for target_row, source_row in enumerate(source_rows, 1):
                    source_dimension = worksheet.row_dimensions.get(source_row)
                    if source_dimension is not None:
                        _copy_dimension_properties(
                            source_dimension, export_sheet.row_dimensions[target_row],
                            ("height", "outlineLevel", "collapsed", "hidden", "thickTop", "thickBot")
                        )
                    for target_col, source_col in enumerate(source_cols, 1):
                        source_cell = worksheet.cell(source_row, source_col)
                        target_cell = export_sheet.cell(target_row, target_col)
                        if source_cell.has_style:
                            for style_name in ("font", "fill", "border", "alignment", "protection"):
                                setattr(target_cell, style_name, copy(getattr(source_cell, style_name)))
                            target_cell.number_format = source_cell.number_format
                        display_key = (sheet_name, source_row, source_col)
                        text = display_texts.get(display_key)
                        if text is None:
                            text = _value_as_export_text(source_cell.value, source_cell.number_format)
                        if text is not None:
                            total_chars += len(text)
                            if total_chars > MAX_REGION_XLSX_CHARS:
                                raise ProcessingError(
                                    f"第 {index} 项区域 XLSX 输出超过 {MAX_REGION_XLSX_CHARS:,} 个字符上限；请缩小范围。"
                                )
                            target_cell.value = text
                            target_cell.data_type = "s"
                for target_col, source_col in enumerate(source_cols, 1):
                    source_dimension = _column_dimension_for_index(worksheet, source_col)
                    target_dimension = export_sheet.column_dimensions[get_column_letter(target_col)]
                    if source_dimension is not None:
                        _copy_dimension_properties(
                            source_dimension, target_dimension,
                            ("width", "bestFit", "outlineLevel", "collapsed", "hidden")
                        )

                for merged in worksheet.merged_cells.ranges:
                    if not _areas_overlap(bounds, (merged.min_col, merged.min_row, merged.max_col, merged.max_row)):
                        continue
                    if (merged.min_row in row_map and merged.max_row in row_map
                            and merged.min_col in col_map and merged.max_col in col_map):
                        export_sheet.merge_cells(
                            start_row=row_map[merged.min_row], start_column=col_map[merged.min_col],
                            end_row=row_map[merged.max_row], end_column=col_map[merged.max_col]
                        )

                temp_handle, temp_name = tempfile.mkstemp(
                    prefix=".wps_region_export_", suffix=".xlsx", dir=str(output_dir)
                )
                os.close(temp_handle)
                temp_path = Path(temp_name)
                try:
                    export_workbook.save(temp_path)
                    if not temp_path.is_file() or temp_path.stat().st_size == 0:
                        raise ProcessingError(f"第 {index} 项区域 XLSX 输出未生成有效文件。")
                    collision = 1
                    while True:
                        target = output_dir / _region_xlsx_filename(output, index, collision, multiple_exports)
                        if target.resolve() in reserved_paths:
                            collision += 1
                            continue
                        try:
                            os.link(temp_path, target)
                            break
                        except FileExistsError:
                            collision += 1
                        except OSError:
                            if target.exists():
                                collision += 1
                                continue
                            try:
                                _copy_file_exclusive(temp_path, target)
                                break
                            except FileExistsError:
                                collision += 1
                    generated.append(target)
                finally:
                    temp_path.unlink(missing_ok=True)
            finally:
                export_workbook.close()

            if bounds != configured_bounds:
                log(f"区域 {sheet_name}!{_area_address(configured_bounds)} 与合并单元格相交，导出范围扩展为 {_area_address(bounds)}。")
            log(
                f"已导出区域 XLSX {sheet_name}!{_area_address(bounds)}：{target.name}；"
                f"选区内隐藏行列及其原始位置均保留；单元格值均为文本。"
            )
        return generated
    except Exception as exc:
        for path in generated:
            path.unlink(missing_ok=True)
        if isinstance(exc, ProcessingError):
            raise
        raise ProcessingError(f"独立区域 XLSX 导出失败：{type(exc).__name__}: {exc}") from exc

def _format_text_matrix(matrix: list[list[Any]]) -> str:
    return "\n".join("\t".join("" if value is None else str(value) for value in row) for row in matrix)


def convert_cell_value(value: Any, mode: str, epoch=WINDOWS_EPOCH) -> Any:
    """Pure value conversion used by both WPS range processing and Linux tests."""
    mode = str(mode).strip().lower()
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None if mode in ("number", "date") else value
    if mode == "text":
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat() if isinstance(value, date) else str(value)
    if mode == "number":
        if isinstance(value, bool):
            return int(value)
        try:
            number = Decimal(str(value).replace(",", "").strip())
        except (InvalidOperation, ValueError) as exc:
            raise ProcessingError(f"无法将 {value!r} 转换为数字。") from exc
        if not number.is_finite():
            raise ProcessingError(f"无法将 {value!r} 转换为有限数字。")
        return int(number) if number == number.to_integral_value() else float(number)
    if mode == "date":
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            try:
                result = from_excel(value, epoch=epoch)
                if isinstance(result, datetime):
                    return result
                if isinstance(result, date):
                    return datetime(result.year, result.month, result.day)
            except (OverflowError, ValueError) as exc:
                raise ProcessingError(f"无法将 Excel 数值 {value!r} 转换为日期。") from exc
        text = str(value).strip()
        normalized = text.replace("/", "-")
        for parser in (datetime.fromisoformat,):
            try:
                return parser(normalized)
            except ValueError:
                pass
        for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(normalized, fmt)
            except ValueError:
                pass
        raise ProcessingError(f"无法将 {value!r} 转换为日期；支持日期单元格、Excel 日期序号或 YYYY-MM-DD/YYYYMMDD 文本。")
    raise ProcessingError(f"不支持的转换类型：{mode}")


def convert_matrix_values(matrix: list[list[Any]], mode: str, epoch=WINDOWS_EPOCH) -> list[list[Any]]:
    converted = []
    for row_index, row in enumerate(matrix, 1):
        converted_row = []
        for col_index, value in enumerate(row, 1):
            try:
                converted_row.append(convert_cell_value(value, mode, epoch))
            except ProcessingError as exc:
                raise ProcessingError(f"第 {row_index} 行第 {col_index} 列：{exc}") from exc
        converted.append(converted_row)
    return converted


def split_text_column(values: list[Any], delimiter: str, column_formats: list[str]) -> tuple[list[list[Any]], int]:
    rows = []
    maximum_parts = 1
    for value in values:
        parts = [None] if value is None else str(value).split(delimiter)
        rows.append(parts)
        maximum_parts = max(maximum_parts, len(parts))
    if maximum_parts > len(column_formats):
        raise ProcessingError(f"分列结果最多有 {maximum_parts} 列，但仅配置了 {len(column_formats)} 个结果列格式。")
    result = []
    for row in rows:
        padded = row + [None] * (maximum_parts - len(row))
        converted = []
        for col_index, value in enumerate(padded):
            try:
                converted.append(convert_cell_value(value, column_formats[col_index]))
            except ProcessingError as exc:
                raise ProcessingError(f"分列结果第 {col_index + 1} 列：{exc}") from exc
        result.append(converted)
    return result, maximum_parts


FILTER_OPERATORS = {
    "equals", "not_equals", "contains", "not_contains", "starts_with", "ends_with",
    "greater", "greater_equal", "less", "less_equal", "blank", "not_blank",
}
FILTER_OPERATOR_LABELS = {
    "equals": "等于", "not_equals": "不等于", "contains": "包含", "not_contains": "不包含",
    "starts_with": "开头为", "ends_with": "结尾为", "greater": "大于",
    "greater_equal": "大于等于", "less": "小于", "less_equal": "小于等于",
    "blank": "为空", "not_blank": "不为空",
}


def _filter_compare_value(value: Any, criteria: str):
    if isinstance(value, (datetime, date)):
        return ("date", value.isoformat())
    try:
        number = Decimal(str(value).replace(",", "").strip())
        if number.is_finite():
            return ("number", number)
    except (InvalidOperation, ValueError, TypeError):
        pass
    return ("text", str(value).casefold())


def _filter_value_matches(value: Any, operator: str, criteria: str) -> bool:
    operator = str(operator).strip().lower()
    if operator == "blank":
        return value is None or (isinstance(value, str) and not value.strip())
    if operator == "not_blank":
        return not (value is None or (isinstance(value, str) and not value.strip()))
    actual_text, wanted_text = str(value if value is not None else "").casefold(), str(criteria).casefold()
    if operator == "equals":
        return _filter_compare_value(value, criteria) == _filter_compare_value(criteria, criteria)
    if operator == "not_equals":
        return not _filter_value_matches(value, "equals", criteria)
    if operator == "contains":
        return wanted_text in actual_text
    if operator == "not_contains":
        return wanted_text not in actual_text
    if operator == "starts_with":
        return actual_text.startswith(wanted_text)
    if operator == "ends_with":
        return actual_text.endswith(wanted_text)
    actual_kind, actual = _filter_compare_value(value, criteria)
    wanted_kind, wanted = _filter_compare_value(criteria, criteria)
    if actual_kind != wanted_kind:
        actual, wanted = actual_text, wanted_text
    try:
        return {"greater": actual > wanted, "greater_equal": actual >= wanted,
                "less": actual < wanted, "less_equal": actual <= wanted}[operator]
    except TypeError:
        return False


def filter_matrix(matrix: list[list[Any]], has_header: bool, column_index: int,
                  operator: str, criteria: str) -> tuple[list[list[Any]], int, set[int]]:
    """Keep matching rows inside the region, compact them, and preserve the header."""
    if not matrix:
        return [], 0, set()
    width = max((len(row) for row in matrix), default=0)
    rows = [list(row) + [None] * (width - len(row)) for row in matrix]
    keep_indexes: set[int] = set()
    result: list[list[Any]] = []
    if has_header:
        keep_indexes.add(0)
        result.append(rows[0])
    removed = 0
    for index in range(1 if has_header else 0, len(rows)):
        row = rows[index]
        if _filter_value_matches(row[column_index], operator, criteria):
            keep_indexes.add(index)
            result.append(row)
        else:
            removed += 1
    result.extend([[None] * width for _ in range(len(rows) - len(result))])
    return result, removed, keep_indexes


def _deduplicate_keep_indexes(matrix: list[list[Any]], has_header: bool = False, key_mode: str = "row",
                              key_indexes: list[int] | None = None) -> tuple[set[int], int]:
    keep_indexes: set[int] = set()
    if not matrix:
        return keep_indexes, 0
    if has_header:
        keep_indexes.add(0)
    seen = set()
    removed = 0
    for index, row in enumerate(matrix[1:] if has_header else matrix, start=1 if has_header else 0):
        if not any(_cell_has_data(value) for value in row):
            keep_indexes.add(index)
            continue
        key = tuple(row) if key_mode == "row" else tuple(row[value] for value in (key_indexes or []))
        try:
            already_seen = key in seen
        except TypeError:
            key = tuple(repr(value) for value in key)
            already_seen = key in seen
        if already_seen:
            removed += 1
        else:
            seen.add(key)
            keep_indexes.add(index)
    return keep_indexes, removed


def deduplicate_matrix(matrix: list[list[Any]], has_header: bool = False, key_mode: str = "row",
                       key_indexes: list[int] | None = None) -> tuple[list[list[Any]], int]:
    if not matrix:
        return [], 0
    width = max((len(row) for row in matrix), default=0)
    rows = [list(row) + [None] * (width - len(row)) for row in matrix]
    keep_indexes, removed = _deduplicate_keep_indexes(rows, has_header, key_mode, key_indexes)
    result = [rows[index] for index in sorted(keep_indexes)]
    result.extend([[None] * width for _ in range(len(rows) - len(result))])
    return result, removed


def _read_text_output_areas(workbook_path: Path, text_outputs: list[dict[str, Any]],
                            log: Callable[[str], None] = lambda _message: None) -> list[dict[str, Any]]:
    if not text_outputs:
        return []
    try:
        with _temporary_workbook_copy(workbook_path, "已保存工作簿", log) as safe_copy:
            workbook = safe_copy.load_workbook(data_only=True, keep_links=False, read_only=True)
            return _read_text_output_areas_loaded(workbook, text_outputs)
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"读取文本输出区域失败：{exc}") from exc


def _read_text_output_areas_loaded(workbook, text_outputs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    try:
        previews = []
        for index, output in enumerate(text_outputs, 1):
            sheet_name = str(output["sheet"]).strip()
            if sheet_name not in workbook.sheetnames:
                raise ProcessingError(f"第 {index} 项文本输出工作表“{sheet_name}”在已保存工作簿中不存在。")
            selected_bounds, selection_kind = _parse_selection_details(output["range"], f"第 {index} 项文本输出")
            worksheet = workbook[sheet_name]
            used_rows = max(1, min(1_048_576, int(worksheet.max_row or 1)))
            used_cols = max(1, min(16_384, int(worksheet.max_column or 1)))
            min_col, min_row, max_col, max_row = selected_bounds
            if selection_kind == "sheet":
                bounds = (1, 1, used_cols, used_rows)
            elif selection_kind == "row":
                bounds = (1, min_row, used_cols, min_row)
            elif selection_kind == "column":
                bounds = (min_col, 1, min_col, used_rows)
            else:
                bounds = selected_bounds
            _check_text_area_size(bounds, f"第 {index} 项文本输出（实际范围 {_area_address(bounds)}）")
            min_col, min_row, max_col, max_row = bounds
            chunks = []
            full_text_parts = []
            total_chars = 0
            for chunk_start in range(min_row, max_row + 1, TEXT_COPY_CHUNK_ROWS):
                chunk_end = min(max_row, chunk_start + TEXT_COPY_CHUNK_ROWS - 1)
                lines = []
                for row in worksheet.iter_rows(min_row=chunk_start, max_row=chunk_end,
                                               min_col=min_col, max_col=max_col, values_only=True):
                    lines.append("\t".join("" if value is None else str(value) for value in row))
                chunk_text = "\n".join(lines)
                separator_length = 1 if full_text_parts else 0
                total_chars += separator_length + len(chunk_text)
                if total_chars > MAX_TEXT_CHARS:
                    raise ProcessingError(f"第 {index} 项文本输出超过 {MAX_TEXT_CHARS:,} 个字符的剪贴板上限；请缩小区域。")
                full_text_parts.append(chunk_text)
                chunks.append({"start_row": chunk_start, "end_row": chunk_end, "text": chunk_text})
            previews.append({"sheet": sheet_name, "range": str(output["range"]),
                             "resolved_range": _area_address(bounds),
                             "text": "\n".join(full_text_parts), "chunks": chunks})
        return previews
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"读取文本输出区域失败：{exc}") from exc


def _detect_csv_encoding(path: Path) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
        try:
            with path.open("rb") as stream:
                while True:
                    chunk = stream.read(1024 * 1024)
                    if not chunk:
                        break
                    decoder.decode(chunk)
                decoder.decode(b"", final=True)
            return encoding
        except UnicodeDecodeError:
            continue
        except OSError as exc:
            raise ProcessingError(f"无法读取 CSV 文件：\n{path}\n{exc}") from exc
    raise ProcessingError(f"CSV 不是有效的 UTF-8 或 GB18030 文本：\n{path}")


def _detect_csv_dialect(path: Path, encoding: str):
    try:
        with path.open("r", encoding=encoding, newline="") as stream:
            sample = stream.read(65536)
        if not sample:
            return csv.excel
        try:
            return csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            return csv.excel
    except (OSError, UnicodeError) as exc:
        raise ProcessingError(f"无法识别 CSV 分隔格式：\n{path}\n{exc}") from exc


def _append_merge_source_rows(worksheet, rows, source_index: int, source_path: Path,
                              start_row: int, end_row: int | None, keep_first_row: bool,
                              current_output_row: int, log: Callable[[str], None]) -> int:
    appended = 0
    for source_row_number, source_row in rows:
        if source_row_number < start_row:
            continue
        if end_row is not None and source_row_number > end_row:
            break
        row = list(source_row)
        if end_row is None and source_row_number >= start_row and not any(_cell_has_data(value) for value in row):
            break
        if source_index > 0 and source_row_number == start_row and not keep_first_row:
            log(f"源文件 {source_path.name} 按设置跳过复制范围首行（第 {start_row} 行）。")
            continue
        if current_output_row + appended > 1_048_576:
            raise ProcessingError("合并结果超过 Excel 的 1,048,576 行上限。")
        if len(row) > 16_384:
            raise ProcessingError(f"源文件 {source_path.name} 第 {source_row_number} 行超过 Excel 的 16,384 列上限。")
        output_row = []
        for col_number, value in enumerate(row, 1):
            if isinstance(value, str):
                if len(value) > 32767:
                    raise ProcessingError(f"源文件 {source_path.name} 第 {source_row_number} 行第 {col_number} 列超过 Excel 单元格 32,767 字符上限。")
                if ILLEGAL_CHARACTERS_RE.search(value):
                    raise ProcessingError(f"源文件 {source_path.name} 第 {source_row_number} 行第 {col_number} 列包含 XLSX 不允许的控制字符。")
                cell = WriteOnlyCell(worksheet, value=value)
                cell.data_type = "s"  # Treat imported text literally; never turn CSV text into an Excel formula.
                output_row.append(cell)
            else:
                output_row.append(value)
        worksheet.append(output_row)
        appended += 1
    return appended


def _merge_only_output_paths(settings: dict[str, Any]) -> tuple[list[tuple[dict[str, Any], Path]], Path]:
    _validate_merge_only_static_settings(settings)
    output_dir = _normalize_path(settings.get("output_dir", ""), "输出目录")
    if not output_dir.is_dir():
        raise ProcessingError(f"输出目录不存在：\n{output_dir}")
    filename = str(settings.get("merge_output_filename", "")).strip()
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    final_path = (output_dir / filename).resolve()
    if final_path.exists():
        raise ProcessingError(f"合并输出文件已存在，为防止覆盖已停止：\n{final_path}")
    validated: list[tuple[dict[str, Any], Path]] = []
    for index, source in enumerate(settings["merge_sources"]):
        source_path = _normalize_path(source.get("path", ""), f"第 {index + 1} 份独立合并源文件")
        if not source_path.is_file():
            raise ProcessingError(f"第 {index + 1} 份独立合并源文件不存在或不是文件：\n{source_path}")
        if source_path.suffix.lower() == ".csv":
            if source_path.stat().st_size > MAX_CSV_FILE_BYTES:
                raise ProcessingError(f"CSV 文件超过 {MAX_CSV_FILE_BYTES // (1024 * 1024)} MiB 安全读取上限：\n{source_path}")
        else:
            _check_xlsx_file(source_path, f"第 {index + 1} 份独立合并源文件")
        if os.path.normcase(str(source_path)) == os.path.normcase(str(final_path)):
            raise ProcessingError("合并输出文件不能与任何源文件同名同路径。")
        validated.append((source, source_path))
    return validated, final_path


def process_merge_only(raw_settings: dict[str, Any], log: Callable[[str], None] = print) -> Path:
    """Merge a dedicated source list into XLSX without opening a template or workflow modules."""
    settings = normalize_settings(raw_settings)
    settings["merge_only"] = True
    validate_static_settings(settings)
    sources, final_path = _merge_only_output_paths(settings)
    fd, temp_name = tempfile.mkstemp(prefix=".merge_only_", suffix=".xlsx", dir=str(final_path.parent))
    os.close(fd)
    staging_path = Path(temp_name)
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("合并结果")
    output_rows = 0
    try:
        log("独立合并模式已启用：不读取模板、不执行源规则或其他工作流，只合并指定文件并输出 XLSX。")
        for source_index, (source, source_path) in enumerate(sources):
            start_row = _positive_row(source.get("copy_start_row", 1), f"第 {source_index + 1} 份源文件复制起始行")
            requested_end = source.get("copy_end_row")
            end_row = _positive_row(requested_end, f"第 {source_index + 1} 份源文件复制结束行") if requested_end not in (None, "") else None
            keep_first = source_index == 0 or bool(source.get("keep_first_row", True))
            if source_path.suffix.lower() == ".csv":
                encoding = _detect_csv_encoding(source_path)
                dialect = _detect_csv_dialect(source_path, encoding)
                try:
                    with source_path.open("r", encoding=encoding, newline="") as stream:
                        reader = csv.reader(stream, dialect)
                        rows = enumerate(reader, start=1)
                        count = _append_merge_source_rows(
                            worksheet, rows, source_index, source_path, start_row, end_row,
                            keep_first, output_rows + 1, log,
                        )
                except (csv.Error, UnicodeError, OSError) as exc:
                    raise ProcessingError(f"读取 CSV 源文件失败：\n{source_path}\n{exc}") from exc
                log(f"CSV {source_path.name}（编码 {encoding}）已读取。")
            else:
                with _temporary_workbook_copy(source_path, "源工作簿", log) as safe_copy:
                    raw_book = safe_copy.load_workbook(read_only=True, data_only=False, keep_links=False)
                    values_book = safe_copy.load_workbook(read_only=True, data_only=True, keep_links=False)
                    try:
                        requested_sheet = str(source.get("source_sheet") or "").strip()
                        if not raw_book.sheetnames or not values_book.sheetnames:
                            raise ProcessingError(f"源工作簿“{source_path.name}”没有可读取的工作表。")
                        sheet_name = _resolve_sheet_name(raw_book.sheetnames, requested_sheet,
                                                         f"源工作簿“{source_path.name}”")
                        if not requested_sheet:
                            log(f"{source_path.name} 未填写源工作表，默认使用第一个工作表“{sheet_name}”。")
                        if sheet_name not in values_book.sheetnames:
                            raise ProcessingError(f"源工作簿“{source_path.name}”中找不到工作表“{sheet_name}”。")
                        raw_sheet = raw_book[sheet_name]
                        values_sheet = values_book[sheet_name]
                        actual_last, _used_columns = _last_nonempty_source_row(raw_sheet)
                        if actual_last < start_row:
                            raise ProcessingError(f"源工作簿“{source_path.name}”工作表“{sheet_name}”在第 {start_row} 行及之后没有可合并内容。")
                        last_row = min(end_row, actual_last) if end_row is not None else actual_last
                        rows = enumerate(values_sheet.iter_rows(min_row=start_row, max_row=last_row, values_only=True),
                                         start=start_row)
                        count = _append_merge_source_rows(
                            worksheet, rows, source_index, source_path, start_row, end_row,
                            keep_first, output_rows + 1, log,
                        )
                    finally:
                        raw_book.close()
                        values_book.close()
                log(f"XLSX {source_path.name} 工作表“{sheet_name}”已读取。")
            output_rows += count
            log(f"已合并源文件 {source_path.name}：追加 {count} 行。")
        if output_rows == 0:
            raise ProcessingError("所有源文件均无可合并行；没有生成输出文件。")
        workbook.save(staging_path)
        _publish_staging_workbook(staging_path, final_path)
        log(f"独立合并完成：共 {len(sources)} 份源文件、{output_rows} 行；结果已另存为 {final_path}。")
        return final_path
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"独立合并失败：{type(exc).__name__}: {exc}") from exc
    finally:
        workbook.close()
        staging_path.unlink(missing_ok=True)


def process_workbooks(raw_settings: dict[str, Any], log: Callable[[str], None] = print,
                      on_text_previews: Callable[[list[dict[str, str]]], None] | None = None,
                      on_region_outputs: Callable[[list[Path]], None] | None = None,
                      on_source_outputs: Callable[[list[Path]], None] | None = None) -> Path:
    """Merge each target's sources first, process that merged copy, then edit the template."""
    settings = normalize_settings(raw_settings)
    if settings.get("merge_only"):
        return process_merge_only(settings, log)
    merged_temp_paths: list[Path] = []
    try:
        return _process_template_workbooks(
            settings, log, on_text_previews, on_region_outputs, on_source_outputs, merged_temp_paths
        )
    finally:
        for merged_path in merged_temp_paths:
            merged_path.unlink(missing_ok=True)


def _process_template_workbooks(settings: dict[str, Any], log: Callable[[str], None],
                                on_text_previews: Callable[[list[dict[str, str]]], None] | None,
                                on_region_outputs: Callable[[list[Path]], None] | None,
                                on_source_outputs: Callable[[list[Path]], None] | None,
                                merged_temp_paths: list[Path]) -> Path:
    """Process normalized template settings and leave merged workbooks to the outer cleanup scope."""
    template_path, final_path = _validate_inputs(settings)
    if any(job.get("sources") for job in settings["jobs"]):
        copy_start, copy_end = column_number(settings["copy_start_col"]), column_number(settings["copy_end_col"])
    else:
        copy_start = copy_end = 1
    all_records: list[list[list[Any]]] = []
    source_rules_used: set[tuple[str, int]] = set()
    source_output_paths: list[Path] = []
    if any(job.get("sources") for job in settings["jobs"]):
        log("先按每个目标表的源文件顺序合并并执行源规则；命中源规则时先另存处理结果副本，再从副本读取模板数据，原始源文件不修改。")
    else:
        log("未配置源文件；将直接在模板临时副本上执行模板范围规则，并保留现有模板数据。")
    for job_idx, job in enumerate(settings["jobs"], 1):
        job_records: list[list[Any]] = []
        if job["sources"]:
            merged = _merge_source_workbooks(job["sources"], settings, copy_start, copy_end, log)
            merged_temp_paths.append(merged["path"])
            job_records, retained_path = _process_merged_source_workbook(
                merged, settings, copy_start, copy_end, log,
                bool(settings.get("retain_merged_sources", settings.get("retain_edited_sources", False))),
                final_path.parent, final_path,
                retained_filename=(str(job.get("merged_source_filename") or "").strip()
                                   or f"{job['template_sheet']}.xlsx"),
            )
            source_rules_used.update(merged["used_rules"])
            if retained_path is not None:
                source_output_paths.append(retained_path)
        all_records.append(job_records)
    for operation_key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
        operation = settings.get(operation_key, {})
        if not operation.get("enabled", False):
            continue
        for idx, rule in enumerate(operation.get("rules", [])):
            if str(rule.get("scope", "target")).lower() == "source" and (operation_key, idx) not in source_rules_used:
                raise ProcessingError(f"第 {idx + 1} 条{MODULE_LABELS[operation_key]}规则指定源工作表“{rule.get('sheet', '')}”，但没有任何已配置源文件使用该工作表；请检查源表名或规则处理对象。")
    _validate_pre_mapping_conflicts(settings, all_records)
    _validate_post_mapping_conflicts(settings, all_records)
    fd, temp_name = tempfile.mkstemp(prefix=f".{final_path.stem}_", suffix=".staging.xlsx", dir=str(final_path.parent))
    os.close(fd)
    staging_path = Path(temp_name)
    generated_region_files: list[Path] = []
    published_final = False
    preserve_final_after_region_failure = False
    try:
        log("校验模板公式和目标子表后，在输出目录创建临时副本。")
        formula_jobs = _create_staging_workbook(template_path, settings, all_records, staging_path, log)
        _calculate_in_wps(staging_path, formula_jobs, log, settings.get("post_mappings", []), settings)
        if not staging_path.is_file() or staging_path.stat().st_size == 0:
            raise ProcessingError("WPS 处理结束但临时工作簿不存在或为空；没有发布输出文件。")
        if final_path.exists():
            raise ProcessingError(f"输出文件在处理过程中被创建；为避免覆盖已停止：\n{final_path}")
        log("工作簿已保存；开始区域 XLSX 导出和文本区域读取（如有）。")
        region_xlsx_files: list[Path] = []
        text_previews = []
        for module_id in settings.get("workflow", DEFAULT_WORKFLOW):
            if module_id == "region_xlsx_output":
                try:
                    region_xlsx_files = _export_table_areas(
                        staging_path, settings.get("region_xlsx_outputs", []), final_path.parent, log,
                        reserved_paths={final_path},
                    )
                    generated_region_files.extend(region_xlsx_files)
                except Exception as region_exc:
                    try:
                        _publish_staging_workbook(staging_path, final_path)
                    except Exception as publish_exc:
                        raise ProcessingError(
                            f"区域 XLSX 输出失败：{region_exc}\nWPS 已完成计算并成功保存工作簿，但无法安全发布到目标路径：\n"
                            f"{final_path}\n发布原因：{publish_exc}"
                        ) from region_exc
                    published_final = True
                    preserve_final_after_region_failure = True
                    if on_region_outputs is not None and region_xlsx_files:
                        on_region_outputs(list(region_xlsx_files))
                    raise ProcessingError(
                        f"区域 XLSX 输出失败：{region_exc}\n"
                        f"WPS 已完成计算并成功保存 XLSX；该文件已保留：\n{final_path}"
                    ) from region_exc
            elif module_id == "text_output":
                text_previews = _read_text_output_areas(staging_path, settings.get("text_outputs", []), log)
        for preview in text_previews:
            log(f"已读取文本区域 {preview['sheet']}!{preview['range']}，内容将在处理完成后显示并可复制。")
        _publish_staging_workbook(staging_path, final_path)
        published_final = True
        staging_path.unlink(missing_ok=True)
        log(f"处理成功，已另存为：{final_path}")
        if on_text_previews is not None:
            on_text_previews(text_previews)
        if on_region_outputs is not None:
            on_region_outputs(list(region_xlsx_files))
        if settings.get("delete_processed_sources_after_success", False):
            protected_paths = {Path(template_path).resolve()}
            for job in settings.get("jobs", []):
                for source in job.get("sources", []):
                    source_path = str(source.get("path") or "").strip()
                    if source_path:
                        protected_paths.add(Path(source_path).expanduser().resolve())
            for source_copy in source_output_paths:
                try:
                    if source_copy.resolve() in protected_paths:
                        log(f"安全保护：跳过删除与原始源文件或模板同路径的副本：{source_copy}")
                        continue
                    source_copy.unlink(missing_ok=True)
                    log(f"模板处理成功，已按设置删除处理源副本：{source_copy}")
                except OSError as exc:
                    log(f"警告：模板处理成功，但无法删除处理源副本，已保留供检查：{source_copy}（{exc}）")
        if on_source_outputs is not None:
            try:
                on_source_outputs([path for path in source_output_paths if path.is_file()])
            except Exception as exc:
                log(f"警告：模板处理已成功，但无法通知界面源副本的保留状态：{exc}")
        return final_path
    except Exception:
        try:
            staging_path.unlink(missing_ok=True)
            for region_path in generated_region_files:
                region_path.unlink(missing_ok=True)
            if published_final and not preserve_final_after_region_failure:
                final_path.unlink(missing_ok=True)
        except Exception as cleanup_exc:
            raise ProcessingError(f"处理失败，同时无法清理临时工作簿：{staging_path}\n清理错误：{cleanup_exc}")
        raise

class App:
    def __init__(self, root):
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
        self.tk, self.filedialog, self.messagebox, self.ttk = tk, filedialog, messagebox, ttk
        self.root = root
        self.root.title(APP_NAME)
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
        except Exception:
            screen_width = screen_height = None
        min_width, min_height, _max_width, _max_height = window_size_limits(screen_width, screen_height)
        width, height = load_window_size(screen_width=screen_width, screen_height=screen_height)
        self.root.resizable(True, True)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(min_width, min_height)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.settings = load_settings()
        self.output_filename_custom = bool(self.settings.get("output_filename_custom", False))
        self._suppress_output_filename_trace = False
        startup_migration_notices = self.settings.pop("_migration_notices", [])
        self.vars: dict[str, Any] = {}
        self.job_vars: dict[str, Any] = {}
        self.operation_enabled: dict[str, Any] = {}
        self._region_output_files: list[Path] = []
        self._source_output_files: list[Path] = []
        self.active_job_index: int | None = None
        self.active_profile_name = str(self.settings.get("config_name") or "")
        self._build()
        self._render_jobs()
        self._render_merge_sources()
        self._render_area_configs()
        self._render_workflow()
        self._refresh_profile_choices()
        if startup_migration_notices:
            self.root.after(200, self._show_configuration_migration_notices, startup_migration_notices)
        self.root.after(150, self._drain_events)

    def _on_close(self):
        try:
            save_window_size(
                self.root.winfo_width(), self.root.winfo_height(),
                screen_width=self.root.winfo_screenwidth(),
                screen_height=self.root.winfo_screenheight(),
            )
        except Exception:
            pass
        self.root.destroy()

    def _on_canvas_configure(self, event):
        requested_width = self.content_frame.winfo_reqwidth()
        requested_height = self.content_frame.winfo_reqheight()
        width = max(event.width, requested_width)
        height = max(event.height, requested_height)
        self.main_canvas.itemconfigure(self.content_window, width=width, height=height)
        self.main_canvas.configure(scrollregion=(0, 0, width, height))

    def _on_content_configure(self, _event=None):
        bounds = self.main_canvas.bbox("all")
        if bounds:
            self.main_canvas.configure(scrollregion=bounds)

    def _entry(self, parent, row, label, key, width=64, browse=None):
        self.ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(8, 8), pady=3)
        var = self.tk.StringVar(value=str(self.settings.get(key, "")))
        self.vars[key] = var
        if key == "output_filename":
            var.trace_add("write", self._on_output_filename_edited)
        self.ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", padx=4, pady=3)
        if browse:
            self.ttk.Button(parent, text="浏览…", command=browse).grid(row=row, column=2, padx=4, pady=3)
        return var

    def _on_output_filename_edited(self, *_args):
        if not self._suppress_output_filename_trace:
            self.output_filename_custom = True

    def _set_output_filename(self, value: str, custom: bool):
        self._suppress_output_filename_trace = True
        try:
            self.vars["output_filename"].set(str(value))
        finally:
            self._suppress_output_filename_trace = False
        self.output_filename_custom = bool(custom)

    def _set_template_selected(self, path: str):
        self.vars["template_path"].set(str(path))
        if not self.output_filename_custom:
            default_name = _template_output_filename(path)
            if default_name:
                self._set_output_filename(default_name, custom=False)

    def _on_merge_mode_changed(self):
        if hasattr(self, "run_button"):
            self.run_button.configure(text="合并并保存" if self.merge_only_var.get() else "开始处理")

    def _render_merge_sources(self):
        if not hasattr(self, "merge_source_tree"):
            return
        for item in self.merge_source_tree.get_children():
            self.merge_source_tree.delete(item)
        for index, source in enumerate(self.settings.get("merge_sources", [])):
            self.merge_source_tree.insert(
                "", "end", iid=str(index),
                values=(source.get("path", ""), source.get("source_sheet", ""),
                        source.get("copy_start_row", 1), source.get("copy_end_row") or "",
                        "是" if index == 0 or source.get("keep_first_row", True) else "否"),
            )

    def _build(self):
        ttk = self.ttk
        viewport = ttk.Frame(self.root)
        viewport.grid(row=0, column=0, sticky="nsew")
        viewport.rowconfigure(0, weight=1)
        viewport.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.main_canvas = self.tk.Canvas(viewport, highlightthickness=0)
        vertical_scrollbar = ttk.Scrollbar(viewport, orient="vertical", command=self.main_canvas.yview)
        horizontal_scrollbar = ttk.Scrollbar(viewport, orient="horizontal", command=self.main_canvas.xview)
        self.main_canvas.configure(yscrollcommand=vertical_scrollbar.set,
                                   xscrollcommand=horizontal_scrollbar.set)
        self.main_canvas.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        outer = ttk.Frame(self.main_canvas, padding=9)
        self.content_frame = outer
        self.content_window = self.main_canvas.create_window((0, 0), window=outer, anchor="nw")
        self.main_canvas.bind("<Configure>", self._on_canvas_configure)
        outer.bind("<Configure>", self._on_content_configure)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=3)
        outer.rowconfigure(7, weight=1)
        files = ttk.LabelFrame(outer, text="模板、输出与源数据列", padding=7)
        files.grid(row=0, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)
        self._entry(files, 0, "模板工作簿（无宏 .xlsx）", "template_path", browse=self._browse_template)
        self._entry(files, 1, "输出目录", "output_dir", browse=self._browse_output_dir)
        self._entry(files, 2, "输出文件名", "output_filename", width=35)
        row = 3
        for key, label in (("copy_start_col", "源数据复制起始列（无源文件可留空）"), ("copy_end_col", "源数据复制末列（有源文件时必填）")):
            ttk.Label(files, text=label).grid(row=row, column=0, sticky="w", padx=(8, 8), pady=3)
            var = self.tk.StringVar(value=str(self.settings.get(key, "")))
            self.vars[key] = var
            ttk.Entry(files, textvariable=var, width=12).grid(row=row, column=1, sticky="w", padx=4, pady=3)
            row += 1
        ttk.Label(files, text="模板前缀匹配长度 N（1–120）").grid(row=5, column=0, sticky="w", padx=(8, 8), pady=3)
        self.vars["template_prefix_length"] = self.tk.StringVar(value=str(self.settings.get("template_prefix_length", 5)))
        ttk.Entry(files, textvariable=self.vars["template_prefix_length"], width=8).grid(row=5, column=1, sticky="w", padx=4, pady=3)
        ttk.Label(files, text="选模板时按文件名（不含扩展名）前 N 个字符匹配已保存配置；无匹配时不会自动套用其他配置。",
                  foreground="#4b5563", wraplength=1180, justify="left").grid(row=6, column=0, columnspan=3, sticky="w", padx=8, pady=3)

        merge_box = ttk.LabelFrame(outer, text="独立源文件合并（不需要模板）", padding=7)
        merge_box.grid(row=1, column=0, sticky="nsew", pady=(7, 0))
        merge_box.columnconfigure(1, weight=1)
        self.merge_only_var = self.tk.BooleanVar(value=bool(self.settings.get("merge_only", False)))
        ttk.Checkbutton(merge_box,
                        text="启用独立合并模式：仅合并此列表，不读取模板、不运行其他处理或工作流",
                        variable=self.merge_only_var, command=self._on_merge_mode_changed).grid(
                            row=0, column=0, columnspan=3, sticky="w", padx=4, pady=(1, 5))
        self.vars["merge_output_filename"] = self.tk.StringVar(value=str(self.settings.get("merge_output_filename", "合并结果.xlsx")))
        ttk.Label(merge_box, text="合并结果文件名（自动保存为 .xlsx）").grid(row=1, column=0, sticky="w", padx=(8, 6), pady=3)
        ttk.Entry(merge_box, textvariable=self.vars["merge_output_filename"], width=35).grid(row=1, column=1, sticky="w", padx=4, pady=3)
        self.merge_source_tree = ttk.Treeview(
            merge_box, columns=("path", "sheet", "from", "to", "keep_header"),
            show="headings", height=4, selectmode="browse",
        )
        for col, title, col_width in (("path", "源文件（首份可为 CSV）", 410), ("sheet", "工作表（空=首表）", 170),
                                      ("from", "复制起始行", 85), ("to", "复制结束行", 110),
                                      ("keep_header", "保留首行", 85)):
            self.merge_source_tree.heading(col, text=title)
            self.merge_source_tree.column(col, width=col_width, anchor="w")
        self.merge_source_tree.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=3, pady=4)
        merge_buttons = ttk.Frame(merge_box)
        merge_buttons.grid(row=3, column=0, columnspan=3, sticky="w", padx=3)
        for label, command in (("添加…", self._add_merge_source), ("编辑…", self._edit_merge_source),
                               ("删除", self._delete_merge_source), ("上移", lambda: self._move_merge_source(-1)),
                               ("下移", lambda: self._move_merge_source(1))):
            ttk.Button(merge_buttons, text=label, command=command).pack(side="left", padx=2)
        ttk.Label(merge_box,
                  text="第一份源文件首行固定保留；后续源文件支持逐个选择是否保留复制范围首行（默认保留）。仅首份可为 .csv；后续文件必须是 .xlsx。",
                  foreground="#4b5563", wraplength=1180, justify="left").grid(
                      row=4, column=0, columnspan=3, sticky="w", padx=5, pady=(4, 1))

        target_box = ttk.LabelFrame(outer, text="目标子表配置（可添加多个；源文件按列表顺序追加）", padding=7)
        target_box.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        target_box.columnconfigure(0, weight=1)
        target_box.rowconfigure(1, weight=1)
        actions = ttk.Frame(target_box)
        actions.grid(row=0, column=0, sticky="ew")
        ttk.Button(actions, text="添加目标表", command=self._add_job).pack(side="left", padx=3)
        ttk.Button(actions, text="删除目标表", command=self._delete_job).pack(side="left", padx=3)
        self.job_tree = ttk.Treeview(target_box, columns=("sheet", "start", "formula", "sources", "mappings"), show="headings", height=4, selectmode="browse")
        for col, title, width in (("sheet", "模板子表", 230), ("start", "粘贴起始行", 90), ("formula", "公式列范围", 120), ("sources", "源文件数（按顺序）", 140), ("mappings", "结果映射", 220)):
            self.job_tree.heading(col, text=title)
            self.job_tree.column(col, width=width, anchor="w")
        self.job_tree.grid(row=1, column=0, sticky="nsew", padx=3, pady=4)
        self.job_tree.bind("<<TreeviewSelect>>", self._on_job_select)

        editor = ttk.LabelFrame(outer, text="选中目标子表的参数", padding=7)
        editor.grid(row=3, column=0, sticky="nsew", pady=(7, 0))
        editor.columnconfigure(0, weight=1)
        editor.columnconfigure(1, weight=1)
        top = ttk.Frame(editor)
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.job_vars = {key: self.tk.StringVar() for key in ("template_sheet", "merged_source_filename", "header_rows", "start_row", "formula_range")}
        self.job_vars["formula_template_row"] = self.tk.StringVar()
        self.job_vars["start_row"].trace_add("write", self._update_formula_template_row_display)
        fields = (("template_sheet", "模板子表名", 25), ("header_rows", "保留标题行数", 10),
                  ("start_row", "粘贴起始行", 13), ("formula_template_row", "公式模板行（自动）", 13),
                  ("formula_range", "公式列范围", 13))
        for i, (key, label, width) in enumerate(fields):
            ttk.Label(top, text=label).grid(row=0, column=i * 2, sticky="w", padx=4)
            state = "readonly" if key == "formula_template_row" else "normal"
            ttk.Entry(top, textvariable=self.job_vars[key], width=width, state=state).grid(row=0, column=i * 2 + 1, sticky="w", padx=4)
        ttk.Label(top, text="合并源文件名（空=模板子表名）").grid(row=1, column=0, sticky="w", padx=4, pady=(4, 0))
        ttk.Entry(top, textvariable=self.job_vars["merged_source_filename"], width=48).grid(
            row=1, column=1, columnspan=9, sticky="w", padx=4, pady=(4, 0))
        self.source_tree = ttk.Treeview(editor, columns=("path", "sheet", "from", "to", "keep_header"), show="headings", height=7)
        for col, title, width in (("path", "源文件路径", 355), ("sheet", "源工作表（空=首表）", 160), ("from", "复制起始行", 85), ("to", "复制结束行", 110), ("keep_header", "保留首行", 80)):
            self.source_tree.heading(col, text=title)
            self.source_tree.column(col, width=width, anchor="w")
        self.source_tree.grid(row=1, column=0, sticky="nsew", padx=3, pady=5)
        source_buttons = ttk.Frame(editor)
        source_buttons.grid(row=2, column=0, sticky="w", padx=3)
        for label, command in (("添加源文件…", self._add_source), ("编辑…", self._edit_source), ("删除", self._delete_source), ("上移", lambda: self._move_source(-1)), ("下移", lambda: self._move_source(1))):
            ttk.Button(source_buttons, text=label, command=command).pack(side="left", padx=2)
        mapping_panel = ttk.LabelFrame(editor, text="可选结果映射：公式结果列 → 目标列", padding=3)
        mapping_panel.grid(row=1, column=1, sticky="nsew", padx=3, pady=5)
        mapping_panel.columnconfigure(0, weight=1)
        self.mapping_tree = ttk.Treeview(mapping_panel, columns=("result", "target", "write_as"), show="headings", height=6)
        for col, title, width in (("result", "公式结果列", 110), ("target", "目标列", 100), ("write_as", "写入类型", 90)):
            self.mapping_tree.heading(col, text=title)
            self.mapping_tree.column(col, width=width, anchor="w")
        self.mapping_tree.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        mapping_buttons = ttk.Frame(mapping_panel)
        mapping_buttons.grid(row=1, column=0, sticky="w", padx=2, pady=2)
        for label, command in (("添加", self._add_mapping), ("编辑", self._edit_mapping), ("删除", self._delete_mapping)):
            ttk.Button(mapping_buttons, text=label, command=command).pack(side="left", padx=2)
        editor.rowconfigure(1, weight=1)
        note = "可只处理模板，无需添加源文件；无源目标子表会保留模板原内容。多份源文件会先按列表顺序纵向合并，再对合并副本执行源规则并用于编辑模板；首份始终保留复制范围首行，后续文件可分别选择是否保留（默认保留，避免误删数据）。合并副本名可按目标表自定义，留空默认用模板子表名；结果映射与后映射均可独立留空。"
        ttk.Label(editor, text=note, foreground="#4b5563", wraplength=1180, justify="left").grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=(5, 0))

        area_box = ttk.LabelFrame(outer, text="模板内部前/后映射、区域处理及文本输出", padding=5)
        area_box.grid(row=4, column=0, sticky="nsew", pady=(7, 0))
        area_box.columnconfigure(0, weight=1)
        area_box.rowconfigure(0, weight=1)
        area_tabs = ttk.Notebook(area_box)
        area_tabs.grid(row=0, column=0, sticky="nsew")
        self.pre_mapping_tree = self._create_area_tab(area_tabs, "前映射", "pre_mappings",
            (("source_sheet", "源工作表", 180), ("source_range", "源范围", 110), ("target_sheet", "目标工作表", 180), ("target_range", "目标范围", 110)))
        self.region_xlsx_tree = self._create_area_tab(area_tabs, "区域 XLSX 输出", "region_xlsx_outputs",
            (("sheet", "模板工作表", 220), ("range", "单元格范围", 180)))
        self.text_tree = self._create_area_tab(area_tabs, "文本预览/复制", "text_outputs",
            (("sheet", "模板工作表", 220), ("range", "单元格范围", 180)))
        self.post_mapping_tree = self._create_area_tab(area_tabs, "后映射", "post_mappings",
            (("source_sheet", "源工作表", 180), ("source_range", "源范围", 110), ("target_sheet", "目标工作表", 180), ("target_range", "目标范围", 110)))
        rule_columns = (("scope", "处理对象", 110), ("sheet", "工作表（留空=首表）", 150), ("range", "处理范围", 120))
        self.custom_formula_tree = self._create_operation_tab(area_tabs, "公式填充", "CustomFormula",
            rule_columns + (("formula_mode", "公式来源", 120), ("formula", "自定义公式", 260), ("seed_cell", "来源单元格", 100)))
        self.text_to_columns_tree = self._create_operation_tab(area_tabs, "文本分列", "TextToColumns",
            rule_columns + (("delimiter", "分隔符", 80), ("output_start_col", "输出起始列", 90), ("column_formats", "结果列格式", 190)))
        self.region_filter_tree = self._create_operation_tab(area_tabs, "区域筛选", "RegionFilter",
            rule_columns + (("filter_column", "条件列", 80), ("operator", "筛选条件", 120), ("criteria", "条件值", 120), ("has_header", "含表头", 80)))
        self.deduplicate_tree = self._create_operation_tab(area_tabs, "去重", "Deduplicate",
            rule_columns + (("has_header", "含表头", 80), ("key_mode", "判重方式", 100), ("key_columns", "关键列", 130)))
        self.format_convert_tree = self._create_operation_tab(area_tabs, "格式转换", "FormatConvert",
            rule_columns + (("mode", "目标类型", 100),))
        self.workflow_tree = self._create_workflow_tab(area_tabs)
        selector_help = ("目标表结果映射可选：将公式列的已计算结果写入公式范围外的目标列；前/后映射的源和目标范围可用列字母（如 A、DG）或 A:A 表示整列；后映射可留空，配置时在首次完整重算及结果映射后逐行复制源区域的已计算值。整列后映射目标在目标表只写 start_row 至本轮数据末行，保留上一行公式模板。范围选择：A1=整张工作表，按 WPS UsedRange 裁剪；D1=整列，A5=整行，B3,B3=单格；两个端点如 D8,B3=自动规范化的矩形。区域筛选专用范围：留空=整张工作表，A/a=整列，A1/a1=单格 A1，也可填 A1:F100。工作表名留空=第一张表，名称匹配不区分大小写。源规则先处理合并后的源副本，再将处理结果另存为独立副本，模板从该副本读取。公式填充支持输入自定义公式或从同表指定单元格取公式，按 WPS 相对引用填满范围；填充会替换所选范围的原有内容。"
                         f"每条处理规则最多 {MAX_PROCESSING_CELLS:,} 格。源文件规则在临时副本中由 WPS 执行，处理结果另存后才用于模板；模板规则作用于结果副本，原始源文件不回写。")
        ttk.Label(area_box, text=selector_help, foreground="#4b5563", wraplength=1180,
                  justify="left").grid(row=1, column=0, sticky="w", padx=4, pady=(4, 1))
        self.retain_merged_sources_var = self.tk.BooleanVar(value=bool(
            self.settings.get("retain_merged_sources", self.settings.get("retain_edited_sources", False))))
        ttk.Checkbutton(area_box, text="保存合并源工作簿：源侧规则命中时，处理结果始终先另存为独立副本并用于模板输入；勾选后无源侧规则时也保留合并副本。未勾选的未处理副本会在任务后清理（原始源及模板不覆盖）。",
                        variable=self.retain_merged_sources_var).grid(row=2, column=0, sticky="w", padx=4, pady=(3, 4))
        self.delete_processed_sources_after_success_var = self.tk.BooleanVar(value=bool(
            self.settings.get("delete_processed_sources_after_success", False)))
        ttk.Checkbutton(area_box, text="模板处理全部成功后删除已另存的源副本（默认保留；开启后仅在模板、区域输出和文本读取均成功后删除；处理中失败时保留副本供排查，不删除原始源文件或模板）。",
                        variable=self.delete_processed_sources_after_success_var).grid(row=3, column=0, sticky="w", padx=4, pady=(0, 4))

        controls = ttk.Frame(outer)
        controls.grid(row=5, column=0, sticky="ew", pady=7)
        self.run_button = ttk.Button(controls, text="开始处理", command=self._run)
        self.run_button.pack(side="left")
        self._on_merge_mode_changed()
        ttk.Button(controls, text="保存配置…", command=self._export_config).pack(side="left", padx=5)
        ttk.Button(controls, text="加载配置…", command=self._import_config).pack(side="left", padx=5)
        ttk.Button(controls, text="保存为默认配置", command=self._save_default).pack(side="left", padx=5)
        ttk.Button(controls, text="恢复默认", command=self._reset_config).pack(side="left", padx=5)
        ttk.Label(controls, text="命名配置").pack(side="left", padx=(12, 3))
        self.profile_choice = self.tk.StringVar(value=self.active_profile_name)
        self.profile_combo = ttk.Combobox(controls, textvariable=self.profile_choice, state="readonly", width=18)
        self.profile_combo.pack(side="left", padx=3)
        ttk.Button(controls, text="新建并保存", command=self._new_profile).pack(side="left", padx=3)
        ttk.Button(controls, text="保存当前", command=self._save_current_profile).pack(side="left", padx=3)
        ttk.Button(controls, text="载入所选", command=self._load_selected_profile).pack(side="left", padx=3)
        ttk.Button(controls, text="设置", command=self._open_settings).pack(side="left", padx=5)
        self.progress = ttk.Progressbar(controls, mode="indeterminate", length=160)
        self.progress.pack(side="right", padx=5)

        preview_frame = ttk.LabelFrame(outer, text="文本区域预览（按行换行、按列 Tab 分隔；复制按钮写入 Windows 剪贴板）", padding=5)
        preview_frame.grid(row=6, column=0, sticky="nsew", pady=(0, 6))
        self.preview_empty_label = ttk.Label(preview_frame, text="处理成功后，配置的每个文本区域会在此显示并提供独立复制按钮。", foreground="#4b5563")
        self.preview_empty_label.pack(anchor="w", padx=4, pady=4)
        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.pack(fill="both", expand=True)
        self.preview_notebook.pack_forget()

        log_frame = ttk.LabelFrame(outer, text="处理日志", padding=5)
        log_frame.grid(row=7, column=0, sticky="nsew")
        self.log_text = self.tk.Text(log_frame, height=8, wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._log("模板模式：每个目标子表维护有序源文件列表，先合并再应用源规则；独立合并模式可单独合并首份 CSV/XLSX 与后续 XLSX。")

    def _create_workflow_tab(self, notebook):
        ttk = self.ttk
        page = ttk.Frame(notebook, padding=6)
        notebook.add(page, text="模块工作流")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)
        ttk.Label(page, text="添加模块并用上移/下移定制执行顺序。核心依赖不可颠倒；公式填充必须早于文本分列、筛选、去重和格式转换。非法顺序会在保存或开始处理时给出具体原因。",
                  wraplength=1080, justify="left").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.workflow_tree = ttk.Treeview(page, columns=("label", "role"), show="headings", height=7, selectmode="browse")
        for column, title, width in (("label", "模块名称", 210), ("role", "配置/依赖说明", 600)):
            self.workflow_tree.heading(column, text=title)
            self.workflow_tree.column(column, width=width, anchor="w")
        self.workflow_tree.grid(row=1, column=0, sticky="nsew")
        controls = ttk.Frame(page)
        controls.grid(row=2, column=0, sticky="w", pady=(5, 0))
        self.module_choice = self.tk.StringVar()
        self.module_combo = ttk.Combobox(controls, textvariable=self.module_choice, state="readonly", width=34)
        self.module_combo.pack(side="left", padx=2)
        ttk.Button(controls, text="添加模块", command=self._add_workflow_module).pack(side="left", padx=2)
        ttk.Button(controls, text="移除所选", command=self._remove_workflow_module).pack(side="left", padx=2)
        ttk.Button(controls, text="上移", command=lambda: self._move_workflow_module(-1)).pack(side="left", padx=2)
        ttk.Button(controls, text="下移", command=lambda: self._move_workflow_module(1)).pack(side="left", padx=2)
        return self.workflow_tree

    def _render_workflow(self):
        for item in self.workflow_tree.get_children():
            self.workflow_tree.delete(item)
        workflow = self.settings.get("workflow", DEFAULT_WORKFLOW)
        descriptions = {
            "source_merge": "先将同一目标表的源文件按顺序纵向合并；首份保留首行，后续首行可选",
            "source_preprocess": "在合并源副本中执行源规则，另存处理结果副本后再供模板读取",
            "template_cleanup": "保留模板格式/公式种子，清理旧数据",
            "pre_mapping": "读取模板原值并在源数据前写入",
            "source_copy": "按目标表及源文件列表追加数据",
            "formula_fill": "按复制数据行数填充公式",
            "formula_recalc": "首次完整重算并保留公式结果",
            "result_mapping": "可选：将公式列已计算结果写到对应目标列",
            "post_mapping": "可选：结果映射后复制源区域的已计算值",
            "CustomFormula": "按自定义公式或同表种子单元格，用 WPS R1C1 相对引用填满所选区域",
            "TextToColumns": "区域处理阶段；规则可指定源文件或目标工作簿",
            "RegionFilter": "按条件筛选指定区域并压缩匹配记录",
            "Deduplicate": "稳定去重并保留首次记录",
            "FormatConvert": "将指定范围转换为文本、数字或日期",
            "final_recalc_save": "所有工作簿写入后最终完整重算并保存",
            "region_xlsx_output": "工作簿最终保存后导出配置的格式化区域 XLSX",
            "text_output": "只读末端文本预览/复制，必须在最终保存后",
        }
        for index, module_id in enumerate(workflow):
            self.workflow_tree.insert("", "end", iid=str(index),
                                      values=(MODULE_LABELS.get(module_id, module_id), descriptions.get(module_id, "")))
        label_to_id = {label: module_id for module_id, label in MODULE_LABELS.items()}
        available = [MODULE_LABELS[key] for key in MODULE_LABELS if key not in workflow]
        self.module_combo.configure(values=available)
        if self.module_choice.get() not in available:
            self.module_choice.set(available[0] if available else "")

    def _change_workflow(self, candidate):
        previous = list(self.settings.get("workflow", DEFAULT_WORKFLOW))
        self.settings["workflow"] = candidate
        try:
            validate_workflow(self.settings)
            _validate_region_module_dependencies(self.settings)
        except (ValueError, ProcessingError) as exc:
            self.settings["workflow"] = previous
            self.messagebox.showerror(APP_NAME, str(exc))
            return False
        self._render_workflow()
        return True

    def _add_workflow_module(self):
        selected_label = self.module_choice.get()
        selected = next((module_id for module_id, label in MODULE_LABELS.items() if label == selected_label), "")
        if selected in MODULE_LABELS and selected not in self.settings.get("workflow", []):
            candidate = add_workflow_module(list(self.settings.get("workflow", [])), selected)
            self._change_workflow(candidate)

    def _remove_workflow_module(self):
        selected = self.workflow_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        candidate = remove_workflow_module(list(self.settings.get("workflow", DEFAULT_WORKFLOW)), index)
        self._change_workflow(candidate)

    def _move_workflow_module(self, direction):
        selected = self.workflow_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        destination = index + direction
        previous = list(self.settings.get("workflow", DEFAULT_WORKFLOW))
        if not 0 <= destination < len(previous):
            return
        candidate = move_workflow_module(previous, index, direction)
        if self._change_workflow(candidate):
            self.workflow_tree.selection_set(str(destination))
            self.workflow_tree.focus(str(destination))

    def _refresh_profile_choices(self):
        names = list_profiles(PROFILE_DIR)
        self.profile_combo.configure(values=names)
        if self.active_profile_name in names:
            self.profile_choice.set(self.active_profile_name)
        elif self.profile_choice.get() not in names:
            self.profile_choice.set("")

    def _create_area_tab(self, notebook, title, key, columns):
        ttk = self.ttk
        page = ttk.Frame(notebook, padding=4)
        notebook.add(page, text=title)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)
        identifiers = tuple(column[0] for column in columns)
        tree = ttk.Treeview(page, columns=identifiers, show="headings", height=4, selectmode="browse")
        for identifier, heading, width in columns:
            tree.heading(identifier, text=heading)
            tree.column(identifier, width=width, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        buttons = ttk.Frame(page)
        buttons.grid(row=1, column=0, sticky="w", pady=(4, 0))
        for action, label in (("add", "添加"), ("edit", "编辑"), ("delete", "删除")):
            ttk.Button(buttons, text=label, command=lambda k=key, t=tree, a=action: self._edit_area_config(k, t, a)).pack(side="left", padx=2)
        return tree

    def _create_operation_tab(self, notebook, title, key, columns):
        ttk = self.ttk
        page = ttk.Frame(notebook, padding=4)
        notebook.add(page, text=title)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)
        enabled = self.tk.BooleanVar(value=self.settings.get(key, {}).get("enabled", False))
        self.operation_enabled[key] = enabled
        ttk.Checkbutton(page, text=f"启用 {title}", variable=enabled).grid(row=0, column=0, sticky="w", pady=(0, 3))
        identifiers = tuple(column[0] for column in columns)
        tree = ttk.Treeview(page, columns=identifiers, show="headings", height=3, selectmode="browse")
        for identifier, heading, width in columns:
            tree.heading(identifier, text=heading)
            tree.column(identifier, width=width, anchor="w")
        tree.grid(row=1, column=0, sticky="nsew")
        buttons = ttk.Frame(page)
        buttons.grid(row=2, column=0, sticky="w", pady=(4, 0))
        for action, label in (("add", "添加"), ("edit", "编辑"), ("delete", "删除"), ("up", "上移"), ("down", "下移")):
            ttk.Button(buttons, text=label, command=lambda k=key, t=tree, a=action: self._manage_operation_rule(k, t, a)).pack(side="left", padx=2)
        return tree

    def _render_area_configs(self):
        for key, tree in (("pre_mappings", self.pre_mapping_tree), ("post_mappings", self.post_mapping_tree),
                          ("region_xlsx_outputs", self.region_xlsx_tree), ("text_outputs", self.text_tree)):
            for item in tree.get_children():
                tree.delete(item)
            columns = tree["columns"]
            for idx, rule in enumerate(self.settings.get(key, [])):
                tree.insert("", "end", iid=str(idx), values=tuple(rule.get(column, "") for column in columns))
        operation_trees = (
            ("CustomFormula", self.custom_formula_tree, ("scope", "sheet", "range", "formula_mode", "formula", "seed_cell")),
            ("TextToColumns", self.text_to_columns_tree, ("scope", "sheet", "range", "delimiter", "output_start_col", "column_formats")),
            ("RegionFilter", self.region_filter_tree, ("scope", "sheet", "range", "filter_column", "operator", "criteria", "has_header")),
            ("Deduplicate", self.deduplicate_tree, ("scope", "sheet", "range", "has_header", "key_mode", "key_columns")),
            ("FormatConvert", self.format_convert_tree, ("scope", "sheet", "range", "mode")),
        )
        for key, tree, columns in operation_trees:
            operation = self.settings.get(key, {"enabled": False, "rules": []})
            self.operation_enabled[key].set(bool(operation.get("enabled", False)))
            for item in tree.get_children():
                tree.delete(item)
            for idx, rule in enumerate(operation.get("rules", [])):
                values = []
                for column in columns:
                    value = rule.get(column, "")
                    if column == "scope":
                        value = "源文件" if value == "source" else "目标工作簿"
                    elif column == "formula_mode":
                        value = {"custom": "自定义公式", "cell": "从单元格取公式"}.get(value, value)
                    elif column == "operator":
                        value = FILTER_OPERATOR_LABELS.get(value, value)
                    elif column == "has_header":
                        value = "是" if value else "否"
                    elif column == "key_mode":
                        value = "整行" if value == "row" else "指定列"
                    elif column == "column_formats":
                        value = ",".join(rule.get(column, ["text"]))
                    elif column == "key_columns":
                        value = ",".join(rule.get(column, []))
                    values.append(value)
                tree.insert("", "end", iid=str(idx), values=tuple(values))

    def _edit_area_config(self, key: str, tree, action: str):
        entries = self.settings.setdefault(key, [])
        selected = tree.selection()
        if action == "edit" and not selected:
            return
        if action == "delete":
            if selected:
                del entries[int(selected[0])]
                self._render_area_configs()
            return
        index = int(selected[0]) if action == "edit" else None
        existing = entries[index] if index is not None else {}
        is_mapping = key in ("pre_mappings", "post_mappings")
        title = {"pre_mappings": "前映射规则", "post_mappings": "后映射规则",
                 "region_xlsx_outputs": "区域 XLSX 输出", "text_outputs": "文本预览区域"}[key]
        fields = (("source_sheet", "源工作表"), ("source_range", "源区域"),
                  ("target_sheet", "目标工作表"), ("target_range", "目标区域")) if is_mapping else (("sheet", "模板工作表"), ("range", "单元格区域"))
        dialog = self.tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        variables = {}
        for row, (field, label) in enumerate(fields):
            variables[field] = self.tk.StringVar(value=str(existing.get(field, "")))
            self.ttk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
            self.ttk.Entry(dialog, textvariable=variables[field], width=34).grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        result = {"value": None}

        def accept():
            value = {field: variables[field].get().strip() for field, _label in fields}
            try:
                if is_mapping:
                    _validate_sheet_name(value["source_sheet"], "映射源")
                    _validate_sheet_name(value["target_sheet"], "映射目标")
                    source_bounds = _parse_area(value["source_range"], "映射源")
                    target_bounds = _parse_area(value["target_range"], "映射目标")
                    source_shape = (source_bounds[2] - source_bounds[0], source_bounds[3] - source_bounds[1])
                    target_shape = (target_bounds[2] - target_bounds[0], target_bounds[3] - target_bounds[1])
                    if source_shape != target_shape:
                        raise ProcessingError("源区域与目标区域的行数和列数必须一致。")
                    value["source_range"] = _area_address(source_bounds)
                    value["target_range"] = _area_address(target_bounds)
                else:
                    _validate_sheet_name(value["sheet"], "输出")
                    bounds, selection_kind = _parse_selection_details(value["range"], "输出")
                    if key == "text_outputs" and selection_kind not in ("sheet", "row", "column"):
                        _check_text_area_size(bounds, "文本输出")
                    elif key == "region_xlsx_outputs" and selection_kind not in ("sheet", "row", "column"):
                        _check_region_xlsx_area_size(bounds, "区域 XLSX 输出")
            except ProcessingError as exc:
                self.messagebox.showerror(APP_NAME, str(exc), parent=dialog)
                return
            result["value"] = value
            dialog.destroy()

        buttons = self.ttk.Frame(dialog)
        buttons.grid(row=len(fields), column=0, columnspan=2, pady=8)
        self.ttk.Button(buttons, text="确定", command=accept).pack(side="left", padx=5)
        self.ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        dialog.wait_window()
        if result["value"] is not None:
            if index is None:
                entries.append(result["value"])
                selected_index = len(entries) - 1
            else:
                entries[index] = result["value"]
                selected_index = index
            self._render_area_configs()
            if str(selected_index) in tree.get_children():
                tree.selection_set(str(selected_index))

    def _operation_rule_dialog(self, key: str, existing: dict[str, Any] | None = None):
        title = {"CustomFormula": "公式填充规则", "TextToColumns": "文本分列规则", "RegionFilter": "区域筛选规则",
                 "Deduplicate": "去重规则", "FormatConvert": "格式转换规则"}[key]
        base_fields = (("scope", "处理对象", "scope"), ("sheet", "工作表名（留空=首表；不区分大小写）", "entry"))
        if key == "CustomFormula":
            fields = base_fields + (("range", "填充范围（A1 表示整表）", "entry"),
                ("formula_mode", "公式来源", "formula_mode"),
                ("formula", "自定义公式（以 = 开头）", "entry"),
                ("seed_cell", "来源单元格（如 B2）", "entry"))
            initial = {"scope": "target", "sheet": "", "range": "", "formula_mode": "custom",
                       "formula": "", "seed_cell": ""}
        elif key == "TextToColumns":
            fields = base_fields + (("range", "单列源范围", "entry"),
                ("delimiter", r"分隔符（可填 \t / \n）", "entry"), ("output_start_col", "输出起始列", "entry"),
                ("column_formats", "结果列格式（text/number/date，逗号分隔）", "entry"))
            initial = {"scope": "target", "sheet": "", "range": "", "delimiter": ",", "output_start_col": "A", "column_formats": ["text"]}
        elif key == "RegionFilter":
            fields = base_fields + (("range", "筛选区域（留空=整页；A/a=整列、A1/a1=单格）", "entry"), ("has_header", "表头", "header"),
                ("filter_column", "条件列（绝对列名）", "entry"), ("operator", "筛选条件", "operator"), ("criteria", "条件值", "entry"))
            initial = {"scope": "target", "sheet": "", "range": "", "has_header": True,
                       "filter_column": "A", "operator": "equals", "criteria": ""}
        elif key == "Deduplicate":
            fields = base_fields + (("range", "整行/选列数据范围", "entry"), ("has_header", "表头", "header"),
                ("key_mode", "判重方式", "keymode"), ("key_columns", "关键列（如 A,C）", "entry"))
            initial = {"scope": "target", "sheet": "", "range": "", "has_header": False, "key_mode": "row", "key_columns": []}
        else:
            fields = base_fields + (("range", "转换范围", "entry"), ("mode", "目标类型", "mode"))
            initial = {"scope": "target", "sheet": "", "range": "", "mode": "text"}
        data = {**initial, **(existing or {})}
        dialog = self.tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        variables = {}
        for row, (field, label, kind) in enumerate(fields):
            self.ttk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
            if kind == "header":
                variable = self.tk.StringVar(value="有表头" if data.get(field, False) else "无表头")
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=("有表头", "无表头"), state="readonly", width=28)
            elif kind == "keymode":
                variable = self.tk.StringVar(value="整行" if data.get(field, "row") == "row" else "指定关键列")
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=("整行", "指定关键列"), state="readonly", width=28)
            elif kind == "scope":
                variable = self.tk.StringVar(value="源文件" if data.get(field, "target") == "source" else "目标工作簿")
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=("目标工作簿", "源文件"), state="readonly", width=28)
            elif kind == "operator":
                variable = self.tk.StringVar(value=FILTER_OPERATOR_LABELS.get(data.get(field, "equals"), "等于"))
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=tuple(FILTER_OPERATOR_LABELS.values()), state="readonly", width=28)
            elif kind == "mode":
                mode_labels = {"text": "文本", "number": "数字", "date": "日期"}
                variable = self.tk.StringVar(value=mode_labels.get(str(data.get(field, "text")), "文本"))
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=tuple(mode_labels.values()), state="readonly", width=28)
            elif kind == "formula_mode":
                formula_mode_labels = {"custom": "自定义公式", "cell": "从单元格取公式"}
                variable = self.tk.StringVar(value=formula_mode_labels.get(str(data.get(field, "custom")), "自定义公式"))
                widget = self.ttk.Combobox(dialog, textvariable=variable, values=tuple(formula_mode_labels.values()), state="readonly", width=28)
            else:
                value = data.get(field, "")
                if field in ("column_formats", "key_columns"):
                    value = ",".join(value)
                variable = self.tk.StringVar(value=str(value))
                widget = self.ttk.Entry(dialog, textvariable=variable, width=34)
            widget.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
            variables[field] = variable
        result = {"value": None}
        def accept():
            rule = {}
            mode_labels = {"文本": "text", "数字": "number", "日期": "date"}
            formula_mode_ids = {"自定义公式": "custom", "从单元格取公式": "cell"}
            operator_ids = {label: operator for operator, label in FILTER_OPERATOR_LABELS.items()}
            for field, _label, _kind in fields:
                value = variables[field].get().strip()
                if field == "scope":
                    rule[field] = "source" if value == "源文件" else "target"
                elif field == "has_header":
                    rule[field] = value == "有表头"
                elif field == "key_mode":
                    rule[field] = "row" if value == "整行" else "columns"
                elif field == "operator":
                    rule[field] = operator_ids.get(value, "equals")
                elif field == "mode":
                    rule[field] = mode_labels.get(value, "text")
                elif field == "formula_mode":
                    rule[field] = formula_mode_ids.get(value, "custom")
                elif field in ("column_formats", "key_columns"):
                    rule[field] = [part.strip().lower() if field == "column_formats" else part.strip().upper()
                                   for part in value.split(",") if part.strip()]
                elif field == "delimiter":
                    rule[field] = value.replace(r"\t", "\t").replace(r"\n", "\n")
                else:
                    rule[field] = value
            if key == "Deduplicate" and rule["key_mode"] == "row":
                rule["key_columns"] = []
            try:
                trial = normalize_settings(self.settings)
                trial[key]["rules"] = [*trial[key].get("rules", []), rule]
                validate_static_settings(trial)
            except ProcessingError as exc:
                self.messagebox.showerror(APP_NAME, str(exc), parent=dialog)
                return
            result["value"] = rule
            dialog.destroy()
        buttons = self.ttk.Frame(dialog)
        buttons.grid(row=len(fields), column=0, columnspan=2, pady=8)
        self.ttk.Button(buttons, text="确定", command=accept).pack(side="left", padx=5)
        self.ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        dialog.wait_window()
        return result["value"]

    def _manage_operation_rule(self, key: str, tree, action: str):
        rules = self.settings.setdefault(key, {"enabled": False, "rules": []}).setdefault("rules", [])
        selected = tree.selection()
        if action in ("edit", "delete", "up", "down") and not selected:
            return
        if action == "delete":
            del rules[int(selected[0])]
        elif action in ("up", "down"):
            index = int(selected[0])
            target = index + (-1 if action == "up" else 1)
            if 0 <= target < len(rules):
                rules[index], rules[target] = rules[target], rules[index]
                self._render_area_configs()
                tree.selection_set(str(target))
                return
        else:
            index = int(selected[0]) if action == "edit" else None
            value = self._operation_rule_dialog(key, rules[index] if index is not None else None)
            if value is None:
                return
            if index is None:
                rules.append(value)
                index = len(rules) - 1
            else:
                rules[index] = value
        self._render_area_configs()
        if rules and action != "delete":
            tree.selection_set(str(index))

    def _mapping_dialog(self, job, existing=None, ignore_index=None):
        tk, ttk = self.tk, self.ttk
        dialog = tk.Toplevel(self.root)
        dialog.title("结果映射")
        dialog.transient(self.root)
        dialog.grab_set()
        vals = {"result_col": tk.StringVar(value=(existing or {}).get("result_col", "")),
                "target_col": tk.StringVar(value=(existing or {}).get("target_col", "")),
                "write_as": tk.StringVar(value=(existing or {}).get("write_as", "value"))}
        for row, key, label in ((0, "result_col", "公式结果列"), (1, "target_col", "写入目标列")):
            ttk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
            ttk.Entry(dialog, textvariable=vals[key], width=15).grid(row=row, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(dialog, text="目标列写入为").grid(row=2, column=0, sticky="w", padx=8, pady=5)
        ttk.Combobox(dialog, textvariable=vals["write_as"], values=("value", "text"), state="readonly", width=12).grid(row=2, column=1, sticky="w", padx=5, pady=5)
        result = {"value": None}
        def accept():
            mapping = {key: vals[key].get().strip() for key in vals}
            try:
                result_col = column_number(mapping["result_col"])
                target_col = column_number(mapping["target_col"])
                formula_start, formula_end = parse_column_range(job["formula_range"])
            except ProcessingError as exc:
                self.messagebox.showerror(APP_NAME, str(exc), parent=dialog)
                return
            if not formula_start <= result_col <= formula_end:
                self.messagebox.showerror(APP_NAME, f"公式结果列必须位于公式范围 {job['formula_range']} 内。", parent=dialog)
                return
            if formula_start <= target_col <= formula_end:
                self.messagebox.showerror(APP_NAME, "结果映射目标列不能位于公式范围内，以免覆盖公式。", parent=dialog)
                return
            mappings = job.setdefault("mappings", [])
            if any(index != ignore_index and column_number(item.get("target_col", "")) == target_col
                   for index, item in enumerate(mappings)):
                self.messagebox.showerror(APP_NAME, "每条结果映射必须写入不同目标列。", parent=dialog)
                return
            if mapping["write_as"] not in ("value", "text"):
                self.messagebox.showerror(APP_NAME, "写入类型只能是 value（值）或 text（文本）。", parent=dialog)
                return
            mapping["result_col"] = get_column_letter(result_col)
            mapping["target_col"] = get_column_letter(target_col)
            result["value"] = mapping
            dialog.destroy()
        buttons = ttk.Frame(dialog)
        buttons.grid(row=3, column=0, columnspan=2, pady=8)
        ttk.Button(buttons, text="确定", command=accept).pack(side="left", padx=5)
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        dialog.wait_window()
        return result["value"]

    def _ensure_result_mapping_workflow(self):
        workflow = list(self.settings.get("workflow", DEFAULT_WORKFLOW))
        if "result_mapping" not in workflow:
            self.settings["workflow"] = add_workflow_module(workflow, "result_mapping")

    def _add_mapping(self):
        job = self._selected_job()
        if job is None:
            return
        mapping = self._mapping_dialog(job)
        if mapping:
            job.setdefault("mappings", []).append(mapping)
            self._ensure_result_mapping_workflow()
            self._show_job(self._current_index())
            self._refresh_job_row(self._current_index())

    def _edit_mapping(self):
        job = self._selected_job()
        selected = self.mapping_tree.selection()
        if job is None or not selected:
            return
        index = int(selected[0])
        mapping = self._mapping_dialog(job, job["mappings"][index], ignore_index=index)
        if mapping:
            job["mappings"][index] = mapping
            self._ensure_result_mapping_workflow()
            self._show_job(self._current_index())
            self.mapping_tree.selection_set(str(index))
            self._refresh_job_row(self._current_index())

    def _delete_mapping(self):
        job = self._selected_job()
        selected = self.mapping_tree.selection()
        if job is not None and selected:
            del job.setdefault("mappings", [])[int(selected[0])]
            self._show_job(self._current_index())
            self._refresh_job_row(self._current_index())

    def _render_jobs(self, select_index: int | None = None):
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)
        for idx, job in enumerate(self.settings.get("jobs", [])):
            source_count = sum(bool(str(source.get("path") or "").strip()) for source in job.get("sources", []))
            mapping_summary = ", ".join(f"{item.get('result_col', '')}→{item.get('target_col', '')}"
                                         for item in job.get("mappings", [])) or "（无）"
            self.job_tree.insert("", "end", iid=str(idx), values=(job.get("template_sheet", ""), job.get("start_row", ""),
                                                                  job.get("formula_range", ""), source_count, mapping_summary))
        if self.settings.get("jobs"):
            index = 0 if select_index is None else max(0, min(select_index, len(self.settings["jobs"]) - 1))
            self.job_tree.selection_set(str(index))
            self.job_tree.focus(str(index))
            self._show_job(index)

    def _current_index(self) -> int | None:
        selected = self.job_tree.selection()
        if not selected:
            return None
        return int(selected[0])

    def _focus_source_location(self, job_index: int, source_index: int):
        job_id = str(job_index)
        self.job_tree.selection_set(job_id)
        self.job_tree.focus(job_id)
        self.job_tree.see(job_id)
        self._show_job(job_index)
        source_id = str(source_index)
        if source_id in self.source_tree.get_children():
            self.source_tree.selection_set(source_id)
            self.source_tree.focus(source_id)
            self.source_tree.see(source_id)

    def _commit_job_editor(self, show_errors=True):
        index = self.active_job_index
        if index is None:
            return True
        if index >= len(self.settings.get("jobs", [])):
            return True
        job = self.settings["jobs"][index]
        job["template_sheet"] = self.job_vars["template_sheet"].get().strip()
        job["merged_source_filename"] = self.job_vars["merged_source_filename"].get().strip()
        job["formula_range"] = self.job_vars["formula_range"].get().strip()
        try:
            job["header_rows"] = int(self.job_vars["header_rows"].get())
            job["start_row"] = int(self.job_vars["start_row"].get())
        except ValueError:
            if show_errors:
                self.messagebox.showerror(APP_NAME, "保留标题行数和粘贴起始行必须为整数。")
            return False
        return True

    def _show_job(self, index: int):
        if not (0 <= index < len(self.settings.get("jobs", []))):
            return
        self.active_job_index = index
        job = self.settings["jobs"][index]
        for key in ("template_sheet", "merged_source_filename", "header_rows", "start_row", "formula_range"):
            self.job_vars[key].set(str(job.get(key, "")))
        self._update_formula_template_row_display()
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
        for src_idx, source in enumerate(job.get("sources", [])):
            if source.get("_optional_placeholder") and not str(source.get("path") or "").strip():
                continue
            sheet_display = str(source.get("source_sheet") or "").strip() or "（首个工作表）"
            keep_header = True if src_idx == 0 else bool(source.get("keep_first_row", True))
            self.source_tree.insert("", "end", iid=str(src_idx), values=(source.get("path", ""), sheet_display,
                                    source.get("copy_start_row", ""), source.get("copy_end_row") or "",
                                    "是" if keep_header else "否"))
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)
        for map_idx, mapping in enumerate(job.get("mappings", [])):
            write_as = "文本" if mapping.get("write_as") == "text" else "值"
            self.mapping_tree.insert("", "end", iid=str(map_idx),
                                     values=(mapping.get("result_col", ""), mapping.get("target_col", ""), write_as))

    def _on_job_select(self, _event=None):
        if not self._commit_job_editor(show_errors=False):
            if self.active_job_index is not None:
                self.job_tree.selection_set(str(self.active_job_index))
            return
        index = self._current_index()
        if index is not None:
            previous_index = self.active_job_index
            if previous_index is not None:
                self._refresh_job_row(previous_index)
            self._show_job(index)
            self._refresh_job_row(index)

    def _update_formula_template_row_display(self, *_args):
        if not self.job_vars or "formula_template_row" not in self.job_vars:
            return
        try:
            start_row = int(self.job_vars["start_row"].get())
            value = str(start_row - 1) if start_row >= 2 else "需从第 2 行开始"
        except (TypeError, ValueError):
            value = "请输入有效行号"
        self.job_vars["formula_template_row"].set(value)

    def _show_configuration_migration_notices(self, notices):
        self.messagebox.showinfo(APP_NAME, "\n\n".join(notices))

    def _refresh_job_row(self, index):
        if not (0 <= index < len(self.settings["jobs"])):
            return
        job = self.settings["jobs"][index]
        source_count = sum(bool(str(source.get("path") or "").strip()) for source in job.get("sources", []))
        mapping_summary = ", ".join(f"{item.get('result_col', '')}→{item.get('target_col', '')}"
                                     for item in job.get("mappings", [])) or "（无）"
        self.job_tree.item(str(index), values=(job.get("template_sheet", ""), job.get("start_row", ""),
                                               job.get("formula_range", ""), source_count, mapping_summary))

    def _add_job(self):
        if not self._commit_job_editor():
            return
        self.settings["jobs"].append({"template_sheet": "新目标表", "merged_source_filename": "", "header_rows": 0, "start_row": 2, "formula_range": "A:A",
                                      "sources": [], "mappings": []})
        self._render_jobs(len(self.settings["jobs"]) - 1)

    def _delete_job(self):
        index = self._current_index()
        if index is None:
            return
        del self.settings["jobs"][index]
        self._render_jobs(min(index, len(self.settings["jobs"]) - 1) if self.settings["jobs"] else None)

    def _selected_job(self):
        if not self._commit_job_editor():
            return None
        index = self._current_index()
        return self.settings["jobs"][index] if index is not None else None

    def _source_dialog(self, existing=None, first_source=False):
        tk, ttk = self.tk, self.ttk
        dialog = tk.Toplevel(self.root)
        dialog.title("源文件设置")
        dialog.transient(self.root)
        dialog.grab_set()
        vals = {"path": tk.StringVar(value=(existing or {}).get("path", "")),
                "source_sheet": tk.StringVar(value=(existing or {}).get("source_sheet", "")),
                "copy_start_row": tk.StringVar(value=str((existing or {}).get("copy_start_row", 2))),
                "copy_end_row": tk.StringVar(value=str((existing or {}).get("copy_end_row") or ""))}
        vals["keep_first_row"] = tk.BooleanVar(value=True if first_source else bool((existing or {}).get("keep_first_row", True)))
        ttk.Label(dialog, text="源工作簿（.xlsx）").grid(row=0, column=0, sticky="w", padx=8, pady=5)
        path_entry = ttk.Entry(dialog, textvariable=vals["path"], width=58)
        path_entry.grid(row=0, column=1, padx=5, pady=5)
        def browse():
            path = self.filedialog.askopenfilename(filetypes=[("Excel 工作簿（无宏）", "*.xlsx")])
            if path:
                vals["path"].set(path)
        ttk.Button(dialog, text="浏览…", command=browse).grid(row=0, column=2, padx=5)
        for row, key, label in ((1, "source_sheet", "源工作表名（留空=首个工作表）"), (2, "copy_start_row", "复制起始行"), (3, "copy_end_row", "复制结束行（留空=遇整行空白）")):
            ttk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
            ttk.Entry(dialog, textvariable=vals[key], width=26).grid(row=row, column=1, sticky="w", padx=5, pady=5)
        header_option = ttk.Checkbutton(dialog, text="保留复制范围首行（首份源文件始终保留）",
                                        variable=vals["keep_first_row"])
        header_option.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        if first_source:
            header_option.configure(state="disabled")
        result = {"value": None}
        def accept():
            try:
                src = {key: vals[key].get().strip() for key in ("path", "source_sheet", "copy_start_row", "copy_end_row")}
                src["keep_first_row"] = bool(vals["keep_first_row"].get())
                if not src["path"]:
                    self.messagebox.showerror(APP_NAME, "请选择或填写源工作簿 .xlsx 路径。", parent=dialog)
                    path_entry.focus_set()
                    return
                src["copy_start_row"] = int(src["copy_start_row"])
                src["copy_end_row"] = int(src["copy_end_row"]) if src["copy_end_row"] else None
                result["value"] = src
                dialog.destroy()
            except ValueError:
                self.messagebox.showerror(APP_NAME, "复制起止行必须填写整数；结束行可以留空。", parent=dialog)
        btns = ttk.Frame(dialog)
        btns.grid(row=5, column=0, columnspan=3, pady=8)
        ttk.Button(btns, text="确定", command=accept).pack(side="left", padx=5)
        ttk.Button(btns, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        dialog.wait_window()
        return result["value"]

    def _merge_source_dialog(self, existing=None, first_source=False):
        tk, ttk = self.tk, self.ttk
        dialog = tk.Toplevel(self.root)
        dialog.title("独立合并源文件设置")
        dialog.transient(self.root)
        dialog.grab_set()
        source = existing or {}
        vals = {
            "path": tk.StringVar(value=source.get("path", "")),
            "source_sheet": tk.StringVar(value=source.get("source_sheet", "")),
            "copy_start_row": tk.StringVar(value=str(source.get("copy_start_row", 1))),
            "copy_end_row": tk.StringVar(value=str(source.get("copy_end_row") or "")),
            "keep_first_row": tk.BooleanVar(value=True if first_source else bool(source.get("keep_first_row", True))),
        }
        first_description = "首份源文件（可选 CSV 或 XLSX）" if first_source else "后续源文件（必须为 XLSX）"
        ttk.Label(dialog, text=first_description).grid(row=0, column=0, sticky="w", padx=8, pady=5)
        path_entry = ttk.Entry(dialog, textvariable=vals["path"], width=58)
        path_entry.grid(row=0, column=1, padx=5, pady=5)

        def browse():
            types = [("CSV 或 Excel 文件", "*.csv *.xlsx"), ("CSV 文件", "*.csv"),
                     ("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")] if first_source else [
                         ("Excel 工作簿（.xlsx）", "*.xlsx"), ("所有文件", "*.*")]
            path = self.filedialog.askopenfilename(filetypes=types)
            if path:
                vals["path"].set(path)

        ttk.Button(dialog, text="浏览…", command=browse).grid(row=0, column=2, padx=5)
        for row, key, label in ((1, "source_sheet", "工作表名（留空=首个工作表；CSV 忽略）"),
                                (2, "copy_start_row", "复制起始行"),
                                (3, "copy_end_row", "复制结束行（留空=遇整行空白）")):
            ttk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=5)
            ttk.Entry(dialog, textvariable=vals[key], width=26).grid(row=row, column=1, sticky="w", padx=5, pady=5)
        keep_option = ttk.Checkbutton(dialog, text="保留复制范围首行", variable=vals["keep_first_row"])
        keep_option.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        if first_source:
            keep_option.configure(state="disabled")
        result = {"value": None}

        def accept():
            try:
                path_text = vals["path"].get().strip()
                if not path_text:
                    raise ProcessingError("请选择或填写源文件路径。")
                suffix = Path(path_text).suffix.lower()
                if first_source and suffix not in (".csv", ".xlsx"):
                    raise ProcessingError("第一份源文件必须是 .csv 或 .xlsx。")
                if not first_source and suffix != ".xlsx":
                    raise ProcessingError("第二份及后续源文件必须是 .xlsx。")
                source_sheet = vals["source_sheet"].get().strip()
                if suffix == ".csv" and source_sheet:
                    raise ProcessingError("CSV 没有工作表，请将工作表名留空。")
                start_row = int(vals["copy_start_row"].get())
                end_text = vals["copy_end_row"].get().strip()
                end_row = int(end_text) if end_text else None
                if start_row < 1 or (end_row is not None and end_row < start_row):
                    raise ProcessingError("复制起始行必须大于零，结束行不得早于起始行。")
                result["value"] = {
                    "path": path_text, "source_sheet": source_sheet,
                    "copy_start_row": start_row, "copy_end_row": end_row,
                    "keep_first_row": True if first_source else bool(vals["keep_first_row"].get()),
                }
                dialog.destroy()
            except (ValueError, ProcessingError) as exc:
                self.messagebox.showerror(APP_NAME, str(exc), parent=dialog)

        buttons = ttk.Frame(dialog)
        buttons.grid(row=5, column=0, columnspan=3, pady=8)
        ttk.Button(buttons, text="确定", command=accept).pack(side="left", padx=5)
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=5)
        dialog.wait_window()
        return result["value"]

    def _add_merge_source(self):
        sources = self.settings.setdefault("merge_sources", [])
        source = self._merge_source_dialog(first_source=not bool(sources))
        if source:
            sources.append(source)
            self._render_merge_sources()

    def _edit_merge_source(self):
        selected = self.merge_source_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        sources = self.settings.setdefault("merge_sources", [])
        source = self._merge_source_dialog(sources[index], first_source=(index == 0))
        if source:
            sources[index] = source
            self._render_merge_sources()
            self.merge_source_tree.selection_set(str(index))

    def _delete_merge_source(self):
        selected = self.merge_source_tree.selection()
        if selected:
            del self.settings.setdefault("merge_sources", [])[int(selected[0])]
            self._render_merge_sources()

    def _move_merge_source(self, direction):
        selected = self.merge_source_tree.selection()
        sources = self.settings.setdefault("merge_sources", [])
        if not selected:
            return
        index = int(selected[0])
        destination = index + direction
        if 0 <= destination < len(sources):
            sources[index], sources[destination] = sources[destination], sources[index]
            self._render_merge_sources()
            self.merge_source_tree.selection_set(str(destination))

    def _add_source(self):
        job = self._selected_job()
        if job is None:
            return
        source = self._source_dialog(first_source=not bool(job.get("sources")))
        if source:
            job["sources"].append(source)
            self._show_job(self._current_index())
            self._refresh_job_row(self._current_index())

    def _edit_source(self):
        job = self._selected_job()
        selected = self.source_tree.selection()
        if job is None or not selected:
            return
        idx = int(selected[0])
        source = self._source_dialog(job["sources"][idx], first_source=(idx == 0))
        if source:
            job["sources"][idx] = source
            self._show_job(self._current_index())
            self.source_tree.selection_set(str(idx))
            self._refresh_job_row(self._current_index())

    def _delete_source(self):
        job = self._selected_job()
        selected = self.source_tree.selection()
        if job is not None and selected:
            del job["sources"][int(selected[0])]
            self._show_job(self._current_index())
            self._refresh_job_row(self._current_index())

    def _move_source(self, direction):
        job = self._selected_job()
        selected = self.source_tree.selection()
        if job is None or not selected:
            return
        idx, dest = int(selected[0]), int(selected[0]) + direction
        if 0 <= dest < len(job["sources"]):
            job["sources"][idx], job["sources"][dest] = job["sources"][dest], job["sources"][idx]
            self._show_job(self._current_index())
            self.source_tree.selection_set(str(dest))

    def _browse_template(self):
        path = self.filedialog.askopenfilename(filetypes=[("Excel 工作簿（无宏）", "*.xlsx")])
        if path:
            self._set_template_selected(path)
            try:
                prefix_length = int(self.vars["template_prefix_length"].get())
                matches = matching_profiles(path, prefix_length, PROFILE_DIR)
            except Exception as exc:
                self.messagebox.showerror(APP_NAME, str(exc))
                return
            if not matches:
                self.messagebox.showinfo(APP_NAME, f"没有已保存配置与模板名“{Path(path).stem}”前 {prefix_length} 个字符匹配。\n未自动载入其他配置；当前界面配置保持不变。")
            elif len(matches) == 1:
                name = matches[0]
                if self.messagebox.askyesno(APP_NAME, f"找到唯一匹配配置“{name}”。是否载入？"):
                    try:
                        self._load_profile_named(name, selected_template_path=path)
                    except Exception as exc:
                        self.messagebox.showerror(APP_NAME, f"载入匹配配置失败：\n{exc}")
            else:
                name = self._choose_profile_dialog(matches)
                if name:
                    try:
                        self._load_profile_named(name, selected_template_path=path)
                    except Exception as exc:
                        self.messagebox.showerror(APP_NAME, f"载入匹配配置失败：\n{exc}")

    def _browse_output_dir(self):
        path = self.filedialog.askdirectory()
        if path:
            self.vars["output_dir"].set(path)

    def _collect(self):
        merge_only = bool(self.merge_only_var.get())
        if not merge_only and not self._commit_job_editor(show_errors=False):
            raise ProcessingError("当前目标表的保留标题行数和粘贴起始行必须是整数。")
        settings = {key: var.get().strip() for key, var in self.vars.items()}
        settings["merge_only"] = merge_only
        settings["output_filename_custom"] = bool(self.output_filename_custom)
        settings["config_version"] = CONFIG_VERSION
        settings["config_name"] = self.active_profile_name
        settings["jobs"] = deepcopy(self.settings["jobs"])
        settings["merge_sources"] = deepcopy(self.settings.get("merge_sources", []))
        settings["workflow"] = deepcopy(self.settings.get("workflow", DEFAULT_WORKFLOW))
        settings["retain_merged_sources"] = bool(self.retain_merged_sources_var.get())
        settings["delete_processed_sources_after_success"] = bool(
            self.delete_processed_sources_after_success_var.get())
        for key in ("pre_mappings", "post_mappings", "region_xlsx_outputs", "text_outputs"):
            settings[key] = deepcopy(self.settings.get(key, []))
        for key in ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
            operation = deepcopy(self.settings.get(key, {"enabled": False, "rules": []}))
            operation["enabled"] = bool(self.operation_enabled[key].get())
            settings[key] = operation
        return normalize_settings(settings)

    def _save_default(self):
        try:
            settings = self._collect()
            save_settings(settings)
            self.settings = settings
            self.messagebox.showinfo(APP_NAME, f"默认配置已保存：\n{CONFIG_PATH}")
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, str(exc))

    def _new_profile(self):
        from tkinter import simpledialog
        name = simpledialog.askstring(APP_NAME, "新配置名称：", parent=self.root)
        if name is not None and name.strip():
            self._save_profile_named(name.strip())

    def _save_current_profile(self):
        name = self.active_profile_name or self.profile_choice.get().strip()
        if not name:
            self._new_profile()
            return
        self._save_profile_named(name)

    def _save_profile_named(self, name):
        try:
            settings = self._collect()
            settings["config_name"] = name
            validate_static_settings(settings)
            if name in list_profiles(PROFILE_DIR) and not self.messagebox.askyesno(APP_NAME, f"配置“{name}”已存在，是否覆盖保存？"):
                return
            path = save_profile(settings, name, PROFILE_DIR)
            self.settings = normalize_settings(settings)
            self.active_profile_name = name
            self.profile_choice.set(name)
            self._refresh_profile_choices()
            self.messagebox.showinfo(APP_NAME, f"已保存命名配置“{name}”：\n{path}")
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, str(exc))

    def _apply_settings_to_gui(self, settings, active_name=None):
        migration_notices = settings.get("_migration_notices") or configuration_migration_notes(settings)
        normalized = normalize_settings(settings)
        self.settings = normalized
        self.active_profile_name = str(active_name if active_name is not None else normalized.get("config_name", ""))
        self.output_filename_custom = bool(normalized.get("output_filename_custom", False))
        self.merge_only_var.set(bool(normalized.get("merge_only", False)))
        self._on_merge_mode_changed()
        self.retain_merged_sources_var.set(bool(normalized.get(
            "retain_merged_sources", normalized.get("retain_edited_sources", False))))
        self.delete_processed_sources_after_success_var.set(bool(
            normalized.get("delete_processed_sources_after_success", False)))
        self._suppress_output_filename_trace = True
        try:
            for key, var in self.vars.items():
                if key in normalized:
                    var.set(str(normalized.get(key, "")))
        finally:
            self._suppress_output_filename_trace = False
        for key, var in self.operation_enabled.items():
            var.set(bool(normalized.get(key, {}).get("enabled", False)))
        self.profile_choice.set(self.active_profile_name)
        self._render_jobs()
        self._render_merge_sources()
        self._render_area_configs()
        self._render_workflow()
        self._refresh_profile_choices()
        if migration_notices:
            self._show_configuration_migration_notices(migration_notices)

    def _load_profile_named(self, name, selected_template_path=None):
        settings = load_profile(name, PROFILE_DIR)
        if selected_template_path:
            settings["template_path"] = selected_template_path
        settings = normalize_settings(settings)
        validate_static_settings(settings)
        self._apply_settings_to_gui(settings, name)

    def _load_selected_profile(self):
        name = self.profile_choice.get().strip()
        if not name:
            self.messagebox.showinfo(APP_NAME, "当前没有可载入的命名配置。")
            return
        try:
            self._load_profile_named(name)
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, f"载入配置失败：\n{exc}")

    def _choose_profile_dialog(self, names):
        dialog = self.tk.Toplevel(self.root)
        dialog.title("选择匹配配置")
        dialog.transient(self.root)
        dialog.grab_set()
        self.ttk.Label(dialog, text="模板前缀匹配到多份配置，请选择：").pack(anchor="w", padx=10, pady=(10, 4))
        choice = self.tk.StringVar(value=names[0])
        combo = self.ttk.Combobox(dialog, textvariable=choice, values=names, state="readonly", width=32)
        combo.pack(padx=10, pady=4)
        result = {"name": None}
        buttons = self.ttk.Frame(dialog)
        buttons.pack(pady=8)
        def accept():
            result["name"] = choice.get()
            dialog.destroy()
        self.ttk.Button(buttons, text="载入所选", command=accept).pack(side="left", padx=4)
        self.ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=4)
        dialog.wait_window()
        return result["name"]

    def _export_config(self):
        try:
            settings = self._collect()
            validate_static_settings(settings)
            path = self.filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON 配置", "*.json")], initialfile="wps_sheet_config.json")
            if path:
                save_settings(settings, path)
                self.settings = settings
                self._render_jobs(self._current_index())
                self.messagebox.showinfo(APP_NAME, f"配置已导出：\n{path}")
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, str(exc))

    def _import_config(self):
        path = self.filedialog.askopenfilename(filetypes=[("JSON 配置", "*.json")])
        if not path:
            return
        try:
            settings = load_settings(path)
            validate_static_settings(settings)
            self._apply_settings_to_gui(settings)
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, f"加载配置失败：\n{exc}")

    def _reset_config(self):
        if not self.messagebox.askyesno(APP_NAME, "恢复程序内置示例配置？不会删除已保存或导出的配置文件。"):
            return
        self._apply_settings_to_gui(deepcopy(DEFAULT_SETTINGS), active_name="")

    def _open_settings(self):
        dialog = self.tk.Toplevel(self.root)
        dialog.title("设置")
        dialog.geometry("840x690")
        dialog.minsize(650, 500)
        dialog.transient(self.root)
        notebook = self.ttk.Notebook(dialog)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        license_page = self.ttk.Frame(notebook, padding=12)
        help_page = self.ttk.Frame(notebook, padding=8)
        config_page = self.ttk.Frame(notebook, padding=12)
        notebook.add(license_page, text="秘钥管理")
        notebook.add(help_page, text="使用说明")
        notebook.add(config_page, text="配置管理")
        self._build_license_settings_tab(license_page)
        self._build_usage_help_tab(help_page)
        self._build_profile_management_tab(config_page)

    def _build_license_settings_tab(self, parent):
        from datetime import date
        from licensing import (LicensingError, evaluate_authorization, install_license,
                               installed_license_path, inspect_license_expiry, machine_code)

        ttk = self.ttk
        parent.columnconfigure(0, weight=1)
        current_machine = machine_code()
        ttk.Label(parent, text="本机机器码", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        machine_var = self.tk.StringVar(value=current_machine)
        machine_row = ttk.Frame(parent)
        machine_row.grid(row=1, column=0, sticky="ew")
        machine_row.columnconfigure(0, weight=1)
        ttk.Entry(machine_row, textvariable=machine_var, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        status_var = self.tk.StringVar(value="正在读取授权状态…")
        ttk.Button(machine_row, text="复制机器码", command=lambda: self._copy_machine_code(current_machine)).grid(
            row=0, column=1)
        ttk.Label(parent, text="当前授权", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=(14, 4))
        ttk.Label(parent, textvariable=status_var, wraplength=760, justify="left").grid(
            row=3, column=0, sticky="w")

        def refresh_status():
            lines = []
            try:
                authorization = evaluate_authorization()
                if authorization.state == "licensed" and authorization.license_expiry:
                    lines.append(f"授权有效，到期日期：{authorization.license_expiry.isoformat()}（含当日）。")
                elif authorization.state == "trial":
                    lines.append(f"当前为试用状态，剩余 {authorization.trial_remaining_days} 天；试用截止：{authorization.trial_expiry.isoformat()}。")
                else:
                    lines.append(authorization.detail or "当前授权不可用。")
            except Exception as exc:
                lines.append(f"读取授权状态失败：{exc}")
            key_path = installed_license_path()
            if key_path.is_file():
                try:
                    expiry = inspect_license_expiry(key_path.read_text(encoding="utf-8"), current_machine)
                    validity = "已过期" if expiry < date.today() else "有效"
                    lines.append(f"已安装秘钥：{validity}；签名绑定本机，到期日期：{expiry.isoformat()}。")
                except (OSError, UnicodeError, LicensingError) as exc:
                    lines.append(f"已检测到秘钥文件，但无法验证其签名或机器码：{exc}")
            else:
                lines.append("当前没有已安装的秘钥文件。")
            status_var.set("\n".join(lines))

        ttk.Separator(parent).grid(row=4, column=0, sticky="ew", pady=14)
        ttk.Label(parent, text="更新 / 修改秘钥", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=5, column=0, sticky="w", pady=(0, 4))
        ttk.Label(parent, text="粘贴为本机签发的新秘钥。只有签名、产品和机器码校验通过且尚未到期时，才会替换当前秘钥。",
                  wraplength=760, justify="left").grid(row=6, column=0, sticky="w", pady=(0, 6))
        token_text = self.tk.Text(parent, height=5, wrap="word")
        token_text.grid(row=7, column=0, sticky="ew")
        ttk.Button(parent, text="验证并更新秘钥", command=lambda: self._update_license_key(
            token_text, current_machine, refresh_status, parent)).grid(row=8, column=0, sticky="w", pady=8)
        ttk.Label(parent, text="秘钥保存在当前 Windows 用户的应用数据目录中；主程序只含验证公钥，不含签发私钥。",
                  foreground="#4b5563", wraplength=760, justify="left").grid(row=9, column=0, sticky="w", pady=(8, 0))
        refresh_status()

    def _copy_machine_code(self, machine):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(machine)
            self.root.update_idletasks()
            self.messagebox.showinfo(APP_NAME, "本机机器码已复制。")
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, f"复制机器码失败：{exc}")

    def _update_license_key(self, token_widget, machine, refresh_status, parent):
        from licensing import LicensingError, install_license

        token = token_widget.get("1.0", "end").strip()
        if not token:
            self.messagebox.showinfo(APP_NAME, "请先粘贴新秘钥。", parent=parent.winfo_toplevel())
            return
        try:
            expiry = install_license(token, machine=machine)
        except (LicensingError, OSError) as exc:
            self.messagebox.showerror(APP_NAME, f"秘钥未更新：{exc}", parent=parent.winfo_toplevel())
            return
        token_widget.delete("1.0", "end")
        refresh_status()
        self.messagebox.showinfo(APP_NAME, f"秘钥已验证并更新，到期日期：{expiry.isoformat()}。", parent=parent.winfo_toplevel())

    def _build_usage_help_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        text = self.tk.Text(parent, wrap="word", padx=10, pady=8)
        scrollbar = self.ttk.Scrollbar(parent, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        guide = """WPS 表格多源汇总处理器｜功能使用说明

一、模板、输出与源数据列
选择无宏 .xlsx 模板、输出目录和输出文件名。源数据复制起始列/末列用于定义从源表复制的列范围；有关联源文件时末列必填，无源文件仅编辑模板时可留空。模板前缀匹配长度用于选择模板时查找已保存的命名配置；只有唯一匹配或明确选择后才载入。

二、目标子表与多源文件
每个目标子表单独设置模板表名、标题行、粘贴起始行和公式列范围。多份源文件先按列表顺序合并为该目标表的工作簿，再对合并结果执行源规则；命中源规则时处理结果必须先另存为独立副本，模板从该副本读取。首份源文件保留复制范围首行，后续源文件可分别选择是否保留复制范围首行（默认保留）。合并源文件名可自定义，留空默认为对应模板子表名。无源侧规则时，“保存合并源工作簿”控制是否保留合并副本。可选勾选“模板处理全部成功后删除已另存的源副本”；默认保留，开启后只在模板、区域输出及文本读取均成功后删除，任何失败都保留副本供排查。源工作表名留空时取首个工作表，复制结束行留空时按首个整行空白边界判断。无源目标表保留模板原数据，仅运行配置的模板规则。

三、结果映射、前映射与后映射
结果映射在每个目标表内配置：首次完整重算后，把指定公式结果列复制为值或文本到公式范围之外的目标列；结果列须在公式范围内。前映射在复制主数据前写入模板已有区域；后映射在首次重算和结果映射之后复制任意源/目标区域。结果映射和后映射彼此独立，均可不配置。

四、公式填充与工作表范围
目标表“公式列范围”按本轮数据行数自动填公式。区域“公式填充”可选择源文件或目标工作簿，使用自定义公式或同表公式种子单元格，并由 WPS 按相对引用填满范围；它会覆盖所选区域原内容。处理规则的工作表名留空时默认第一张表，名称匹配不区分大小写。区域筛选范围仅留空时表示整表；A/a 表示整列，A1/a1 表示单格 A1，A1:F100 表示矩形。其他常见范围：D1=整列，A5=整行，B3,B3=单格，A1:F100=矩形。规则按 UsedRange 裁剪并设有单条单元格上限。

五、区域处理功能
“文本分列”按分隔符拆分指定单列并选择输出起始列及各列格式，输出区域须为空。“区域筛选”按列和值筛选记录，只压缩所选区域，不删除工作表物理行。“去重”按整行或指定关键列去重，保留首次出现记录。“格式转换”将范围转换为文本、数字或日期。规则可指定源文件或目标工作簿；源端公式填充、文本分列、筛选、去重和格式转换均在源文件合并后、模板编辑前对合并临时副本执行，处理结果先另存为独立副本，模板从该副本读取；原始文件不会被改写。默认保留已另存副本，可选在模板处理全部成功后删除；失败时会保留副本供排查。

六、区域 XLSX 输出与文本预览
“区域 XLSX 输出”添加模板表和范围；区域 XLSX 会独立导出并保留单元格样式。“文本预览/复制”添加工作表范围；处理并保存成功后显示预览，可复制完整区域或分块内容，列之间以 Tab 分隔。

七、模块工作流
“模块工作流”可调整处理模块顺序；添加/移除可选模块并用上移/下移排序。必需模块、结果映射、后映射、公式填充、区域操作和输出模块之间有依赖约束；非法顺序会在保存或运行时提示。默认顺序先合并源文件，再处理合并副本和模板，完成重算、映射、区域操作和保存，最后输出区域 XLSX 及文本预览。

八、独立源文件合并
勾选“启用独立合并模式”，仅使用本页独立源文件列表和“合并结果文件名”；模板、目标表、源规则、映射、区域操作及模块工作流均不会运行。首份源可为 .csv 或 .xlsx，后续源必须为 .xlsx；CSV 自动尝试 UTF-8/UTF-8 BOM、GB18030 与常见分隔符。XLSX 工作表名留空时使用首表。首份复制范围首行固定保留，后续可逐份选择保留或跳过；结果始终另存为 XLSX，不覆盖源文件或已有输出。

九、配置管理与秘钥管理
“保存配置…”导出当前界面配置为 JSON，“加载配置…”导入 JSON；命名配置可新建、保存、载入，并在“设置 → 配置管理”导出当前配置或所选命名配置，也可删除明确选中的命名配置。删除命名配置不会删除当前正在编辑的设置或其他配置文件。
“设置 → 秘钥管理”显示本机机器码、当前授权状态和当前秘钥到期日期；粘贴新秘钥后点击“验证并更新秘钥”，仅通过本机绑定、签名和有效期验证的秘钥才会替换现有秘钥。更新秘钥不会更改程序内置公钥。

十、数据安全提示
主工作簿和原始源文件不会被直接覆盖；模板处理使用安全临时副本，独立合并以只读方式读取 CSV/XLSX 并将结果写入新的 XLSX。命中源侧规则时，处理结果另存为独立源副本并作为模板输入；默认保留，可选在模板处理全部成功后删除。发生失败时保留副本供排查。无源侧规则时，合并源副本由保留选项控制。秘钥与业务配置分别存放；主程序不包含签名私钥。恢复默认只替换当前界面配置，不删除命名配置或已导出的 JSON 文件。
"""
        text.insert("1.0", guide)
        text.configure(state="disabled")

    def _build_profile_management_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        self.ttk.Label(parent, text="导出当前配置会保存当前界面内容为 JSON；命名配置是保存在本机的可复用配置档。",
                       wraplength=760, justify="left").grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.ttk.Button(parent, text="导出当前配置…", command=self._export_config).grid(
            row=1, column=0, sticky="w", pady=(0, 14))
        profile_box = self.ttk.LabelFrame(parent, text="命名配置", padding=10)
        profile_box.grid(row=2, column=0, sticky="new")
        profile_box.columnconfigure(0, weight=1)
        choice = self.tk.StringVar(value="")
        combo = self.ttk.Combobox(profile_box, textvariable=choice, state="readonly", width=42)
        combo.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=4)

        def refresh_profiles():
            names = list_profiles(PROFILE_DIR)
            combo.configure(values=names)
            if choice.get() not in names:
                choice.set(names[0] if names else "")
            return names

        def export_selected_profile():
            name = choice.get().strip()
            if not name:
                self.messagebox.showinfo(APP_NAME, "请先选择要导出的命名配置。", parent=parent.winfo_toplevel())
                return
            try:
                profile = load_profile(name, PROFILE_DIR)
                profile = normalize_settings(profile)
                validate_static_settings(profile)
                destination = self.filedialog.asksaveasfilename(
                    parent=parent.winfo_toplevel(), defaultextension=".json",
                    filetypes=[("JSON 配置", "*.json")], initialfile=f"{name}.json")
                if destination:
                    save_settings(profile, destination)
                    self.messagebox.showinfo(APP_NAME, f"命名配置已导出：\n{destination}", parent=parent.winfo_toplevel())
            except Exception as exc:
                self.messagebox.showerror(APP_NAME, f"导出命名配置失败：\n{exc}", parent=parent.winfo_toplevel())

        def delete_selected_profile():
            name = choice.get().strip()
            if not name:
                self.messagebox.showinfo(APP_NAME, "请先选择要删除的命名配置。", parent=parent.winfo_toplevel())
                return
            if not self.messagebox.askyesno(APP_NAME, f"确定删除命名配置“{name}”吗？此操作只删除该配置档文件，不影响当前界面内容。",
                                            parent=parent.winfo_toplevel()):
                return
            try:
                delete_profile(name, PROFILE_DIR)
                if getattr(self, "active_profile_name", "") == name:
                    self.active_profile_name = ""
                    self.settings["config_name"] = ""
                    self.profile_choice.set("")
                refresh_profiles()
                self._refresh_profile_choices()
                self.messagebox.showinfo(APP_NAME, f"已删除命名配置“{name}”。", parent=parent.winfo_toplevel())
            except Exception as exc:
                self.messagebox.showerror(APP_NAME, f"删除命名配置失败：\n{exc}", parent=parent.winfo_toplevel())

        buttons = self.ttk.Frame(profile_box)
        buttons.grid(row=1, column=0, sticky="w", pady=6)
        self.ttk.Button(buttons, text="刷新列表", command=refresh_profiles).pack(side="left", padx=(0, 5))
        self.ttk.Button(buttons, text="导出所选配置…", command=export_selected_profile).pack(side="left", padx=5)
        self.ttk.Button(buttons, text="删除所选配置…", command=delete_selected_profile).pack(side="left", padx=5)
        self.ttk.Label(profile_box, text=f"本机配置档目录：{PROFILE_DIR}", foreground="#4b5563",
                       wraplength=740, justify="left").grid(row=2, column=0, sticky="w", pady=(7, 0))
        refresh_profiles()

    def _log(self, message: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _copy_preview_to_clipboard(self, text: str):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update_idletasks()
            self._log("已将文本区域复制到系统剪贴板。")
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, f"写入 Windows 剪贴板失败：{exc}")

    def _render_text_previews(self, previews: list[dict[str, Any]]):
        self.text_previews = deepcopy(previews)
        for tab in self.preview_notebook.tabs():
            self.preview_notebook.forget(tab)
        if not previews:
            self.preview_notebook.pack_forget()
            self.preview_empty_label.pack(anchor="w", padx=4, pady=4)
            return
        self.preview_empty_label.pack_forget()
        self.preview_notebook.pack(fill="both", expand=True)
        for idx, preview in enumerate(previews, 1):
            page = self.ttk.Frame(self.preview_notebook, padding=5)
            page.columnconfigure(0, weight=1)
            page.rowconfigure(2, weight=1)
            resolved = preview.get("resolved_range", preview["range"])
            range_text = f"{preview['range']}（实际读取 {resolved}）" if resolved != preview["range"] else preview["range"]
            self.ttk.Label(page, text=f"{preview['sheet']}!{range_text}").grid(row=0, column=0, sticky="w", padx=4, pady=3)
            chunks = preview.get("chunks") or [{"start_row": 1, "end_row": 1, "text": preview["text"]}]
            chunk_box = None
            if len(chunks) > 1:
                labels = [f"第 {n} 块（行 {chunk['start_row']}-{chunk['end_row']}）" for n, chunk in enumerate(chunks, 1)]
                chunk_var = self.tk.StringVar(value=labels[0])
                chunk_box = self.ttk.Combobox(page, textvariable=chunk_var, values=labels, state="readonly", width=32)
                chunk_box.grid(row=1, column=0, sticky="w", padx=4, pady=3)
            text_widget = self.tk.Text(page, height=5, wrap="none")
            text_widget.insert("1.0", chunks[0]["text"])
            text_widget.configure(state="disabled")
            text_widget.grid(row=2, column=0, sticky="nsew", padx=4, pady=3)
            if chunk_box is not None:
                def show_chunk(_event=None, box=chunk_box, items=chunks, widget=text_widget):
                    selected_index = max(0, box.current())
                    widget.configure(state="normal")
                    widget.delete("1.0", "end")
                    widget.insert("1.0", items[selected_index]["text"])
                    widget.configure(state="disabled")
                chunk_box.bind("<<ComboboxSelected>>", show_chunk)
            def copy_current_chunk(items=chunks, box=chunk_box):
                selected_index = max(0, box.current()) if box is not None else 0
                self._copy_preview_to_clipboard(items[selected_index]["text"])
            copy_label = f"复制第 {idx} 个区域的当前分块" if len(chunks) > 1 else f"复制第 {idx} 个区域"
            self.ttk.Button(page, text=copy_label, command=copy_current_chunk).grid(row=3, column=0, sticky="w", padx=4, pady=4)
            self.preview_notebook.add(page, text=f"{idx}. {preview['sheet']}!{preview['range']}")

    def _run(self):
        try:
            settings = self._collect()
            validate_static_settings(settings)
            missing_source = None if settings.get("merge_only") else _first_missing_source_location(settings)
            if missing_source is not None:
                self._focus_source_location(*missing_source)
                raise ProcessingError(_missing_source_path_message(settings, missing_source))
            save_settings(settings)
            self.settings = settings
        except Exception as exc:
            self.messagebox.showerror(APP_NAME, str(exc))
            return
        self.run_button.configure(state="disabled")
        self.progress.start(10)
        self._render_text_previews([])
        self._region_output_files = []
        self._source_output_files = []
        self._log("开始独立合并……" if settings.get("merge_only") else "开始处理……")
        def work():
            try:
                path = process_workbooks(settings, lambda text: self.events.put(("log", text)),
                                         lambda previews: self.events.put(("text_previews", previews)),
                                         lambda paths: self.events.put(("region_outputs", paths)),
                                         lambda paths: self.events.put(("source_outputs", paths)))
                self.events.put(("success", (str(path), bool(settings.get("merge_only")))))
            except Exception as exc:
                self.events.put(("error", f"{exc}\n\n详细诊断：\n{traceback.format_exc()}"))
        threading.Thread(target=work, daemon=True).start()

    def _drain_events(self):
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "log":
                    self._log(str(payload))
                elif event == "text_previews":
                    self._render_text_previews(payload)
                elif event == "region_outputs":
                    self._region_output_files = [Path(path) for path in payload]
                elif event == "source_outputs":
                    self._source_output_files.extend(Path(path) for path in payload)
                elif event in ("success", "error"):
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                    if event == "success":
                        path_text, merge_only = payload if isinstance(payload, tuple) else (str(payload), False)
                        self._log(f"合并完成：{path_text}" if merge_only else f"完成：{path_text}")
                        xlsx_count = sum(path.suffix.lower() == ".xlsx" for path in self._region_output_files)
                        output_note = f"\n区域 XLSX：{xlsx_count} 个。" if self._region_output_files else ""
                        if self._source_output_files:
                            output_note += f"\n已保存合并源工作簿：{len(self._source_output_files)} 个。"
                        if merge_only:
                            self.messagebox.showinfo(APP_NAME, f"独立合并成功，已保存 XLSX：\n{path_text}")
                        else:
                            self.messagebox.showinfo(APP_NAME, f"处理成功，已保存新文件：\n{path_text}{output_note}")
                    else:
                        self._log("处理未全部完成；请查看详细诊断中的已保存文件路径及清理信息。")
                        self.messagebox.showerror(APP_NAME, str(payload))
        except queue.Empty:
            pass
        self.root.after(150, self._drain_events)


def self_check() -> int:
    try:
        validate_static_settings(DEFAULT_SETTINGS)
        expected = {"EB:EN": (132, 144), "EB:EU": (132, 151), "EB:FV": (132, 178)}
        for text, answer in expected.items():
            if parse_column_range(text) != answer:
                raise AssertionError(f"列范围换算错误：{text}")
        legacy = {"sources": ["1.xlsx", "2.xlsx", "3.xlsx"], "source_start_row": 2,
                  "jobs": [{"template_sheet": "一", "source_sheet": "源", "header_rows": 2, "start_row": 3, "formula_range": "EB:EN", "result_col": "EJ", "target_col": "DG"},
                           {"template_sheet": "二", "source_sheet": "源", "header_rows": 2, "start_row": 3, "formula_range": "EB:EU", "result_col": "EQ", "target_col": "DG"},
                           {"template_sheet": "三", "source_sheet": "源", "header_rows": 3, "start_row": 4, "formula_range": "EB:FV", "result_col": "FQ", "target_col": "DG"}]}
        if len(normalize_settings(legacy)["jobs"]) != 3:
            raise AssertionError("v1 配置迁移失败")
        print(f"静态检查通过：v{CONFIG_VERSION} 默认配置有效；公式范围/映射、两阶段映射、区域公式填充和四类区域处理配置、源/目标处理范围及旧配置迁移均通过。")
        print("未连接 WPS COM；未读取真实业务文件；未验证 Windows GUI 或 EXE。")
        return 0
    except Exception as exc:
        print(f"静态检查失败：{exc}", file=sys.stderr)
        return 1


def _startup_log_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_STATE_HOME") or Path.home() / ".local" / "state")
    return base / "WpsSheetAutomation" / "logs" / "startup.log"


def _write_startup_log(stage: str, detail: str = "") -> Path | None:
    path = _startup_log_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as stream:
            stream.write(f"\n[{datetime.now().astimezone().isoformat(timespec='seconds')}] {stage}\n")
            if detail:
                stream.write(detail.rstrip() + "\n")
        return path
    except OSError:
        return None


def main() -> int:
    if "--self-check" in sys.argv:
        return self_check()
    if sys.platform != "win32":
        print("此程序的公式重算与最终保存要求 Windows 版 WPS 表格 COM。请在 Windows 11 + WPS 表格环境运行。")
        print("当前执行环境不是 Windows；没有生成或验证 EXE。可运行 python main.py --self-check 和假数据单元测试。")
        return 2
    root = None
    stage = "导入 Tk 图形界面组件"
    _write_startup_log("启动开始")
    try:
        import tkinter as tk
        stage = "创建 Tk 主窗口"
        root = tk.Tk()
        root.withdraw()
        _write_startup_log("主窗口已创建，开始授权检查")
        stage = "显示授权窗口"
        from licensing import authorize_startup
        if not authorize_startup(root):
            root.destroy()
            _write_startup_log("用户关闭授权窗口，程序退出")
            return 0
        _write_startup_log("授权通过，开始构建主界面")
        stage = "构建主界面"
        App(root)
        root.deiconify()
        _write_startup_log("主界面已显示")
        stage = "运行主界面消息循环"
        root.mainloop()
        return 0
    except Exception as exc:
        detail = traceback.format_exc()
        log_path = _write_startup_log(f"{stage}失败：{exc}", detail)
        message = f"程序启动失败（{stage}）：\n{exc}"
        if log_path is not None:
            message += f"\n\n诊断日志：\n{log_path}"
        try:
            if root is not None:
                root.deiconify()
                root.lift()
            from tkinter import messagebox
            messagebox.showerror(APP_NAME, message, parent=root)
        except Exception:
            pass
        print(message + "\n\n" + detail, file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
