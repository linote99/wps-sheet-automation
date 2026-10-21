"""Validation and grouping helpers for per-rule source-file selections.

This module intentionally has no dependency on the GUI or WPS automation code.
"""
from __future__ import annotations

import os
import re
from copy import deepcopy
from typing import Any

OPERATIONS = (
    "CustomFormula",
    "TextToColumns",
    "RegionFilter",
    "Deduplicate",
    "FormatConvert",
)
_WINDOWS_INVALID = set('<>:"/\\|?*')
_RESERVED_DEVICE = re.compile(r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$", re.IGNORECASE)


def _has_control(value: str) -> bool:
    return any(ord(char) < 32 or ord(char) == 127 for char in value)


def _absolute_path(value: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(os.path.expanduser(value))))


def _normalise_output_filename(value: Any, source_paths: list[str]) -> str:
    if value is None or value == "":
        if not source_paths:
            # This value is not used by legacy rules, but keeping it deterministic
            # makes validation useful on its own.
            return "_已处理.xlsx"
        stem = os.path.splitext(os.path.basename(source_paths[0]))[0]
        return _normalise_output_filename(f"{stem}_已处理.xlsx", [])
    if not isinstance(value, str):
        raise ValueError("source_output_filename 必须是字符串或空值。")
    if not value or value.strip() == "":
        if not source_paths:
            return "_已处理.xlsx"
        stem = os.path.splitext(os.path.basename(source_paths[0]))[0]
        return _normalise_output_filename(f"{stem}_已处理.xlsx", [])
    if _has_control(value):
        raise ValueError("source_output_filename 不得包含控制字符。")
    if "/" in value or "\\" in value:
        raise ValueError("source_output_filename 只能是文件名，不能包含路径。")
    if any(char in _WINDOWS_INVALID for char in value):
        raise ValueError("source_output_filename 包含 Windows 无效字符。")
    if value[-1] in " .":
        raise ValueError("source_output_filename 不得以空格或点结尾。")
    has_xlsx_suffix = value.lower().endswith(".xlsx")
    if has_xlsx_suffix:
        stem = value[:-5]
    else:
        stem = value
    if not stem.strip(" ."):
        raise ValueError("source_output_filename 的文件名主体不能为空。")
    if _RESERVED_DEVICE.match(stem):
        raise ValueError("source_output_filename 不能使用 Windows 保留设备名。")
    # A name with a non-xlsx extension is not an xlsx output name.
    if not has_xlsx_suffix and "." in stem:
        raise ValueError("source_output_filename 只能省略扩展名或使用 .xlsx。")
    if len((stem + ".xlsx").encode("utf-16-le")) // 2 > 240:
        raise ValueError("source_output_filename 过长，请缩短以保留同名避让序号空间。")
    return f"{stem}.xlsx"


def validate_source_selection(rule: dict[str, Any], label: str = "源文件规则") -> None:
    """Validate one rule's source selection, raising :class:`ValueError`.

    Missing fields retain legacy behavior.  Validation deliberately does not
    check whether selected files exist.
    """
    if not isinstance(rule, dict):
        raise ValueError(f"{label} 必须是对象。")
    scope = str(rule.get("scope", "target")).strip().lower()
    raw_paths = rule.get("source_paths", [])
    if raw_paths is None:
        raise ValueError(f"{label} source_paths 必须是列表。")
    if not isinstance(raw_paths, list):
        raise ValueError(f"{label} source_paths 必须是列表。")
    if scope != "source" and raw_paths:
        raise ValueError(f"{label} 只有 scope=source 时才能选择源文件。")
    normalised: list[str] = []
    seen: set[str] = set()
    for index, path in enumerate(raw_paths, 1):
        if not isinstance(path, str) or not path.strip():
            raise ValueError(f"{label} source_paths 第 {index} 项必须是非空字符串。")
        if _has_control(path):
            raise ValueError(f"{label} source_paths 第 {index} 项不得包含控制字符。")
        normal = _absolute_path(path)
        if normal in seen:
            raise ValueError(f"{label} source_paths 不能包含重复规范化路径。")
        seen.add(normal)
        normalised.append(normal)
        if scope == "source" and not path.lower().endswith(".xlsx"):
            raise ValueError(f"{label} 选定源文件必须是 .xlsx。")
    _normalise_output_filename(rule.get("source_output_filename"), normalised)


def plan_selected_source_groups(settings: dict[str, Any]) -> list[dict[str, Any]]:
    """Plan enabled source rules into stable groups without mutating settings."""
    if not isinstance(settings, dict):
        raise ValueError("settings 必须是对象。")
    groups: dict[tuple[tuple[str, ...], str], dict[str, Any]] = {}
    for operation_key in OPERATIONS:
        operation = settings.get(operation_key, {})
        if operation is None:
            continue
        if not isinstance(operation, dict):
            raise ValueError(f"{operation_key} 必须是对象。")
        rules = operation.get("rules", [])
        if rules is None:
            rules = []
        if not isinstance(rules, list):
            raise ValueError(f"{operation_key}.rules 必须是列表。")
        for rule in rules:
            validate_source_selection(rule, f"{operation_key} 规则")
        if not operation.get("enabled", False):
            continue
        for rule in rules:
            if str(rule.get("scope", "target")).strip().lower() != "source" or not rule.get("source_paths", []):
                continue
            paths = [_absolute_path(path) for path in rule["source_paths"]]
            output = _normalise_output_filename(rule.get("source_output_filename"), paths)
            key = (tuple(paths), output)
            if key not in groups:
                groups[key] = {"source_paths": list(paths), "output_filename": output, "rules": {}}
            grouped = groups[key]["rules"].setdefault(operation_key, [])
            grouped.append(deepcopy(rule))
    return list(groups.values())
