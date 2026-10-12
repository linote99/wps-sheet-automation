"""Workflow helpers for WPS Sheet Automation.

Keep this project module's name distinct from the unrelated third-party
``workflow`` distribution handled by PyInstaller's contrib hook.
"""

from __future__ import annotations

import hashlib
import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Any


MODULE_LABELS = {
    "source_merge": "源文件合并",
    "source_preprocess": "源文件规则处理",
    "template_cleanup": "模板清理",
    "pre_mapping": "前映射",
    "source_copy": "多源数据复制",
    "formula_fill": "公式填充",
    "formula_recalc": "首次完整重算",
    "result_mapping": "结果映射",
    "post_mapping": "后映射",
    "CustomFormula": "区域公式填充",
    "TextToColumns": "文本分列",
    "RegionFilter": "区域筛选",
    "Deduplicate": "去重",
    "FormatConvert": "格式转换",
    "final_recalc_save": "最终重算并保存",
    "region_xlsx_output": "区域 XLSX 输出",
    "text_output": "文本区域剪贴板输出",
}

DEFAULT_WORKFLOW = list(MODULE_LABELS)
CORE_MODULES = {
    "source_merge", "source_preprocess", "template_cleanup", "source_copy", "formula_fill", "formula_recalc",
    "final_recalc_save",
}

# The UI may reorder independent region processors and terminal readers. These
# edges encode only correctness constraints shared by every valid workflow.
DEPENDENCIES = (
    ("source_merge", "source_preprocess", "源文件必须先按顺序合并，再执行其他源文件处理"),
    ("source_preprocess", "template_cleanup", "源文件规则必须先处理完成，再开始编辑模板"),
    ("template_cleanup", "source_copy", "模板清理必须先于源数据写入，否则可能清掉刚复制的数据"),
    ("source_copy", "formula_fill", "源数据复制必须先完成，才能按最终数据行数填充公式"),
    ("formula_fill", "formula_recalc", "公式填充必须先于首次完整重算"),
    ("formula_recalc", "final_recalc_save", "公式必须先完整重算，再最终重算并保存"),
    ("final_recalc_save", "region_xlsx_output", "区域 XLSX 必须在工作簿最终重算并保存后导出"),
    ("final_recalc_save", "text_output", "文本必须在工作簿最终重算并保存后读取"),
)


def validate_workflow(settings: dict[str, Any]) -> None:
    workflow = settings.get("workflow", DEFAULT_WORKFLOW)
    if not isinstance(workflow, list):
        raise ValueError("工作流 workflow 必须是模块 ID 数组。")
    unknown = [item for item in workflow if not isinstance(item, str) or item not in MODULE_LABELS]
    if unknown:
        raise ValueError(f"工作流包含未知模块：{unknown[0]!r}。")
    if len(set(workflow)) != len(workflow):
        raise ValueError("工作流中不能重复添加同一模块。")
    missing_core = CORE_MODULES.difference(workflow)
    if missing_core:
        module_id = sorted(missing_core)[0]
        raise ValueError(f"不能移除必需模块“{MODULE_LABELS[module_id]}”；处理结果依赖此步骤。")

    required = []
    if settings.get("pre_mappings"):
        required.append(("pre_mapping", "配置了前映射规则"))
    if settings.get("post_mappings"):
        required.append(("post_mapping", "配置了后映射规则"))
    if any(isinstance(job, dict) and job.get("mappings") for job in settings.get("jobs", [])):
        required.append(("result_mapping", "配置了结果映射规则"))
    operation_map = (("CustomFormula", "CustomFormula"), ("TextToColumns", "TextToColumns"),
                     ("RegionFilter", "RegionFilter"), ("Deduplicate", "Deduplicate"),
                     ("FormatConvert", "FormatConvert"))
    for key, module_id in operation_map:
        operation = settings.get(key, {})
        if isinstance(operation, dict) and operation.get("enabled", False):
            required.append((module_id, f"已启用 {MODULE_LABELS[module_id]}"))
    if settings.get("text_outputs"):
        required.append(("text_output", "配置了文本输出区域"))
    if settings.get("region_xlsx_outputs"):
        required.append(("region_xlsx_output", "配置了区域 XLSX 输出"))
    for module_id, reason in required:
        if module_id not in workflow:
            raise ValueError(f"{reason}，但工作流中移除了“{MODULE_LABELS[module_id]}”模块；请重新添加该模块或清空/禁用对应配置。")

    positions = {module_id: index for index, module_id in enumerate(workflow)}
    edges = list(DEPENDENCIES)
    edges.extend((
        ("template_cleanup", "pre_mapping", "前映射需在模板旧内容清理后应用"),
        ("pre_mapping", "source_copy", "前映射必须在主源数据复制前完成"),
        ("formula_recalc", "result_mapping", "结果映射必须读取首次完整重算后的公式结果"),
        ("result_mapping", "post_mapping", "后映射应在结果映射之后执行"),
        ("result_mapping", "final_recalc_save", "结果映射必须在最终重算和保存前完成"),
        ("formula_recalc", "post_mapping", "后映射必须在首次完整重算后应用"),
        ("post_mapping", "final_recalc_save", "后映射必须在最终重算和保存前完成"),
    ))
    region_modules = ("CustomFormula", "TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert")
    for module_id in region_modules:
        edges.append(("formula_recalc", module_id, f"{MODULE_LABELS[module_id]}必须在首次完整重算后执行"))
        edges.append((module_id, "final_recalc_save", f"{MODULE_LABELS[module_id]}必须在最终重算和保存之前执行"))
        if "result_mapping" in positions:
            edges.append(("result_mapping", module_id, f"{MODULE_LABELS[module_id]}必须在结果映射之后执行"))
        if "post_mapping" in positions:
            edges.append(("post_mapping", module_id, f"{MODULE_LABELS[module_id]}必须在后映射之后执行"))
    for module_id in ("TextToColumns", "RegionFilter", "Deduplicate", "FormatConvert"):
        edges.append(("CustomFormula", module_id, f"{MODULE_LABELS[module_id]}必须在区域公式填充后执行"))
    for before, after, reason in edges:
        if before in positions and after in positions and positions[before] >= positions[after]:
            raise ValueError(f"工作流顺序无效：“{MODULE_LABELS[before]}”必须位于“{MODULE_LABELS[after]}”之前；{reason}。")


def add_workflow_module(workflow: list[str], module_id: str) -> list[str]:
    if module_id not in MODULE_LABELS:
        raise ValueError(f"未知工作流模块：{module_id!r}。")
    if module_id in workflow:
        raise ValueError(f"模块“{MODULE_LABELS[module_id]}”已在工作流中。")
    candidate = list(workflow)
    rank = list(MODULE_LABELS).index(module_id)
    insert_at = next((index for index, current in enumerate(candidate)
                      if current in MODULE_LABELS and list(MODULE_LABELS).index(current) > rank), len(candidate))
    candidate.insert(insert_at, module_id)
    return candidate


def move_workflow_module(workflow: list[str], index: int, direction: int) -> list[str]:
    candidate = list(workflow)
    destination = index + direction
    if not 0 <= index < len(candidate) or not 0 <= destination < len(candidate):
        return candidate
    candidate[index], candidate[destination] = candidate[destination], candidate[index]
    return candidate


def remove_workflow_module(workflow: list[str], index: int) -> list[str]:
    candidate = list(workflow)
    if 0 <= index < len(candidate):
        candidate.pop(index)
    return candidate


def profile_filename(name: str) -> str:
    clean_name = str(name).strip()
    if not clean_name:
        raise ValueError("配置名称不能为空。")
    if len(clean_name) > 80 or any(ord(ch) < 32 for ch in clean_name):
        raise ValueError("配置名称最多 80 个字符，且不能包含控制字符。")
    slug = re.sub(r"[^\w.-]+", "_", clean_name, flags=re.UNICODE).strip("._-") or "config"
    digest = hashlib.sha256(clean_name.casefold().encode("utf-8")).hexdigest()[:12]
    return f"{slug[:60]}-{digest}.json"


def list_profiles(profile_dir: Path | str) -> list[str]:
    directory = Path(profile_dir)
    profiles: list[tuple[str, Path]] = []
    if not directory.exists():
        return []
    for path in directory.glob("*.json"):
        try:
            content = json.loads(path.read_text(encoding="utf-8"))
            name = str(content.get("config_name") or "").strip()
            if name:
                profiles.append((name, path))
        except (OSError, ValueError, AttributeError):
            continue
    return sorted((name for name, _ in profiles), key=str.casefold)


def save_profile(settings: dict[str, Any], name: str, profile_dir: Path | str) -> Path:
    clean_name = str(name).strip()
    filename = profile_filename(clean_name)
    payload = deepcopy(settings)
    payload["config_name"] = clean_name
    directory = Path(profile_dir)
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / filename
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def load_profile(name: str, profile_dir: Path | str) -> dict[str, Any]:
    target = Path(profile_dir) / profile_filename(name)
    if not target.is_file():
        raise ValueError(f"找不到已保存配置“{name}”。")
    result = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(result, dict):
        raise ValueError(f"已保存配置“{name}”不是有效 JSON 对象。")
    result["config_name"] = str(result.get("config_name") or name)
    return result


def delete_profile(name: str, profile_dir: Path | str) -> Path:
    """Delete exactly one named profile file and return the removed path."""
    target = Path(profile_dir) / profile_filename(name)
    if not target.is_file():
        raise ValueError(f"找不到已保存配置“{name}”。")
    target.unlink()
    return target


def _template_stem(value: str) -> str:
    # Saved paths are commonly Windows paths even during Linux static tests.
    return Path(str(value).replace("\\", "/")).stem


def matching_profiles(template_path: str, prefix_length: int, profile_dir: Path | str) -> list[str]:
    try:
        length = int(prefix_length)
    except (TypeError, ValueError) as exc:
        raise ValueError("模板匹配前缀长度 N 必须是整数。") from exc
    if not 1 <= length <= 120:
        raise ValueError("模板匹配前缀长度 N 必须介于 1 和 120。")
    target_prefix = _template_stem(template_path)[:length].casefold()
    if not target_prefix:
        return []
    matches = []
    for name in list_profiles(profile_dir):
        try:
            profile = load_profile(name, profile_dir)
        except (OSError, ValueError):
            continue
        saved_path = str(profile.get("template_path") or "").strip()
        if saved_path and _template_stem(saved_path)[:length].casefold() == target_prefix:
            matches.append(name)
    return sorted(matches, key=str.casefold)
