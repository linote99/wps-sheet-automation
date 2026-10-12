# 测试结果（v5）

日期：2026-10-05。验证环境为 Linux，Python 3.12.3，openpyxl 3.1.5。

## 执行结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

**全部通过：** Python 语法检查、JSON 语法检查、自检以及 **35 项 unittest** 均通过。自检确认 v5 默认配置、公式范围/映射、两阶段映射、三类独立区域处理配置及旧版配置迁移有效。

独立版本包 `WpsSheetAutomation_v5_scope_selector_20261005.zip`：`unzip -t` 和 `zipfile.testzip()` CRC 校验通过，共 19 个成员；解压至隔离临时目录后再次运行 Python 语法检查、JSON 检查、自检及完整 unittest，**35 项通过，0 项失败**。

## 主要覆盖

- v5 默认 JSON 与旧配置迁移；源列表、前/后映射、图片/文本区域及 `TextToColumns` / `Deduplicate` / `FormatConvert` 各自独立的启用状态/规则列表配置保存和载入往返，包含自定义输出范围语法。
- 前映射在模板清理之前读值，可独立指定源/目标工作表；前映射目标与主数据写入区域冲突会拒绝处理；模板原件字节保持不变。
- 后映射在首次重算后执行；假 WPS 事件序列验证“首次重算 → 公式结果映射 → 后映射 → 文本分列 → 去重 → 值类型转换 → 再次完整重算 → 保存”。
- 共享选择器的 `A1` 全表、`A5` 整行、`D1` 整列、相同端点单格、不同及反向端点矩形语义；非法坐标、XFD/1048576 边界与越界拒绝。
- 区域处理单项 500,000 格限制、PNG 2,000 行/列及 100,000 格限制、文本 100,000 格/2,000,000 字符限制；整表文本按实际已用范围解析并以最多 500 行分块浏览/复制。
- 文本分列、稳定去重、文本/数字/日期转换纯逻辑；分列目标已有内容时拒绝覆盖。文本区域按行换行、列间 Tab 分隔；模拟复制按钮向剪贴板对象写入精确载荷。
- 多区域图片 PNG 命名包含工作表、范围与序号；已有文件不覆盖；模拟导出失败会清理临时资源。文本输出区域读取不创建文本文件。
- 端到端假数据流水线验证图片导出和文本预览读取发生在工作簿保存之后。
- 既有多源顺序追加、源结束行/全空行边界、公式种子填充、筛选/排序 XML、安全临时副本与失败清理测试仍通过。

测试仅使用自动生成的合成 XLSX、假 WPS/COM 对象与模拟 PNG 导出器；没有访问或修改用户上传的真实工作簿。

## 未验证范围

**未在 Windows 11 + WPS/COM 上实机验证。** Linux 无法启动 WPS COM；真实 WPS 的完整重算行为、`Range.CopyPicture`/临时图表 PNG 导出、Windows Tk 剪贴板、GUI 交互、Windows EXE 打包及真实模板/源文件端到端处理仍须在 Windows + WPS 环境验收。假对象只验证纯逻辑、命名、资源清理与调用顺序，不构成 WPS COM 兼容性证明。未修改 Android/Harmony 工作区。
