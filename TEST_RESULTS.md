# 测试结果（v6）

日期：2026-10-05。验证环境：Linux，Python 3.12，openpyxl 3.1.5；测试只使用自动生成的虚构 XLSX 和假 WPS/COM 对象。

## 执行结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法检查、自检、JSON 检查和 **38 项 unittest 全部通过**。测试模拟配置文件与假 COM/WPS 调用；没有打开真实业务工作簿，也没有声称已验证 Windows GUI 或 WPS。

新版本归档为 `WpsSheetAutomation_v6_modular_workflow_20261005.zip`。ZIP 完整性与 CRC 检查通过，共 19 个成员；隔离解压后再次运行语法检查、JSON 检查、自检及单元测试，**38 项通过，0 项失败**。

## v6 新增覆盖

- 命名配置创建/保存/载入往返；按配置关联模板名测试 0、1、多份匹配，测试自定义 N 和非法 N。
- 工作流模块添加、移除、排序及实际执行；必需步骤错误顺序会具体说明依赖原因。
- 分列与去重/格式转换规则区域重叠时，拒绝会先修改输入或占用分列输出的顺序；不重叠模块允许按用户指定次序执行。
- PNG 与文本只读末端模块可互换；假处理验证二者都在工作簿最终保存之后且实际按配置顺序调用。
- 旧 v5 配置自动迁移到 v6 默认工作流和前缀长度 N=5；示例配置、GUI 配置收集和 Windows 构建脚本包含新结构。

原有多源数据追加、前/后映射、公式填充与重算、区域转换、范围校验、PNG 导出清理、文本预览/剪贴板及失败清理等回归测试继续通过。

## 未验证范围

**未在 Windows 11 + WPS/COM 上实机验证。** Linux 无法启动 WPS COM；真实 WPS 公式重算、`Range.CopyPicture`/图表 PNG 导出、Windows 剪贴板、实际 GUI 交互和 EXE 打包仍须在 Windows 11 + WPS 环境验收。没有访问或修改用户真实工作簿，也没有修改 Android/Harmony 工程。
