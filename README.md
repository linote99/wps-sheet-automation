# WPS 表格多源汇总处理器

面向 **Windows 11 + WPS 表格 + 无宏 XLSX** 的 GUI 工具。程序使用模板临时副本生成主结果，不覆盖模板或源文件。

## 本次更新（v14）

- 将 `TextToColumns`、`Deduplicate`、`FormatConvert` 在界面中的名称改为**文本分列、去重、格式转换**，并新增**区域筛选**。
- 移除独立区域 XLSX 导出及其关联的 PNG 输出；文本预览/复制仍保留。
- 文本分列、区域筛选、去重、格式转换的**每条规则均可选“源文件”或“目标工作簿”**。源文件规则只处理本次读取到的内存数据，不回写源 XLSX；目标工作簿规则处理输出模板。
- 旧配置中的 `image_outputs`、`export_region_pngs`、`open_region_exports` 和 `image_output` 工作流模块会被忽略并移除，配置升级为 v9。
- 修复 Windows 启动时授权窗口可能随隐藏主窗口一起不可见的问题；启动失败会弹出错误详情并写入 `%LOCALAPPDATA%\\WpsSheetAutomation\\logs\\startup.log`。

## 数据与模板设置

可以配置多个目标工作表；每个目标表可依次添加多份源文件，设置文件、工作表、复制起始/结束行和顺序。源工作表留空时使用该源文件的第一张表。复制列范围为全局设置，数据依次追加；结束行留空时按所有已使用列寻找首个整行空白边界，否则使用最后非空行。

公式从目标模板“粘贴起始行”的上一行读取并填充。公式结果列映射在首次完整重算后执行；前映射在主数据复制前执行，后映射在结果列映射后执行。前/后映射均为模板工作簿内的区域拷贝，源/目标尺寸必须一致。

## 区域处理

默认处理顺序是**文本分列 → 区域筛选 → 去重 → 格式转换**。可在“模块工作流”页调整顺序；规则列表内也可排序。区域操作在最终完整重算和保存之前执行。

每条规则要选择处理对象、工作表和指定区域。选择“源文件”时，规则仅应用于配置的源工作簿中同名工作表；若同名工作表未被任何源文件使用，程序会提示。选择“目标工作簿”时，规则应用于模板临时副本。源文件侧转换后的值或筛选/去重后保留的记录会进入主数据复制流程，源文件本身保持不变。

**区域筛选**根据指定条件列和条件保留匹配记录。条件包括等于、不等于、包含、不包含、开头为、结尾为、大于/大于等于/小于/小于等于、为空和不为空；文本匹配不区分大小写，数字条件支持数值比较。可设置是否有表头。筛选会在所选矩形区域内压缩保留的记录；为了让整条记录继续对齐，通常应把记录的所有相关列都包含在筛选范围内。

**去重**可按区域整行，或按区域内一列/多列作为关键字段；稳定保留第一次出现的记录，有表头时保留首行。它只压缩所选区域，不删除整张工作表的物理行，因此如需其他列随记录一起移动，应把这些列也包含在区域中。

**文本分列**从指定单列按分隔符拆分，逐列设置文本、数字或日期类型和输出起始列；目标区域必须为空，避免覆盖现有数据。源文件侧分列结果只有在全局复制列范围覆盖其输出列时才会被复制到目标。

**格式转换**把指定范围转换为文本、数字或日期，并设置相应数字格式。日期支持日期单元格、Excel 日期序号及常见日期文本格式。含公式的目标区域不会被转换；源侧规则也会拒绝含公式的区域，以避免依赖未保存的公式缓存。

区域规则每项最多处理 500,000 个单元格。目标侧区域筛选、去重和格式转换若触及公式区域或公式单元格，会停止处理以防破坏公式。

## 文本预览与配置

文本预览/复制仍可配置多个工作表范围。完成保存后，界面显示区域内容，并可按行换行、列间 Tab 的方式复制；不生成独立输出文件。

配置版本为 **9**。处理规则的 JSON 结构示例：

```json
{
  "config_version": 9,
  "workflow": [
    "template_cleanup", "pre_mapping", "source_copy", "formula_fill",
    "formula_recalc", "result_mapping", "post_mapping", "TextToColumns",
    "RegionFilter", "Deduplicate", "FormatConvert", "final_recalc_save", "text_output"
  ],
  "TextToColumns": {
    "enabled": true,
    "rules": [{
      "scope": "source", "sheet": "明细", "range": "D2:D100",
      "delimiter": "|", "output_start_col": "F", "column_formats": ["text", "number"]
    }]
  },
  "RegionFilter": {
    "enabled": true,
    "rules": [{
      "scope": "target", "sheet": "汇总", "range": "A1:F100",
      "has_header": true, "filter_column": "C", "operator": "greater_equal", "criteria": "10"
    }]
  },
  "Deduplicate": {
    "enabled": true,
    "rules": [{
      "scope": "target", "sheet": "汇总", "range": "A1:F100",
      "has_header": true, "key_mode": "columns", "key_columns": ["A", "C"]
    }]
  },
  "FormatConvert": {
    "enabled": true,
    "rules": [{"scope": "source", "sheet": "明细", "range": "C2:C100", "mode": "number"}]
  }
}
```

旧版独立区域文件和 PNG 设置会在载入时删除；未受影响的模板、源文件、映射、筛选/转换规则及文本预览配置会继续保留。默认配置见 `default_config.json`。

## Windows 构建

main-only 包附带 `install_and_build.bat` 和 Inno Setup 脚本。Windows 上需要 Python 3.10+、WPS 表格和可选的 Inno Setup 6；运行构建脚本后会执行语法检查、自检、静态回归测试，并生成 PyInstaller 程序目录；检测到 Inno Setup 时还会生成 v14 安装器。当前代码在 Linux 上完成静态及合成工作簿测试；Windows/WPS COM 与 EXE 仍需在目标 Windows 环境验证。
