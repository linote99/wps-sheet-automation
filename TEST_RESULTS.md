# 测试结果（区域 XLSX 与 PNG 导出选项）

日期：2026-10-06。验证环境：Linux，Python 3.12.3，openpyxl 3.1.5。测试工作簿均为自动生成的虚构 XLSX；未读取或修改用户模板、源文件或真实输出。

## 回归覆盖

区域 XLSX 输出保留所选范围内全部行列及其原始相对位置；隐藏行列不会删除或压缩，导出后仍保持隐藏。单元格字体（含颜色）、填充、边框、对齐、数字格式、行高、列宽和合并区域继续复制；输出值仍写为文本，而非公式/数值。

Windows 环境会尝试在净化后的临时副本上调用 WPS COM `Range.Text` 获取数值/日期显示文本；窄列返回 `####` 时忽略该结果并回退至 `number_format` 格式化。Linux 回归覆盖千分位、百分比、小数精度、前导零、自定义零值显示、日期、布尔值，并用模拟 COM 验证 `Text` 优先、`####` 被排除、只读打开及关闭顺序。

PNG 选项新增回归覆盖：配置默认关闭；旧 `open_region_exports` 字段迁移到 `export_region_pngs`；图像由已生成的区域 XLSX 经 WPS 图表剪贴/导出接口生成；同名 PNG 不覆盖；转换完成后才删除对应 XLSX；模拟转换失败时保留区域 XLSX 与已保存主工作簿。假 WPS 对象用于验证调用和文件生命周期，不等同真实 WPS 渲染。

## 验证结果

在源码目录 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法检查、JSON 解析、静态自检均通过；完整测试 **73 项通过、0 项失败**。

## 验证边界

当前环境没有 Windows/WPS COM，未在 Windows 11 + WPS 上实测图片渲染、宏 GUI 或文件锁行为。区域 PNG 生成依赖 WPS 图表剪贴/导出 COM 接口，需在目标 WPS 环境首次验收。遇到 COM 不可用或 `####` 溢出时，纯 Python `number_format` 回退尽量保留常用显示格式；地区设置、条件格式和复杂自定义格式可能导致显示差异。所有读取和转换都基于处理副本，不修改原模板或源文件。
