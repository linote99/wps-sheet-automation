# 测试结果（自定义区域独立 XLSX 文本输出）

日期：2026-10-06。验证环境：Linux，Python 3.12.3。测试数据全部为自动生成的虚构 XLSX；WPS COM、Windows GUI 和文件关联打开均未在 Windows + WPS 实机环境验证。

## 本次行为与回归覆盖

原 PNG 图片输出已替换为每个所选区域生成一个独立 `.xlsx` 文件。导出值写为文本；复制单元格字体、填充、边框、对齐、数字格式、行高、列宽和合并布局。合并范围闭包回归确认 `统计!AH1:BO44` 扩展至 `统计!AH1:CL44`。选区内隐藏整行和整列会从导出副本中剔除，并将其余可见内容压缩排列；原工作簿隐藏状态及内容保持不变。

回归还覆盖多区域多文件、不覆盖既有同名文件、100,000 格/2,000,000 字符安全上限、合并闭包、文本单元格类型、样式/行列尺寸迁移、隐藏轴全部剔除时拒绝空导出、部分导出失败时清理本轮区域文件，以及区域导出失败后保留并准确报告 WPS 已保存主 XLSX。

新增 `open_region_exports` 设置和 GUI 复选框，默认关闭。处理成功时，Windows 版通过系统默认文件关联打开本轮生成的区域文件；自动打开失败仅写入日志，不影响已保存结果。配置字段 `image_outputs` 与工作流 ID `image_output` 保留以兼容旧配置，含义已改为区域 XLSX 输出。

## 执行结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法检查、JSON 检查和静态自检通过；完整单元测试 **65 项通过，0 项失败**。已有配置、工作流、模板默认输出名、公式模板上一行规则、映射、区域处理、主 XLSX 保存保护及文本预览功能均保留。

## 验证边界

未读取用户业务工作簿或修改模板。测试验证的是 openpyxl 区域文件生成与假 WPS 流程；没有连接 WPS COM，也没有在 Windows 上实际运行 GUI、PyInstaller EXE 或 `os.startfile` 文件关联打开。目标 Windows + WPS 环境仍需实机验收。
