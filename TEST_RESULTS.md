# 测试结果（PNG 导出失败后保留已保存 XLSX）

日期：2026-10-06。验证环境：Linux，Python 3.12.3。测试使用自动生成的虚构 XLSX、假 WPS/COM 对象和模拟图片导出器；没有使用真实业务工作簿。

## 本次行为与回归覆盖

WPS 完成计算并成功保存暂存工作簿后，如果随后的 PNG 导出失败，程序现在会以不覆盖现有目标文件的方式发布已保存 XLSX，并在图片错误提示中明确说明 XLSX 已保留及其完整路径。WPS 保存成功前发生错误仍会清理暂存文件、不发布最终 XLSX；文本区域输出失败仍按原策略清理暂存工作簿和本轮图片，不发布最终文件。

新增回归测试模拟 WPS 保存已编辑工作簿后 PNG 导出抛错，验证最终 XLSX 文件存在、可由 openpyxl 重新打开且包含 WPS 保存后的单元格值，并检查错误信息包含准确目标路径。另覆盖保存前失败不发布、文本失败策略不变，以及并发/已有目标 XLSX 不会被覆盖。保留现有公式模板行自动使用 start_row 上一行、pywin32 安装与启动入口、合并单元格图片范围闭包扩展及不修改隐藏行列状态等功能。

## 执行结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法检查、JSON 检查和静态自检通过；完整单元测试 **59 项通过，0 项失败**。安装脚本/启动器静态断言、公式模板行、图片范围扩展与隐藏属性保护等既有覆盖也全部通过。

独立新版归档为 [WpsSheetAutomation_v7_png_failure_xlsx_preserved_20261006.zip](/workspace/wps_sheet_automation/WpsSheetAutomation_v7_png_failure_xlsx_preserved_20261006.zip)，共 20 个成员。ZIP CRC 与 `unzip -t` 校验通过；在独立临时目录解压后重新运行语法检查、JSON 检查、静态自检及完整单元测试，**59 项通过，0 项失败**。

## 验证边界

测试只使用虚构数据和模拟 WPS/COM。**未在 Windows + WPS 实机验证**，因此实际 WPS 计算/保存、PNG 截图、Windows GUI 与 EXE 仍待目标环境验收。
