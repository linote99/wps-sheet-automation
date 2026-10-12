# 测试结果（模板默认输出名与WPS图片导出重试）

日期：2026-10-06。验证环境：Linux，Python 3.12.3。测试数据全部为自动生成的虚构 XLSX，图片导出由假 WPS/COM 对象模拟；未在 Windows + WPS 实机测试。

## 本次行为与回归覆盖

选择模板后，输出文件名默认采用模板完整 basename（含 `.xlsx`）；模板切换会更新自动默认名，用户手动编辑的自定义名不被覆盖。命名配置显式保存自动/自定义状态：载入自动命名配置时使用本次所选模板名，载入自定义配置时保留保存的名称；旧配置按旧通用默认名和模板文件名兼容推断。同目录、同名输出会明确拒绝，原模板不被覆盖，并提示改名或选择其他输出目录。

图片导出保留 WPS `CopyPicture → 临时图表 Paste → Chart.Export` 路径，加入可用对象激活、范围选择、COM 消息泵与短等待，并尝试屏幕位图、`xlPicture`、再试位图三轮；不新增大型系统依赖。最终失败时错误包含配置范围、实际闭包范围及每轮失败阶段；现有 PNG 失败后保留已保存 XLSX 的处理继续生效。

模拟 `统计!AH1:BO44` 与 `AH1:CL1` 合并标题相交，完整合并闭包得到 `统计!AH1:CL44`。测试确认配置范围仍为 `AH1:BO44`、隐藏行列快照不变；还覆盖格式回退、瞬时 `Chart.Export` 失败重试、持续失败诊断/清理，以及 PNG 失败后已保存 XLSX 存在、可读且错误含准确路径。

## 执行结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法检查、JSON 检查和静态自检通过；完整单元测试 **64 项通过，0 项失败**。此前功能也全部保留并通过，包括公式模板行自动取 `start_row - 1`、pywin32 安装和启动脚本、合并单元格闭包扩展与隐藏行列状态不变、PNG 失败后保留 XLSX。

独立归档为 `WpsSheetAutomation_v8_template_filename_png_retry_20261006.zip`。归档 CRC、`unzip -t`、隔离目录解压后的语法/JSON/静态检查和完整单元测试均已验证通过。

## 验证边界

真实模板只读检查了工作表、合并区域和隐藏行列元数据；未读取单元格业务值，未修改模板或用户机器。假 WPS 可验证重试顺序、范围闭包、隐藏状态保护和错误路径，但不能证明真实 WPS 会成功生成 PNG。**未在 Windows 11 + WPS/COM 实机验证**；实际 COM、GUI 和 EXE 仍需目标环境验收。
