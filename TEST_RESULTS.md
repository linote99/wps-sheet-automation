# 测试结果（机器码授权与 UFI-TOOLS 秘钥生成器）

日期：2026-10-07。验证环境：Linux，Python 3.12.3，openpyxl 3.1.5，cryptography 50.0.1。工作簿回归使用自动生成的虚构 XLSX；未读取或修改真实模板、源文件或业务输出。

## v16：模板独立编辑、WPS 源副本处理与区域输出

在源码目录 `/workspace/wps_sheet_automation/` 验证：

```bash
python3 -m py_compile main.py wps_automation_workflow.py licensing.py keygen.py make_fake_data.py test_static.py test_licensing.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py test_licensing.py
```

编译、JSON 解析、v11 静态自检通过；完整主程序与授权回归 **96 项通过、0 项失败**。新增覆盖包括：无源模板的完整处理仍保留模板内容、无源模式不需要填写复制起始/末列、源侧规则在 WPS 临时副本中先于模板处理（端到端时序验证），勾选时另存编辑后的源副本且原文件字节不变、整表/整行/整列按 WPS `UsedRange` 裁剪、区域 XLSX 的填充色/字体颜色/数字格式与显示文本、区域 XLSX 始终输出而 PNG 可选，以及旧配置工作流迁移。

由源码目录执行的主程序与授权全套回归为 96 项通过；从 main-only v16 发布目录单独执行的构建前静态套件为 **85 项通过、0 项失败**，安装器元数据和输出名均为 v16。

未在 Windows/WPS COM、GUI 或 EXE 环境实测；Fake WPS/COM 仅用于验证调用与数据流，不能替代目标电脑的实机验收。

## v15 区域筛选范围简写验证

已验证区域筛选 `A` 解析为整列、`a1` 解析为单格、矩形写法和旧 `A1` 整表语义保持；整列静态配置不按 1,048,576 行触发上限。合成 XLSX/COM 测试确认目标侧按 `UsedRange` 行数筛选，源侧处理到整张源表最后数据行（包括条件列为空但其他列有值的行），且原源文件字节不变。专项测试与当时完整回归 **90 项通过**。

## Windows / WPS 验收边界

当前执行环境是 Linux，没有可切换的 Windows 设备，因此没有生成或运行 Windows `.exe`，也没有实际启动 Windows DPAPI、Tkinter 授权窗口或 WPS COM。`install_and_build.bat` 可在目标 Windows 11 + WPS 电脑执行构建与静态检查；仍需在目标机验收 WPS COM 源文件处理、副本保存、模板处理、区域 PNG 导出及最终 EXE 行为。私钥会被复制到管理员专用的秘钥生成器目录，切勿将该目录分发给软件用户。

离线授权只能提高普通复制和误用的门槛，无法保证抵抗专业逆向或彻底阻止用户回拨本地系统日期。PyInstaller 打包不是 Python 源码强加密；详细限制与私钥保管流程见 `LICENSING_README.md`。
