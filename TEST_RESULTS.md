# 主程序包验证结果

日期：2026-10-07。验证环境：Linux x86_64、Python 3.12、openpyxl。所有工作簿测试使用自动生成的合成数据；未读取或修改真实业务文件。

在 v14 main-only 发布目录执行：

```bash
python3 -m py_compile main.py wps_automation_workflow.py licensing.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py
```

语法编译、默认配置 JSON、自检及静态/合成工作簿回归检查通过：**75 项通过，0 项失败**。新增测试覆盖隐藏根窗口下授权弹窗的显示顺序，以及窗口化启动异常的日志落盘。

完整源码目录另有授权模块测试，**11 项通过**；加上 main-only 回归，总计 86 项通过。当前执行环境不是 Windows，未生成或运行 Windows EXE、Inno Setup 安装器，也未验证 WPS COM。需在 Windows 11 上用此包重新构建并安装后确认实际窗口行为。

若新版本仍无法启动，可查看 `%LOCALAPPDATA%\\WpsSheetAutomation\\logs\\startup.log`，其中包含启动阶段和异常堆栈。
