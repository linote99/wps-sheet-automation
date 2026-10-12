# 主程序包验证结果

日期：2026-10-07。验证环境：Linux x86_64，Python 3.12.3，openpyxl 3.1.5，cryptography 50.0.1。测试仅使用源码、自检和自动生成的合成数据；未读取或修改真实业务模板、源文件或业务输出。

## 已通过的检查

在主程序源码目录执行：

```bash
python3 -m py_compile main.py workflow.py licensing.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py
```

Python 语法编译、默认配置 JSON、自检均通过；静态与合成工作簿回归 **74 项通过，0 项失败**。新增安装器回归检查验证 Inno Setup 安装脚本按当前用户安装、从 PyInstaller 主程序目录收集文件，并且安装定义与构建脚本不引用密钥生成器或签发私钥。启动错误提示现已统一指向实际构建日志 `build_main.log`。

## Windows 安装包状态

当前可用执行环境为 Linux；没有 Windows、PyWin32/Windows COM、Inno Setup 编译器或 WPS，因此没有生成、启动或实测 Windows `.exe`。源码包现已包含 Inno Setup 安装定义 `WpsSheetAutomation.iss`，并更新 `install_and_build.bat`：在 Windows 上通过语法、自检、回归及 PyInstaller 构建后，如果找到 Inno Setup 6 的 `ISCC.exe`，会生成 `dist\\WpsSheetAutomation_Setup_v12.exe`；若找不到编译器，会明确说明只生成了 PyInstaller 程序目录、没有生成安装器。

要产出单文件安装器，需在 Windows 11 构建机安装 Python 3.10+ 和 Inno Setup 6，并在联网状态下运行 `install_and_build.bat`。最终仍需在目标 Windows 11 + WPS 环境启动安装后的主程序，验收真实 WPS COM 流程。此主程序包与密钥生成器、签发私钥分开，安装器不会包含它们。
