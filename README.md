# WPS 表格多源汇总处理器 v24 主程序源码包

适用范围：Windows 11 + WPS 表格 + 无宏 XLSX。v24 / 配置 v18。详细功能和操作说明见 `USER_GUIDE.md`。

## v24 更新

- 主窗口和 Windows 程序图标使用 Lino 笔记标志。
- 设置中新增“关于”，介绍运行方式和适用范围，并提供 Lino 笔记、GitHub 主页和项目源码链接。
- 首页启动显示版本提醒；选择“不再提示”后偏好保存在本机，不混入业务配置。
- 保留 v23 的源规则处理副本保存/读取流程及模板成功后可选清理行为。

## 在 Windows 上构建主程序

1. 在安装了 WPS 表格的 Windows 11 电脑上解压本目录。
2. 安装 Python 3.10 或更高版本，并确保 Python Launcher（`py`）可用。
3. 双击 `install_and_build.bat`。该脚本创建 `.venv`、安装依赖、执行语法检查、自检和静态回归，然后构建主程序目录。
4. 构建结果位于 `dist\WpsSheetAutomation\WpsSheetAutomation.exe`；源码调试入口为 `run_app.bat`。

首次构建需要联网安装依赖。脚本只构建主程序，不构建授权密钥生成器，不复制签发私钥。实际 WPS COM、公式重算、Windows 界面及 EXE 仍须在目标 Windows/WPS 环境验收；当前已通过的测试来自 Linux 静态/合成测试和 Xvfb 界面烟测。

## 包含内容与安全说明

包含 v24 主程序源码、工作流、默认配置、测试、说明和 Lino 图标资源。

本包不含 EXE、Inno Setup 安装器、密钥生成器、签发私钥、`.venv`、构建输出或缓存。授权签发私钥必须留在管理员本机，绝不可放入主程序或普通发行包。
