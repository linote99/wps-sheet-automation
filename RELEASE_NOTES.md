# WPS v26 主程序源码构建包

程序版本 v26，配置版本 v19。主程序来自仓库提交 `6457abea0b542f4539531ac291deb27b07b5a5ff`，与当前本地主程序一致。

此包只提供主程序源码、图标、配置、测试及 Windows 构建文件。两个构建入口均使用 main-only 构建脚本。已排除签发器、keygen、管理员插件、私钥、许可证文件、历史备份、缓存、虚拟环境和构建产物；许可模块中未被主程序使用的签发函数及配套编码辅助函数也已移除，保留公钥验签与用户激活功能。

README、CHANGELOG 与 TEST_RESULTS 是项目原有历史文档，其中旧版本签发器相关描述不代表本包包含这些工具。

## 构建

在 Windows 安装 Python 3、WPS 表格和 Inno Setup 6 或 7，运行 `install_and_build.bat`（或 `install_and_build_main_only.bat`）。依赖在 requirements.txt 中列出。

## 验证边界

发布前检查包内容、8 个 Python 文件语法、ZIP 完整性与 SHA-256；本次 Linux 环境运行 main-only 静态/模拟回归测试 143 项，全部通过。没有进行 Windows/WPS 真机验收，也没有在本环境构建 EXE；不将历史回归测试结果视为本次真机测试。
