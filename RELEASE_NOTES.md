# WPS v27 主程序源码构建包

程序版本 v27，配置 schema v20。此版本合并 v26 的选定源文件处理功能与 main 上的连续整行空白停止阈值功能，且保留两边的历史提交，不重写或移动旧标签。

配置 v20 会兼容两种历史 v19 配置：保留已有的空行阈值和源文件选择设置；缺少阈值时默认为 1 行，缺少独立源处理选项时沿用旧模板流程。

本包包含主程序源码、图标、默认配置、测试与 Windows 构建文件。发布包不含授权签发器、签名私钥、许可证、缓存、虚拟环境或构建产物；普通 main-only 构建入口不会构建或分发签发器。

## 构建

在 Windows 安装 Python 3、WPS 表格和 Inno Setup 6 或 7，运行 `install_and_build.bat`（或 `install_and_build_main_only.bat`）。依赖见 `requirements.txt`。安装器输出名为 `WpsSheetAutomation_Setup_v27.exe`。

## 验证边界

本次 Linux 环境运行的语法检查、自检及回归测试结果记录在 `TEST_RESULTS.md`。Linux 环境不包含 Windows/WPS COM 和 Inno Setup，未生成或验收 Windows 安装器；模拟 COM 测试不替代目标 Windows + WPS 实机验收。
