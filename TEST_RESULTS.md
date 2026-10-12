# 主程序包验证结果

日期：2026-10-07。验证环境：Linux x86_64、Python 3.12、openpyxl。所有工作簿测试均使用自动生成的合成数据；未读取或修改真实业务模板、源文件或业务输出。

## 已通过的检查

在 v13 main-only 发布目录执行：

```bash
python3 -m py_compile main.py wps_automation_workflow.py licensing.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py
```

语法编译、默认配置 JSON、自检及静态/合成工作簿回归检查通过：**73 项通过，0 项失败**。新增用例验证本地工作流模块改用 `wps_automation_workflow.py`，避免 PyInstaller hooks-contrib 的 `hook-workflow.py` 因缺少第三方 `workflow` distribution metadata 而失败；并覆盖区域筛选比较和区域压缩、四类处理规则可配置为源文件、源文件规则只影响本次读取结果而不改写原件、目标工作簿筛选、旧区域 XLSX/PNG 配置迁移清除，以及不再生成区域导出旁路文件。

此外，Linux 临时环境使用 PyInstaller 6.22.3 与 pyinstaller-hooks-contrib 2026.8 成功完成 main-only onedir 冒烟打包；分析图包含 `wps_automation_workflow`，且未调用 `hook-workflow.py`。此项仅验证同名 hook 冲突已消除，不代表生成或验收了 Windows EXE。

## Windows 安装包状态

当前执行环境为 Linux，没有 Windows、WPS COM 或 Inno Setup 编译器；因此没有生成或实测 Windows EXE，也未验证实际 WPS COM 执行。main-only 发布包的构建脚本会在 Windows 上运行语法检查、自检、静态回归及 PyInstaller；检测到 Inno Setup 6 时生成 v13 安装器。

Windows 11 + WPS 环境仍需实际启动程序验证 COM 读写、重算、筛选/转换及最终保存流程。发布包不包含密钥生成器或签发私钥。
