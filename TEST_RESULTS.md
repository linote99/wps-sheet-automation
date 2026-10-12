# 测试结果（WPS 表格多源汇总处理器）

## v23 源规则副本另存、读取及成功后可选清理（配置 v18）

当前 Linux 环境已通过语法编译、默认配置 JSON 解析、`main.py --self-check` 及完整静态/授权回归：**123 项通过、0 项失败**。新增端到端检查：源规则运行后先发布独立 `.xlsx` 副本，模板处理开始前副本已存在，模板读取函数实际使用该副本；默认成功后保留副本，启用删除时仅在模板任务成功后删除，模拟处理失败时副本仍保留，原始源文件和模板字节均不变。Fake WPS/COM 验证调用顺序和数据流，不替代 Windows/WPS 实机验收。

## v22 区域规则默认值与大小写匹配（配置 v17）

已在当前 Linux 环境执行语法编译、默认配置 JSON 解析、`python3 main.py --self-check` 及完整静态/授权回归：**122 项通过、0 项失败**。新增覆盖包括：空工作表名默认首表、工作表名称大小写不敏感、区域筛选空范围按 UsedRange 整表处理、`A`/`a` 整列、`A1`/`a1` 单格且大小写等价；空范围整页、源端空范围筛选先修改合并安全副本后再把筛选结果交给模板、公式填充/文本分列/筛选/去重/格式转换五类源规则均在合并源副本上映射执行，以及 v16→v17 配置迁移说明。

未连接 WPS COM，未读取真实业务文件，未在 Windows/WPS/EXE 中验收；Fake WPS/COM 回归仅验证调用顺序和数据流，不替代目标 Windows 环境验收。

## v21 模板合并源副本命名与可选保存（配置 v16）

已在当前 Linux 环境执行：语法编译、`main.py --self-check`、默认配置 JSON 解析及完整静态/授权回归；结果为 **118 项通过、0 项失败**。新增覆盖自定义名/空名默认使用模板子表名、可选保存与原件保护、未保存时合并临时工作簿贯穿模板处理并在任务结束后清理，以及旧 `retain_edited_sources` 到新保存字段的迁移。另在 Xvfb 下构建 Tk 界面并验证文件名字段和配置往返。

未连接 WPS COM、未在 Windows/WPS/EXE 中验收，也未读取真实业务文件。

## v20 独立源文件合并（配置 v15）

已在当前 Linux 环境执行：

```bash
python3 -m py_compile main.py wps_automation_workflow.py licensing.py keygen.py make_fake_data.py test_static.py test_licensing.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py test_licensing.py
```

语法检查、JSON 解析与 v15 自检通过；静态/授权回归 **116 项通过、0 项失败**。新增覆盖无模板 CSV+XLSX 独立合并、后续源首行保留/跳过、UTF-8 BOM 与 GB18030 CSV 读取、XLSX 首表默认行为、自定义输出名、源文件字节不变、拒绝非法输入顺序及不覆盖既有结果；并继续覆盖 v19 源合并与区域 XLSX-only 输出。另在 Linux Xvfb 下构建 Tk 窗口，检查独立合并列表、模式按钮、无模板配置收集与静态校验。

未连接 WPS COM，未读取/修改真实业务文件，也未在 Windows GUI、WPS 或 EXE 环境验收；仍需在目标 Windows 环境验证 GUI 表现及大文件处理体验。

v18 验证记录日期：2026-10-07。验证环境：Linux，Python 3.12.3，openpyxl 3.1.5，cryptography 50.0.1。工作簿回归使用自动生成的虚构 XLSX；未读取或修改真实模板、源文件或业务输出。

## v18 设置中心、秘钥管理与配置管理

源码目录 `/workspace/wps_sheet_automation/` 已通过 `py_compile`、默认配置 JSON 解析、`main.py --self-check` 及 `python3 -m unittest -v test_static.py test_licensing.py`；完整回归 **112 项通过、0 项失败**。新增测试覆盖命名配置仅删除选定档案，以及可读取已签名秘钥到期日但仍拒绝过期授权。使用 Xvfb 启动 Tk 界面并构建设置窗口，确认“秘钥管理 / 使用说明 / 配置管理”三个页面创建成功。

main-only v18 发布副本 `/workspace/releases/split/WpsSheetAutomation_main_v18/` 已通过编译、v13 静态自检、默认 JSON 检查及静态套件 **100 项通过、0 项失败**。已核对发布目录不含 `.pem` 或私钥文件；未运行 Windows 安装器、WPS COM 或 EXE，仍须在目标 Windows + WPS 环境验收。

## v17 修订版：可选结果映射与 WPS 区域公式填充双模式

在源码目录 `/workspace/wps_sheet_automation/` 验证：

```bash
python3 -m py_compile main.py wps_automation_workflow.py licensing.py keygen.py make_fake_data.py test_static.py test_licensing.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -q test_static.py test_licensing.py
```

编译、JSON 解析及 v13 静态自检通过；完整主程序与授权回归 **110 项通过、0 项失败**，其中静态主程序套件 **99 项通过**。新增覆盖结果映射与后映射独立并存、结果映射可选、工作流顺序、旧配置迁移、目标冲突保护及 WPS 实际读写顺序；区域公式填充也不能覆盖结果映射目标。并覆盖自定义公式双模式、公式模板区域保护以及源文件先重算后读取等。

main-only 发布目录 `/workspace/releases/split/WpsSheetAutomation_main_v17/` 已独立通过编译、JSON 解析、v13 自检及静态套件 **103 项通过、0 项失败**；结果映射和后映射均可独立留空，配置时先后顺序正确且冲突保护有效。Inno Setup 产品版本和安装器输出名仍为 v17。

Fake WPS/COM 验证调用和数据流，不等同真实 WPS 函数兼容性测试；当前 Linux 环境未运行 Windows GUI、WPS COM 或 EXE，仍需在目标 Windows + WPS 环境验收公式输入语法、函数计算和相对引用结果。

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

历史 v17 测试曾按当时规则将 `A1` 视为整表；当前 v22 已统一改为 `A1`/`a1` 单格、空范围整页，`A`/`a` 整列；整列静态配置不按 1,048,576 行触发上限。合成 XLSX/COM 测试确认目标侧按 `UsedRange` 行数筛选，源侧处理到整张源表最后数据行（包括条件列为空但其他列有值的行），且原源文件字节不变。专项测试与当时完整回归 **90 项通过**。

## Windows / WPS 验收边界

当前执行环境是 Linux，没有可切换的 Windows 设备，因此没有生成或运行 Windows `.exe`，也没有实际启动 Windows DPAPI、Tkinter 授权窗口或 WPS COM。`install_and_build.bat` 可在目标 Windows 11 + WPS 电脑执行构建与静态检查；仍需在目标机验收 WPS COM 源文件合并与规则处理、副本保存、模板处理、区域 XLSX 导出及最终 EXE 行为。私钥会被复制到管理员专用的秘钥生成器目录，切勿将该目录分发给软件用户。

离线授权只能提高普通复制和误用的门槛，无法保证抵抗专业逆向或彻底阻止用户回拨本地系统日期。PyInstaller 打包不是 Python 源码强加密；详细限制与私钥保管流程见 `LICENSING_README.md`。
