# 测试结果（区域 XLSX 安全读取修复）

日期：2026-10-06。验证环境：Linux，Python 3.12.3，openpyxl 3.1.5。全部工作簿均为自动生成的虚构 XLSX；未读取、修改用户模板或输出。

## 回归覆盖

区域 XLSX 导出以及保存后文本区域读取，现统一先对 WPS staging 建立受限的安全临时副本，再由 openpyxl 加载副本。区域输出仍为每个所选区域单独生成 XLSX，值写为文本，保留样式、合并与行列尺寸，剔除并压缩区域中的隐藏行列；PNG 失败时保留主 XLSX 的既有流程未改动。区域文件使用对应模板工作表名，单份无序号，多份按配置顺序追加 `1`、`2`、`3`；主工作簿默认文件名仍为模板 basename。

新增合成 XLSX 回归在已保存 staging 的 worksheet XML 中写入非法自动筛选范围，验证区域 XLSX 和文本区域读取均通过 sanitized 临时副本成功，并确认源 XLSX 原字节及 `mtime_ns` 均不变。另用错误 `workbook.xml` 验证不相关结构错误不会被静默删除、会明确报错，且流程仍保留并报告已保存主 XLSX。安全副本异常清理模拟 Windows sharing violation，确认 close-before-unlink、有界重试、失败时给出警告但不把成功的导出改判失败。

既有回归继续覆盖多区域独立文件、模板工作表命名及后缀、同名文件不覆盖、数据文本类型、样式/合并/尺寸、隐藏项剔除、自动打开选项、主 XLSX 保留、模板 basename 默认名、映射与工作流、公式模板行、配置迁移及文本输出。

## 验证结果

在 `/workspace/wps_sheet_automation/` 执行：

```bash
python3 -m py_compile main.py workflow.py make_fake_data.py test_static.py
python3 -m json.tool default_config.json
python3 main.py --self-check
python3 -m unittest -v test_static.py
```

语法和配置检查通过，静态自检通过，完整测试 **69 项通过、0 项失败**。发行包 `WpsSheetAutomation_v10_region_xlsx_safe_read_worksheet_names_20261006.zip` 的 ZIP CRC 检查通过；在独立临时目录完整解压后再次运行语法检查、配置检查、自检和测试，**69 项通过、0 项失败**。

## 验证边界

测试未连接 WPS COM，也未在 Windows + WPS/COM 环境实机验证；Windows 文件锁由模拟异常覆盖，自动打开和 GUI 也未在 Windows 实机测试。