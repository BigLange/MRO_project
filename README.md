# MRO 恢复工程

当前范围以[旧版基线](docs/LEGACY-BASELINE.md)和[项目交接](HANDOFF.md)为准。用户要求先恢复旧版，后续新增或调整功能由用户明确提出。

2026-11-10已按用户要求调整[浮动比例输入和80%两端分布](docs/CUSTOMER-CHANGES.md)：比例可填数字110、百分比或旧文字，也可读取该列公式保存的结果；公式重新计算保存后需重新导入表1。辅助生成约80%落在两端各25%内，约20%在中间。

## 使用

1. 打开主程序，重新导入表1 `检测标准表.xlsx`。当前标准库为 `standard-catalog.v4.json`，按旧规则读取104个唯一条码产品。
2. 辅助程序导入表2（首表A列条码、B列数量），或手工填写条码、抽样基数、最大值。点击“预览样品数”，再点击“生成标准表”，生成一份原格式8列称重Excel。
3. 主程序分析该称重表，填写抽样基数或点击“一键生成标准抽样基数”，选择报告日期，导出各商品独立报告和单独汇总。
4. 报告直接使用原始Excel模板，不附加来源标记、复核文字、追溯页、说明页或新栏目。

旧版的样品补足/裁减、默认感官结果、日期、判定文字和编码排序均按已核对的旧DLL恢复。具体边界及依据见旧版基线文档。

## 构建

Windows .NET 8 SDK：

```powershell
.\scripts\Test-Windows.ps1 -Package
# PowerShell 7.4+，完成Release构建后可比对旧DLL：
.\scripts\Compare-LegacyGenerator.ps1
.\scripts\Compare-LegacyWorkflow.ps1
```

脚本支持 `-DotnetPath`，优先使用本机项目专用SDK；发布为Windows x64自包含程序。

- 主程序：`artifacts/windows/MainApp/Mro.MainApp.exe`
- 辅助程序：`artifacts/windows/HelperApp/Mro.HelperApp.exe`
- 构建日志、ZIP发布包：`artifacts/`

17组Core检查、辅助生成24组差分、主程序补足/裁减24组差分已通过。界面与打印验收见[Windows检查](docs/WINDOWS-TEST.md)。完整原源码、所有DPI、客户Excel/WPS打印和安装/卸载仍未完全验收。

原始发布目录 `original-release/`、原标准表、参考报告及图像字节保持不变。原表单仅填入本次数据，避免带入历史结果；技术上的暂存写入和失败回滚不改变报告内容。未自行提交或推送Git。
