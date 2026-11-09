# MRO 恢复工程

本节点交接见 **[HANDOFF.md](HANDOFF.md)**。其中列明恢复依据、差异、Windows 测试和后续 Git 上传范围。

原始发布文件完整保存在 **[original-release/MRO-Windows-x64](original-release/MRO-Windows-x64/)**，包含 EXE、DLL、PDB、运行配置、Excel 和安装脚本，供 Windows 对照及后续恢复。文件哈希和验证说明见 [original-release/README.md](original-release/README.md)。

两个窗口的 XAML 从保留的 BAML 解析得到，原 Logo、窗口图标、尺寸、布局与样式已复用。窗口事件和 Core 数据流程为重新实现，**不是恢复了全部原业务源码，也未验证能否编译和运行**。

本版提供标准表导入、称重 XLSX 导入、条码关联、抽样数量校验、真实测量报告、辅助称重表导出及共享操作日志。报告统计已导入的测量值。

## 回家在 Windows 上验证

安装 **.NET 8 SDK（Windows x64）**，重新打开 PowerShell。仅安装 Desktop Runtime 无法编译。也可用 Visual Studio 2022 17.8+ 并安装 .NET 桌面开发与 .NET 8 SDK。

在本目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Test-Windows.ps1 -Package
```

脚本会构建 `Mro.sln`、运行 Core 控制台检查，成功后发布到：

- `artifacts/windows/MainApp/Mro.MainApp.exe`
- `artifacts/windows/HelperApp/Mro.HelperApp.exe`

第一次还原或发布可能需要联网下载 Windows 目标包。日志保存在 `artifacts/windows-check-*.log`。脚本会另生成带安装脚本的 ZIP，但不会自动安装；解压后可单独运行 Install-MRO.cmd。

手动启动和界面验收见 [Windows 检查清单](docs/WINDOWS-TEST.md)。任何构建失败，请保留完整日志继续修复；当前没有把未运行的测试标记为通过。

## 源码结构

```text
apps/Mro.MainApp/       主窗口 XAML、资源与新事件代码
apps/Mro.HelperApp/     辅助窗口 XAML、资源与新事件代码
apps/Shared/            绑定模型、状态、日志和通用窗口支持
src/Mro.Core/           XLSX 读写、标准目录、测量数据、校验和报告
tests/Mro.Core.Specs/   不依赖第三方测试包的运行检查
SharedContent/         原标准表、抽样配置示例和旧报告格式参考
scripts/               Windows 构建检查与跨平台静态检查
docs/                  恢复证据、限制和待测项
```

### 当前与旧版的差异

- UI 的结构和资源来自发布包；少量按钮及说明文字改为真实测量流程，详见 `docs/ui-recovery.json`。显示效果仍需在相同 Windows 字体、DPI 和窗口尺寸下比较。
- 辅助窗口“最大值(g)”列改为只读“实测样品数”，导入称重表后再核对生产数量。删除按钮仅删除空条目。
- 已从原报告制作空白样式模板，保留字体、边框、合并、尺寸和页面设置。汇总每页15产品、明细每页100记录，另有测量追溯和说明。打印效果仍待 Windows 验证。
- 旧版 T1/T2 分类及计数函数已按 IL 恢复，按用户选择标注待法规复核。平均值和数量偏差单独显示；感官结果不假定正常。
- 增加“导入生产表”，兼容首表 A 条码/B 数量，另支持命名表头；数量可手工核实。恢复路径偏好、打包、独立安装和卸载脚本，这些操作尚未在 Windows 运行。
- 仅支持 `.xlsx` 导入。`ReferenceReports` 中的 `.xls` 仅保留为格式参考。
- 数据保存到 `%LOCALAPPDATA%\MRO-Recovered`，与旧应用目录分开；主、辅程序共享 `standard-catalog.v2.json` 和 `operations.jsonl`。

### 抽样规则配置

已根据原程序集确认按上限分段：1–20全检，21–40→32，41–60→35，61–80→47，81–100→49，101–200→64，201–300→67，301–400→81，401–500→81，501–100000→98。原标准表的规则行按同样方式读取，并包含允许 T1/T2 数量。

21–31等小批量可能得到大于批量的应抽数，这是旧版规则的实际问题；恢复版阻止这种数据导出。此口径尚未独立核实法规。自定义时可将 `SharedContent/sampling-plan.example.json` 复制为 `sampling-plan.json`，编辑后重新导入标准库；不要把旧版业务口径当作法规认证。

### 生产表和路径偏好

生产表支持首工作表无表头 A列条码/B列正整数数量；或命名列 `条码`/`编码`/`产品编号`、`生产数量`/`抽样基数`。可附名称及生产日期；条码与编码同时出现时必须一致。未知、重复和缺量条目报错。

主程序保存标准表、称重表和输出目录；辅助程序保存输入和输出目录。两程序更新同一偏好文件时合并字段，失效路径回退，损坏 JSON 保存前备份。

### 称重数据格式

需要唯一表头，包含 `扫描`、`重量`、`单位`；可附 `序号`、`日期`、`时间`、`皮重`、`毛重`。扫描列按条码精确关联；不会自动按名称或另一列“编码”猜测。

- 支持 `g`、`kg`、克、千克、公斤及单位前后空格；“重量”作为净重输入。
- 公式测量值、非正重量、未知单位、缺失条码和重复记录 ID 会报错。
- 标准表有两组重复条码，关联这些条码时会报歧义；数值型条码会提示核对，文本前导零会保留。
- 已导入的所有记录参与平均值计算，不删除低于标准规格的记录；缺测时不自动补值。
- 来源文件、工作表和行号用于回查，不等于经过仪器签名认证。

### Mac 上能检查什么

```bash
python3 scripts/verify_recovery.py
```

该脚本只检查 XAML/XML、资源哈希、绑定、事件与工程引用，不执行 C#。有 .NET 8 SDK 时可单独构建并运行 Core 检查：

```bash
dotnet run --project tests/Mro.Core.Specs -c Release
```

当前 Mac 完成的检查记录见 [恢复状态](docs/RECOVERY-STATUS.md)。
