# MRO 恢复交接 · 2026-11-09

**本节点已完成源码与资源整理；尚未通过 .NET 编译、Core 运行测试或 Windows 界面验收。** 当前 Mac 没有 .NET SDK，无法联网下载工具。已准备 Windows 检查脚本，不能把静态检查当成运行通过。

用户已确认：判定部分先恢复原程序口径，标注“待法规复核”。Git 交接状态见文末。

## 本节点产物

| 部分 | 当前实现与依据 | 待验证 |
| --- | --- | --- |
| 界面 | 两个窗口由原 BAML 解析；复用 Logo、图标、尺寸、配色、样式和绑定。各增加一个“导入生产表”按钮 | WPF 编译、与旧版的 DPI/字体/布局对照 |
| 抽样规则 | 核对原 `SamplingPlanResolver` 和标准表读取器 IL，补齐上限分段和允许 T1/T2 数量 | Windows Core 检查、确认业务及法规适用性 |
| 汇总和明细 | 全部真实记录参与平均值、最小值、最大值、T1/T2 统计；保留来源文件、工作表、行和记录 ID | Excel/WPS 数值及往返导入 |
| 报告排版 | 从原报告复制字体、填充、边框、合并、行列尺寸和页面设置，制作清空历史结果的模板；汇总15产品/页、明细100记录/页 | 打印排版，不能宣称像素级一致 |
| 原版判定 | 恢复单件分类和批次计数口径；平均值、数量偏差和法规复核状态另列 | 不是经过法规认证的整体合格判定 |
| 生产表 | 兼容原导入器的首表 A 列条码、B 列数量；另支持命名表头和显式编码关联 | 客户实际生产表样例 |
| 路径偏好 | 恢复7个原字段，增加主程序生产表目录；JSON保存、跨进程合并更新、损坏文件备份 | 双程序切换、关闭重开、文件失效回退 |
| 打包安装 | 发布与ZIP打包、SHA-256清单、独立安装目录、恢复版快捷方式、保留用户数据的卸载 | 所有 PowerShell/安装操作均未运行 |

核心源码是按已核对证据和现有数据流程重建；没有恢复完整原始 C#、注释或 Git 历史。

## 已核实的原程序行为

批量区间：1–20 全检；21–40→32；41–60→35；61–80→47；81–100→49；101–200→64；201–300→67；301–400→81；401–500→81；501–100000→98。允许 T1 数量依次为 0、1、1、2、2、3、3、4、4、5，允许 T2 均为0。

原单件分类：实测≥规格为 Pass；规格−允许短缺量≤实测<规格为 T1；更低为 T2。短缺量来自标准表 J 列。原批次计数函数只检查 T1 数量和是否存在 T2，**不检查平均值**。

这些是旧版实现事实，不是已核实的法规定义。恢复版分别显示原计数结果、平均值条件和数量偏差。感官结果不默认“正常”，生产日期不默认与检验日期相同。

需注意：21–31 件会得到32个应抽样数，存在应抽数大于批量的旧版问题；恢复版阻止这种数据导出。标准表 H 列文字要求与 J 列短缺量也可能不一致，需复核。重复标准条码、未知生产条目不会静默选取或丢弃。

原导出器实际写出各产品明细文件和一个汇总文件。恢复版采用合并工作簿，增加“测量追溯”和“报告说明”，对超过100条的明细分页保留全部记录；这是有意改动。

## Windows 第一轮操作

1. 安装 .NET 8 SDK（Windows x64），重新打开 PowerShell。
2. 切到本目录，运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Test-Windows.ps1 -Package
```

脚本先构建解决方案，再运行 Core 检查，成功后发布两个程序并生成带安装脚本的 ZIP。日志在 `artifacts/windows-check-*.log`；可直接启动：

```text
artifacts/windows/MainApp/Mro.MainApp.exe
artifacts/windows/HelperApp/Mro.HelperApp.exe
```

3. 主程序重新导入标准表。目录已升级为 `standard-catalog.v2.json`，旧恢复初版的 v1 不自动沿用。
4. 先对照两个窗口，再用真实生产表和称重表验证导入、统计与报告。详细步骤见 `docs/WINDOWS-TEST.md`。
5. 单独测试安装包：解压ZIP，运行 `Install-MRO.cmd`；关闭程序后再测试卸载脚本。构建脚本本身不执行安装。

应用数据：`%LOCALAPPDATA%\MRO-Recovered`。安装程序：`%LOCALAPPDATA%\Programs\MRO-Recovered\releases\...`。旧版应用和旧版快捷方式不作为安装目标。卸载保留标准库、路径偏好和日志。

若失败，保留完整构建日志；运行错误再看 `%LOCALAPPDATA%\MRO-Recovered\startup-errors.log`。优先修编译/启动，再核对数值，最后处理视觉与打印差异。

## 检查与证据

- `original-release/MRO-Windows-x64/`：保留的原始发布目录，含程序、DLL、PDB、Excel、资源及安装/卸载脚本。`original-release/sha256.json` 用于校验文件；只排除 macOS 的 `.DS_Store`。
- `scripts/verify_recovery.py`：XAML、资源、绑定、事件、工程引用和空白模板静态检查。
- `docs/syntax-checks.json`：C# 和 PowerShell 语法解析结果，不等于编译或执行结果。
- `tests/Mro.Core.Specs/Program.cs`：抽样边界、判定、生产表、偏好、模板分页和数据保留的运行检查；待 Windows/.NET 执行。
- `docs/recovered-business-rules.json`：规则与来源。
- `docs/report-template-recovery.json`：模板来源、排版恢复与有意差异。
- `docs/resource-manifest.json`：资源和工作簿哈希。
- `docs/recovery-evidence/supported-il.json`：所列导入、规则、设置和排版方法的静态 IL 证据，没有执行程序集。完整解析脚本保留在本地上一级 `recovery-assessment/`，构建不依赖它们。

## Git 交接

仓库根目录为本地 `MRO-Recovered/`。恢复工程源码与 `original-release/` 下的原始发布文件一同提交，便于 Windows 上继续对照；构建仍不依赖原 DLL。不要纳入 `bin/`、`obj/`、`artifacts/` 等生成目录。

上传目标：`https://github.com/BigLange/MRO_project.git`，本地分支为 `main`，工程根目录即仓库根目录。已准备本地提交与远程配置；当前执行环境无法连接 GitHub 的443端口，远端分支状态尚未核实，推送结果以实际命令输出为准。客户端新需求、正式法规核验和 Windows 实测修复属于下一节点。
