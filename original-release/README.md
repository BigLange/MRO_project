# 原始 Windows 发布文件

`MRO-Windows-x64/` 是保留发布目录的逐字节副本，包含两个应用的 EXE、DLL、PDB、运行配置、Excel、图标及安装/卸载脚本。仅排除 macOS 的 `.DS_Store`。

此目录用于继续核对原程序的资源、结构和行为。恢复工程源码位于仓库根目录的 `apps/`、`src/` 等目录；新恢复版安装包需要在 Windows 构建后生成。

原发布包安装入口为 `MRO-Windows-x64/Install-MRO.cmd`，原程序文件位于其 `Mro.MainApp/`、`Mro.HelperApp/` 子目录；Git 提交过程不执行这些程序或脚本。

`sha256.json` 记录所有25个发布文件的大小与 SHA-256。Git 属性已禁止此目录的文本换行转换，Windows 检出后应保持原字节。

在仓库根目录的 PowerShell 中校验：

```powershell
$releaseRoot = Join-Path (Get-Location) "original-release\MRO-Windows-x64"
$manifest = Get-Content ".\original-release\sha256.json" -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($file in $manifest.files) {
    $path = Join-Path $releaseRoot $file.path
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $file.sha256) {
        throw "File mismatch: $($file.path)"
    }
}
Write-Host "All original release hashes match."
```
