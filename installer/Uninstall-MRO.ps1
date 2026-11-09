$ErrorActionPreference = "Stop"
$installRoot = Join-Path $env:LOCALAPPDATA "Programs\MRO-Recovered"
$markerPath = Join-Path $installRoot ".managed-install.json"
if (-not (Test-Path -LiteralPath $markerPath)) { throw "No managed MRO-Recovered installation found." }
$marker = Get-Content -LiteralPath $markerPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($marker.Product -ne "MRO-Recovered") { throw "Installation marker does not match." }
$prefix = [IO.Path]::GetFullPath($installRoot).TrimEnd('\') + '\'
Get-CimInstance Win32_Process -Filter "Name = 'Mro.MainApp.exe' OR Name = 'Mro.HelperApp.exe'" | ForEach-Object {
    if ($_.ExecutablePath -and $_.ExecutablePath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Close both recovered apps before uninstalling." }
}
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell
foreach ($name in @("马里奥产品抽检分类系统（恢复版）", "马里奥产品抽检辅助系统（恢复版）")) {
    $shortcutPath = Join-Path $desktop ($name + ".lnk")
    if (Test-Path -LiteralPath $shortcutPath) {
        $shortcut = $shell.CreateShortcut($shortcutPath)
        if ($shortcut.TargetPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $shortcutPath }
    }
}
Remove-Item -LiteralPath $installRoot -Recurse -Force
Write-Host "Recovered apps removed. User data in $env:LOCALAPPDATA\MRO-Recovered was retained."
