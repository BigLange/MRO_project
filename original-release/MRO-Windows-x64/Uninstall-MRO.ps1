$ErrorActionPreference = "Stop"

$installRoot = Join-Path $env:LOCALAPPDATA "MRO"
$desktop = [Environment]::GetFolderPath("Desktop")
$mainShortcutName = "马里奥产品抽检分类系统"
$helperShortcutName = "马里奥产品抽检辅助系统"

if (Test-Path $installRoot)
{
    Remove-Item -LiteralPath $installRoot -Recurse -Force
}

$mainShortcut = Join-Path $desktop "$mainShortcutName.lnk"
$helperShortcut = Join-Path $desktop "$helperShortcutName.lnk"

if (Test-Path $mainShortcut) { Remove-Item -LiteralPath $mainShortcut -Force }
if (Test-Path $helperShortcut) { Remove-Item -LiteralPath $helperShortcut -Force }

Write-Host "MRO desktop installation removed."
