$ErrorActionPreference = "Stop"

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultInstallRoot = Join-Path $env:LOCALAPPDATA "MRO"
$installRoot = $defaultInstallRoot
$mainShortcutName = "马里奥产品抽检分类系统"
$helperShortcutName = "马里奥产品抽检辅助系统"

function Stop-InstalledAppProcesses([string]$root)
{
    if (-not (Test-Path $root))
    {
        return
    }

    $rootFull = [System.IO.Path]::GetFullPath($root).TrimEnd("\") + "\"
    $processNames = @("Mro.MainApp.exe", "Mro.HelperApp.exe")
    foreach ($processName in $processNames)
    {
        Get-CimInstance Win32_Process -Filter "Name = '$processName'" | ForEach-Object {
            $path = $_.ExecutablePath
            if ([string]::IsNullOrWhiteSpace($path))
            {
                return
            }

            $processFull = [System.IO.Path]::GetFullPath($path)
            if (-not $processFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase))
            {
                return
            }

            Write-Host "Closing running app: $processName"
            Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
            Wait-Process -Id $_.ProcessId -Timeout 10 -ErrorAction SilentlyContinue
        }
    }
}

function Remove-PathWithRetry([string]$path)
{
    if (-not (Test-Path $path))
    {
        return
    }

    for ($attempt = 1; $attempt -le 3; $attempt++)
    {
        try
        {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
            return
        }
        catch
        {
            if ($attempt -eq 3)
            {
                throw "Could not replace '$path'. Close the MRO apps and run Install-MRO.cmd again. Original error: $($_.Exception.Message)"
            }

            Start-Sleep -Seconds 1
        }
    }
}

function Assert-InstalledFile([string]$path)
{
    if (-not (Test-Path $path))
    {
        throw "Required installed file is missing: $path"
    }
}

if (-not (Test-Path $installRoot))
{
    New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
}

$mainSource = Join-Path $packageRoot "Mro.MainApp"
$helperSource = Join-Path $packageRoot "Mro.HelperApp"
$sharedContentSource = Join-Path $packageRoot "SharedContent"
$assetsSource = Join-Path $packageRoot "Assets"
$mainTarget = Join-Path $installRoot "Mro.MainApp"
$helperTarget = Join-Path $installRoot "Mro.HelperApp"
$sharedContentTarget = Join-Path $installRoot "SharedContent"
$assetsTarget = Join-Path $installRoot "Assets"
$mainShortcutIconTarget = Join-Path $assetsTarget "mro-main-red-icon.ico"
$helperShortcutIconTarget = Join-Path $assetsTarget "mro-helper-green-icon.ico"

if (-not (Test-Path $mainSource) -or -not (Test-Path $helperSource))
{
    throw "Install files were not found. Extract the zip package first, then run Install-MRO.cmd from the extracted folder."
}

Stop-InstalledAppProcesses $installRoot

Remove-PathWithRetry $mainTarget
Remove-PathWithRetry $helperTarget
Remove-PathWithRetry $sharedContentTarget
Remove-PathWithRetry $assetsTarget

Copy-Item -LiteralPath $mainSource -Destination $mainTarget -Recurse
Copy-Item -LiteralPath $helperSource -Destination $helperTarget -Recurse
if (Test-Path $sharedContentSource)
{
    Copy-Item -LiteralPath $sharedContentSource -Destination $sharedContentTarget -Recurse
}
if (Test-Path $assetsSource)
{
    Copy-Item -LiteralPath $assetsSource -Destination $assetsTarget -Recurse
}

Assert-InstalledFile (Join-Path $mainTarget "Mro.MainApp.exe")
Assert-InstalledFile (Join-Path $helperTarget "Mro.HelperApp.exe")
Assert-InstalledFile (Join-Path $sharedContentTarget "检测标准表.xlsx")
Assert-InstalledFile (Join-Path $sharedContentTarget "W250905.xlsx")
Assert-InstalledFile (Join-Path $sharedContentTarget "检测报告1.xlsx")
Assert-InstalledFile (Join-Path $sharedContentTarget "检测报告2.xlsx")
Assert-InstalledFile $mainShortcutIconTarget
Assert-InstalledFile $helperShortcutIconTarget

$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell

$mainShortcutPath = Join-Path $desktop "$mainShortcutName.lnk"
$helperShortcutPath = Join-Path $desktop "$helperShortcutName.lnk"
if (Test-Path $mainShortcutPath) { Remove-Item -LiteralPath $mainShortcutPath -Force }
if (Test-Path $helperShortcutPath) { Remove-Item -LiteralPath $helperShortcutPath -Force }

$mainShortcut = $shell.CreateShortcut($mainShortcutPath)
$mainShortcut.TargetPath = Join-Path $mainTarget "Mro.MainApp.exe"
$mainShortcut.WorkingDirectory = $mainTarget
$mainShortcut.IconLocation = "$mainShortcutIconTarget,0"
$mainShortcut.Save()

$helperShortcut = $shell.CreateShortcut($helperShortcutPath)
$helperShortcut.TargetPath = Join-Path $helperTarget "Mro.HelperApp.exe"
$helperShortcut.WorkingDirectory = $helperTarget
$helperShortcut.IconLocation = "$helperShortcutIconTarget,0"
$helperShortcut.Save()

Write-Host "Installed to: $installRoot"
Write-Host "Desktop shortcuts created:"
Write-Host "  $mainShortcutName"
Write-Host "  $helperShortcutName"
