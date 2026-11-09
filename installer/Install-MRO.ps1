$ErrorActionPreference = "Stop"
$packageRoot = $PSScriptRoot
$installRoot = Join-Path $env:LOCALAPPDATA "Programs\MRO-Recovered"
$product = "MRO-Recovered"
$markerPath = Join-Path $installRoot ".managed-install.json"
$names = @("马里奥产品抽检分类系统（恢复版）", "马里奥产品抽检辅助系统（恢复版）")
$apps = @("Mro.MainApp", "Mro.HelperApp")
$selfContained = @($apps | Where-Object { -not (Test-Path -LiteralPath (Join-Path $packageRoot "$_\coreclr.dll")) }).Count -eq 0
if (-not $selfContained) {
    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) { throw "Install .NET 8 Desktop Runtime (Windows x64) first." }
    $runtimes = & dotnet --list-runtimes
    if ($LASTEXITCODE -ne 0 -or -not ($runtimes -match '^Microsoft.WindowsDesktop.App 8\.')) { throw "Install .NET 8 Desktop Runtime (Windows x64) first." }
}
$manifestPath = Join-Path $packageRoot "package-manifest.json"
if (-not (Test-Path -LiteralPath $manifestPath)) { throw "Package manifest is missing. Run scripts/Test-Windows.ps1 -Package first." }
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$sourcePrefix = [IO.Path]::GetFullPath($packageRoot).TrimEnd('\') + '\'
foreach ($file in $manifest) {
    $candidate = [IO.Path]::GetFullPath((Join-Path $packageRoot $file.Path))
    if (-not $candidate.StartsWith($sourcePrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Invalid package path." }
    if ((Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash -ne $file.Sha256) { throw "Package hash mismatch: $($file.Path)" }
}
foreach ($app in $apps) {
    if (-not (Test-Path -LiteralPath (Join-Path $packageRoot "$app\$app.exe"))) { throw "Published app missing: $app" }
}
if (Test-Path -LiteralPath $installRoot) {
    if (-not (Test-Path -LiteralPath $markerPath)) { throw "Existing folder is not a managed MRO-Recovered installation." }
    $marker = Get-Content -LiteralPath $markerPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($marker.Product -ne $product) { throw "Existing installation marker does not match." }
}
$installedPrefix = [IO.Path]::GetFullPath($installRoot).TrimEnd('\') + '\'
Get-CimInstance Win32_Process -Filter "Name = 'Mro.MainApp.exe' OR Name = 'Mro.HelperApp.exe'" | ForEach-Object {
    if ($_.ExecutablePath -and $_.ExecutablePath.StartsWith($installedPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Close both recovered apps before reinstalling." }
}
$version = (Get-Date -Format "yyyyMMdd-HHmmss") + "-" + [Guid]::NewGuid().ToString("N").Substring(0,8)
$releaseRoot = Join-Path $installRoot "releases\$version"
New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null
try {
    foreach ($app in $apps) { Copy-Item -LiteralPath (Join-Path $packageRoot $app) -Destination (Join-Path $releaseRoot $app) -Recurse }
    Copy-Item -LiteralPath (Join-Path $packageRoot "Uninstall-MRO.ps1") -Destination (Join-Path $installRoot "Uninstall-MRO.ps1") -Force
    $desktop = [Environment]::GetFolderPath("Desktop")
    $shell = New-Object -ComObject WScript.Shell
    for ($i = 0; $i -lt $apps.Count; $i++) {
        $shortcutPath = Join-Path $desktop ($names[$i] + ".lnk")
        $shortcut = $shell.CreateShortcut($shortcutPath)
        if ((Test-Path -LiteralPath $shortcutPath) -and $shortcut.TargetPath -and -not $shortcut.TargetPath.StartsWith($installedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "An unrelated desktop shortcut already uses this name: $shortcutPath"
        }
        $target = Join-Path $releaseRoot ($apps[$i] + "\" + $apps[$i] + ".exe")
        $shortcut.TargetPath = $target
        $shortcut.WorkingDirectory = Split-Path -Parent $target
        $shortcut.IconLocation = "$target,0"
        $shortcut.Save()
    }
    @{ Product = $product; Version = $version; Release = $releaseRoot } | ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding UTF8
}
catch {
    # Keep the staged release for diagnosis; do not delete or overwrite a previous release.
    throw "Installation did not finish. Previous releases and user data were retained. $($_.Exception.Message)"
}
Write-Host "Installed to $releaseRoot"
Write-Host "User data and preferences remain in $env:LOCALAPPDATA\MRO-Recovered"
