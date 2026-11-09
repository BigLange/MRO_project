$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$published = Join-Path $projectRoot "artifacts\windows"
$stamp = (Get-Date -Format "yyyyMMdd-HHmmss") + "-" + [Guid]::NewGuid().ToString("N").Substring(0,8)
$packageRoot = Join-Path $projectRoot ("artifacts\MRO-Recovered-Windows-x64-" + $stamp)
foreach ($app in @("MainApp", "HelperApp")) {
    if (-not (Test-Path -LiteralPath (Join-Path $published "$app\Mro.$app.exe"))) { throw "Missing published $app. Run Test-Windows.ps1 -Package." }
}
New-Item -ItemType Directory -Path $packageRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $published "MainApp") -Destination (Join-Path $packageRoot "Mro.MainApp") -Recurse
Copy-Item -LiteralPath (Join-Path $published "HelperApp") -Destination (Join-Path $packageRoot "Mro.HelperApp") -Recurse
foreach ($file in @("Install-MRO.ps1", "Install-MRO.cmd", "Uninstall-MRO.ps1")) {
    Copy-Item -LiteralPath (Join-Path $projectRoot "installer\$file") -Destination (Join-Path $packageRoot $file)
}
$prefix = [IO.Path]::GetFullPath($packageRoot).TrimEnd('\') + '\'
$manifest = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File | ForEach-Object {
    @{ Path = $_.FullName.Substring($prefix.Length).Replace('\','/'); Sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
})
ConvertTo-Json -InputObject $manifest -Depth 4 | Set-Content -LiteralPath (Join-Path $packageRoot "package-manifest.json") -Encoding UTF8
$zipPath = Join-Path $projectRoot ("artifacts\MRO-Recovered-Windows-x64-" + $stamp + ".zip")
Compress-Archive -LiteralPath $packageRoot -DestinationPath $zipPath
Write-Host "Package: $zipPath"
