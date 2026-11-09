param([switch]$Package)
$ErrorActionPreference = "Stop"
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "Install the .NET 8 SDK for Windows x64, then reopen PowerShell. Desktop Runtime alone cannot build this project."
}
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location $projectRoot
$transcriptStarted = $false
try {
    New-Item -ItemType Directory -Path "artifacts" -Force | Out-Null
    $logPath = Join-Path $projectRoot ("artifacts\windows-check-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
    Start-Transcript -Path $logPath | Out-Null
    $transcriptStarted = $true
    function Invoke-Dotnet([string[]]$CommandArguments) {
        & dotnet @CommandArguments
        if ($LASTEXITCODE -ne 0) { throw "dotnet failed with exit code $LASTEXITCODE. See $logPath" }
    }
    Invoke-Dotnet -CommandArguments @("--info")
    Invoke-Dotnet -CommandArguments @("build", "Mro.sln", "-c", "Release", "--nologo")
    Invoke-Dotnet -CommandArguments @("run", "--project", "tests/Mro.Core.Specs", "-c", "Release", "--no-build")
    if ($Package) {
        Invoke-Dotnet -CommandArguments @("publish", "apps/Mro.MainApp", "-c", "Release", "-r", "win-x64", "--self-contained", "false", "-o", "artifacts/windows/MainApp")
        Invoke-Dotnet -CommandArguments @("publish", "apps/Mro.HelperApp", "-c", "Release", "-r", "win-x64", "--self-contained", "false", "-o", "artifacts/windows/HelperApp")
        Write-Host "Published: artifacts/windows/MainApp and artifacts/windows/HelperApp"
        & (Join-Path $PSScriptRoot "Package-Windows.ps1")
    }
    Write-Host "Build and Core checks passed. Window rendering still needs the manual checklist in docs/WINDOWS-TEST.md."
    Write-Host "Log: $logPath"
}
finally {
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
    Pop-Location
}
