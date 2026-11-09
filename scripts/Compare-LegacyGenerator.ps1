#requires -Version 7.4
param()
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$original = [Reflection.Assembly]::LoadFile((Join-Path $projectRoot 'original-release\MRO-Windows-x64\Mro.HelperApp\Mro.Core.dll'))
$recovered = [Reflection.Assembly]::LoadFile((Join-Path $projectRoot 'src\Mro.Core\bin\Release\net8.0\Mro.Core.dll'))
$cases = 0
$changedCases = 0
foreach ($seed in @(1, 7, 29, 1234)) {
    foreach ($count in @(1, 2, 3, 7, 32, 98)) {
        $script:legacyRandom = [Random]::new($seed)
        $script:recoveredRandom = [Random]::new($seed)
        $old = [Activator]::CreateInstance($original.GetType('Mro.Core.Helper.RandomPassWeightGenerator'), [object[]]@([Func[double]]{ $script:legacyRandom.NextDouble() }))
        $new = [Activator]::CreateInstance($recovered.GetType('Mro.Core.PassWeightGenerator'), [object[]]@([Func[double]]{ $script:recoveredRandom.NextDouble() }))
        $expected = $old.GenerateWeights([decimal]90, [decimal]99, $count)
        $actual = $new.Generate([decimal]90, [decimal]99, $count)
        if ($expected.Count -ne $actual.Count) { throw 'Count mismatch.' }
        $low = @($actual | Where-Object { $_ -le [decimal]92.25 }).Count
        $high = @($actual | Where-Object { $_ -ge [decimal]96.75 }).Count
        $quota = [int][Math]::Round($count * 0.8, [MidpointRounding]::AwayFromZero)
        if ($low + $high -ne $quota -or [Math]::Abs($low - $high) -gt 1) { throw "80/20 quota mismatch: seed $seed, count $count." }
        if (@($actual | Where-Object { $_ -lt 90 -or $_ -gt 99 }).Count) { throw 'Weight outside bounds.' }
        if (($expected -join ',') -ne ($actual -join ',')) { $changedCases++ }
        $cases++
    }
}
Write-Host "PASS $cases deterministic 80/20 quota cases; $changedCases differ from the original generator (intentional user change on 2026-11-10)."
