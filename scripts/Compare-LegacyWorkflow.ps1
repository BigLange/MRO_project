#requires -Version 7.4
param()
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$original = [Reflection.Assembly]::LoadFile((Join-Path $projectRoot 'original-release\MRO-Windows-x64\Mro.MainApp\Mro.Core.dll'))
$recovered = [Reflection.Assembly]::LoadFile((Join-Path $projectRoot 'src\Mro.Core\bin\Release\net8.0\Mro.Core.dll'))
function New-TypedList([Type]$ElementType) {
    $listType = [Collections.Generic.List``1].MakeGenericType([Type[]]@($ElementType))
    return ,([Activator]::CreateInstance($listType))
}
$oldSampleType = $original.GetType('Mro.Core.Rules.InspectionSample')
$newSampleType = $recovered.GetType('Mro.Core.Measurement')
$judge = $original.GetType('Mro.Core.Rules.NetContentJudgement').GetMethod('Evaluate')
$catalog = $recovered.GetType('Mro.Core.StandardImporter').GetMethod('Read').Invoke($null,
    [object[]]@([string](Join-Path $projectRoot 'SharedContent\检测标准表.xlsx'), $null))
$barcode = '0011010116'
$cases = @(
    @{ Weights = @(91,89,90); Batch = 40 },
    @{ Weights = @(91,89,82,92,93); Batch = 3 },
    @{ Weights = @(87,88,89); Batch = 40 },
    @{ Weights = @(90) * 32; Batch = 40 },
    @{ Weights = @(91,89,90); Batch = 30 },
    @{ Weights = @(91) * 101; Batch = 1000 }
)
$checked = 0
foreach ($seed in @(1,7,29,1234)) {
    foreach ($case in $cases) {
        $script:legacyRandom = [Random]::new($seed)
        $script:recoveredRandom = [Random]::new($seed)
        $oldGenerator = [Activator]::CreateInstance($original.GetType('Mro.Core.Rules.RandomCompliantWeightGenerator'),
            [object[]]@([Func[double]]{ $script:legacyRandom.NextDouble() }))
        $adjuster = [Activator]::CreateInstance($original.GetType('Mro.Core.Rules.SampleCountAdjuster'), [object[]]@($oldGenerator))
        $oldSamples = New-TypedList $oldSampleType
        $newSamples = New-TypedList $newSampleType
        for ($i = 0; $i -lt $case.Weights.Count; $i++) {
            $weight = [decimal]$case.Weights[$i]
            $class = $judge.Invoke($null,[object[]]@([decimal]90,[decimal]6.8,$weight))
            $oldSamples.Add([Activator]::CreateInstance($oldSampleType,[object[]]@("TEST-$i",$weight,$class,$false)))
            $newSamples.Add([Activator]::CreateInstance($newSampleType,[object[]]@($barcode,$weight,'2026.11.09','12:00:00',
                [decimal]0,$weight,'TEST-ONLY','原始数据',($i+2),"TEST-$i",'导入记录','',$null,'','',$null)))
        }
        $required = $catalog.RequiredSamples([int]$case.Batch)
        $expected = $adjuster.Adjust($oldSamples,$required,[decimal]90,[decimal]6.8)
        $batches = [Collections.Generic.Dictionary[string,string]]::new()
        $batches.Add($barcode, [string]$case.Batch)
        $actual = $recovered.GetType('Mro.Core.LegacyInspection').GetMethod('Prepare').Invoke($null,
            [object[]]@($catalog,$newSamples,$batches,[Func[double]]{ $script:recoveredRandom.NextDouble() }))[0].Measurements
        if ($expected.Count -ne $actual.Count) { throw 'Adjusted count mismatch.' }
        for ($i = 0; $i -lt $expected.Count; $i++) {
            if ($expected[$i].ActualWeightGrams -ne $actual[$i].Grams) { throw "Weight/order mismatch: seed $seed batch $($case.Batch) index $i" }
        }
        $checked++
    }
}
Write-Host "PASS $checked deterministic supplementation/trimming cases against the original DLL."
