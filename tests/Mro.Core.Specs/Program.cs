using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Mro.Core;

// Explicit test fixtures, never production measurements. No random data generator.
var temporary = Path.Combine(Path.GetTempPath(), "Mro-Recovered-Specs-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(temporary);
var passed = 0;
void Assert(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
void Reject(Action action, string message)
{
    try { action(); }
    catch (Exception ex) when (ex is InvalidDataException or IOException) { return; }
    throw new Exception(message);
}
void Test(string name, Action action)
{
    action(); passed++; Console.WriteLine("PASS " + name);
}
string Temp(string name) => Path.Combine(temporary, name);
string Fixture(string name, string unit, params decimal[] values)
{
    var rows = new List<object?[]> { ["序号", "日期", "时间", "皮重", "重量", "毛重", "单位", "扫描"] };
    rows.AddRange(values.Select((v, i) => new object?[] { i + 1, "TEST-ONLY", $"12:00:{i:00}", 0, v, v, unit, "0011010116" }));
    var path = Temp(name);
    XlsxWriter.Write(path, [new("测试数据_非实际测量", rows)]);
    return path;
}

try
{
    var standard = Path.Combine(AppContext.BaseDirectory, "SharedContent", "检测标准表.xlsx");
    var catalog = StandardImporter.Read(standard);
    Test("standard table: product rows separated from rules, leading-zero identifiers preserved", () =>
    {
        Assert(catalog.Products.Count == 106, "Expected 106 products.");
        Assert(catalog.Resolve("0011010116").ProductName == "迷你毛毛虫（面包）", "Barcode changed.");
        Assert(catalog.Products.Any(p => p.Barcode == "4801000353"), "Numeric barcode changed.");
        Assert(catalog.ImportWarnings.Any(w => w.Contains("重复条码")), "Missing duplicate warning.");
        Reject(() => catalog.Resolve("0011010178"), "Ambiguous barcode must not resolve.");
        Reject(() => catalog.Resolve("AF12566"), "Unknown barcode must not resolve.");
    });
    Test("sampling: recovered upper-bound intervals and limits", () =>
    {
        Assert(catalog.RequiredSamples(3) == 3 && catalog.RequiredSamples(40) == 32, "Explicit rules changed.");
        Assert(catalog.RequiredSamples(30) == 32, "Recovered rule is an upper bound, not a single typical value.");
        (int Lower, int Upper, int Count, int T1)[] recovered = [
            (21,40,32,1), (41,60,35,1), (61,80,47,2), (81,100,49,2), (101,200,64,3),
            (201,300,67,3), (301,400,81,4), (401,500,81,4), (501,100000,98,5) ];
        foreach (var (lower, upper, count, t1) in recovered)
        {
            Assert(catalog.RequiredSamples(lower) == count && catalog.RequiredSamples(upper) == count, $"Boundary error at {lower}/{upper}.");
            Assert(catalog.SamplingRuleFor(lower).AllowedT1Count == t1, "Recovered T1 allowance changed.");
        }
        Reject(() => catalog.RequiredSamples(100001), "Out-of-range batch accepted.");
        Reject(() => StandardImporter.ValidateRules([new(1, 20, 0, true), new(10, 40, 5)]), "Overlapping plans accepted.");
    });
    var inputPath = Fixture("test-input-kg.xlsx", " kg", 0.091m, 0.089m, 0.090m);
    var inputHash = DataFiles.Sha256(inputPath);
    var measurements = MeasurementImporter.Read(inputPath);
    var batches = new Dictionary<string, string> { ["0011010116"] = "3" };
    Test("measured import: kg to g, all values retained including below-standard measurements", () =>
    {
        Assert(measurements.Select(m => m.Grams).SequenceEqual(new decimal[] { 91m, 89m, 90m }), "Measurement altered.");
        Assert(measurements.All(m => m.Barcode == "0011010116"), "Leading zeros lost.");
        Assert(measurements[0].SourceRow == 2 && measurements[2].SourceRow == 4, "Source row lost.");
        Assert(InspectionValidator.Validate(catalog, measurements, batches)[0].AverageGrams == 90m, "Average must include low result.");
    });
    Test("insufficient, unknown, duplicate and invalid data block export", () =>
    {
        Reject(() => WorkbookExports.Report(Temp("must-not-exist.xlsx"), catalog, measurements,
            new Dictionary<string, string> { ["0011010116"] = "40" }, DateTime.Today), "Missing measurements must block report.");
        Assert(!File.Exists(Temp("must-not-exist.xlsx")), "Invalid report was created.");
        Reject(() => InspectionValidator.Validate(catalog, measurements, new Dictionary<string, string> { ["0011010116"] = "0" }), "Zero batch accepted.");
        Reject(() => InspectionValidator.Validate(catalog, measurements.Select(m => m with { Barcode = "UNKNOWN" }).ToList(),
            new Dictionary<string, string> { ["UNKNOWN"] = "3" }), "Unknown barcode accepted.");
        Reject(() => InspectionValidator.Validate(catalog, [measurements[0], measurements[0], measurements[2]], batches), "Duplicate record accepted.");
        Reject(() => MeasurementImporter.Read(Fixture("negative.xlsx", "g", -1m)), "Negative weight accepted.");
        Reject(() => MeasurementImporter.Read(Fixture("bad-unit.xlsx", "???", 91m)), "Unknown unit accepted.");
    });
    Test("formula values are rejected rather than trusting cached values", () =>
    {
        var formulaFile = Fixture("formula.xlsx", "g", 91m);
        using (var zip = ZipFile.Open(formulaFile, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument document;
            using (var stream = entry.Open()) document = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            document.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == "E2").AddFirst(new XElement(ns + "f", "90+1"));
            entry.Delete();
            using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            document.Save(output);
        }
        Reject(() => MeasurementImporter.Read(formulaFile), "Formula accepted as real measurement.");
    });
    Test("report: summary and per-product detail agree, source rows preserved", () =>
    {
        WorkbookExports.Report(Temp("test-report.xlsx"), catalog, measurements, batches, new(2026, 11, 9));
        var sheets = XlsxReader.Read(Temp("test-report.xlsx"));
        Assert(sheets.Count == 3, "Expected summary, detail, metadata sheets.");
        var summary = sheets[0].Rows.Single(r => r.Number == 3);
        Assert(summary.At(1) == "0011010116" && summary.At(7) == "3", "Summary identifiers/count changed.");
        Assert(decimal.Parse(summary.At(8), CultureInfo.InvariantCulture) == 90m, "Summary average changed.");
        var detail = sheets[1].Rows.Where(r => r.Number >= 3).ToArray();
        Assert(detail.Length == 3 && detail[1].At(4) == "89", "Low result lost in report.");
        Assert(detail[1].At(9) == "3" && detail[1].At(10) == measurements[1].RecordId, "Source trace lost.");
    });
    Test("helper export/import round trip preserves measurement values and record identity", () =>
    {
        var output = Temp("test-roundtrip.xlsx");
        WorkbookExports.Measurements(output, catalog, measurements, batches);
        var again = MeasurementImporter.Read(output);
        Assert(again.SequenceEqual(measurements), "Helper export changed source records.");
    });
    Test("legacy judging boundaries, mean condition and record-count mismatch remain visible", () =>
    {
        Assert(RecoveredJudgement.Classify(100m, 5m, 100m) == LegacySampleJudgement.Pass, "Nominal boundary changed.");
        Assert(RecoveredJudgement.Classify(100m, 5m, 95m) == LegacySampleJudgement.T1, "Shortage boundary changed.");
        Assert(RecoveredJudgement.Classify(100m, 5m, 94.999m) == LegacySampleJudgement.T2, "T2 boundary changed.");
        var product = catalog.Resolve("0011010116") with { StandardGrams = 100m, AllowedShortageGrams = 5m };
        var belowMean = Enumerable.Range(0, 32).Select(i => measurements[0] with { Grams = i == 0 ? 99m : 100m, RecordId = "TEST-MEAN-" + i }).ToList();
        var metrics = RecoveredJudgement.Evaluate(new(product, 40, 32, belowMean), catalog.SamplingRuleFor(40));
        Assert(metrics.LegacyCountPass == true && !metrics.AverageMeetsNominal, "Original count result and mean must be separate.");
        Assert(metrics.ReviewConclusion.Contains("未满足"), "Low mean concealed by count-only pass.");
        var extra = belowMean.Select(m => m with { Grams = 101m }).Append(measurements[0] with { Grams = 101m, RecordId = "TEST-EXTRA" }).ToList();
        Assert(RecoveredJudgement.Evaluate(new(product, 40, 32, extra), catalog.SamplingRuleFor(40)).ReviewConclusion.Contains("计划数不同"), "Excess records concealed.");
    });
    Test("production import: original A/B layout, named code layout, duplicates and unknown identifiers", () =>
    {
        var legacy = Temp("production-legacy.xlsx");
        XlsxWriter.Write(legacy, [new("TEST-ONLY", [["0011010116", 3]])]);
        var entries = ProductionImporter.Read(legacy, catalog);
        Assert(entries.Single().BatchSize == 3 && entries[0].Barcode == "0011010116", "Original A/B input changed.");
        var named = Temp("production-named.xlsx");
        XlsxWriter.Write(named, [new("TEST-ONLY", [["编码", "生产数量", "生产日期"], ["M2", 3, "2026-11-08"]])]);
        var namedEntry = ProductionImporter.Read(named, catalog).Single();
        Assert(namedEntry.Barcode == "0011010116" && namedEntry.ProductionDate == "2026-11-08", "Named code input failed.");
        var duplicated = Temp("production-duplicate.xlsx");
        XlsxWriter.Write(duplicated, [new("TEST-ONLY", [["0011010116", 3], ["0011010116", 4]])]);
        Reject(() => ProductionImporter.Read(duplicated, catalog), "Duplicate batches were silently combined.");
        var unknown = Temp("production-unknown.xlsx");
        XlsxWriter.Write(unknown, [new("TEST-ONLY", [["UNKNOWN", 3]])]);
        Reject(() => ProductionImporter.Read(unknown, catalog), "Unknown production row silently skipped.");
        var zero = Temp("production-zero.xlsx");
        XlsxWriter.Write(zero, [new("TEST-ONLY", [["0011010116", 0]])]);
        Reject(() => ProductionImporter.Read(zero, catalog), "Zero production quantity accepted.");
    });
    Test("preferences: updates preserve other app fields, damaged JSON backed up", () =>
    {
        var path = Temp("preferences.json"); var store = new UserFolderPreferenceStore(path);
        store.Update(old => old with { MainStandardWorkbookPath = standard });
        new UserFolderPreferenceStore(path).Update(old => old with { HelperEntryWorkbookDirectory = temporary });
        Assert(store.Load().MainStandardWorkbookPath == standard && store.Load().HelperEntryWorkbookDirectory == temporary, "Other app preferences overwritten.");
        File.WriteAllText(path, "{broken");
        store.Load(); Assert(store.LoadWarning is not null, "Corrupt preference warning missing.");
        store.Update(old => old with { MainOutputDirectory = temporary });
        Assert(Directory.GetFiles(temporary, "preferences.json.corrupt-*").Length == 1, "Corrupt settings not backed up.");
        Assert(UserFolderPreferenceStore.GetExistingFile(Temp("nonexistent.xlsx")) is null, "Missing file restored.");
    });
    Test("template report: styles retained, historical results cleared, all records paginated", () =>
    {
        var template = Path.Combine(AppContext.BaseDirectory, "SharedContent", "ReportTemplates", "MRO-ReportTemplate.xlsx");
        var output = Temp("test-template-report.xlsx");
        var full = Enumerable.Range(0, 101).Select(i => measurements[0] with { Grams = 90m, SourceRow = i + 2, RecordId = "TEST-PAGE-" + i }).ToList();
        WorkbookExports.Report(output, catalog, full, new Dictionary<string, string> { ["0011010116"] = "1000" }, new(2026,11,9), template);
        var sheets = XlsxReader.Read(output);
        Assert(sheets.Count == 5, "Expected summary, two detail pages, trace and notes.");
        Assert(sheets[0].Rows.Single(r => r.Number == 7).At(3) == "迷你毛毛虫（面包）", "Summary mapped to wrong cell.");
        Assert(sheets[1].Rows.Single(r => r.Number == 3).At(1).StartsWith("样品1."), "First detail sample misplaced.");
        Assert(sheets[2].Rows.Single(r => r.Number == 3).At(1).StartsWith("样品101."), "Samples beyond 100 lost.");
        Assert(sheets[3].Rows.Count(r => r.Number >= 3) == 101, "Trace does not contain all measurements.");
        Assert(sheets[0].Rows.Single(r => r.Number == 7).At(9) == "未录入", "Sensory result fabricated.");
        Assert(sheets[1].Rows.Single(r => r.Number == 33).At(3) == "未提供", "Production date fabricated.");
        Assert(sheets.SelectMany(s => s.Rows).All(r => r.Cells.Values.All(c => !c.Text.Contains("66.1g"))), "Historical result leaked.");
        Assert(TemplateReportExporter.DetailSlot(1) == "A3" && TemplateReportExporter.DetailSlot(26) == "C3" && TemplateReportExporter.DetailSlot(100) == "G27", "Original sample layout changed.");
        using var original = ZipFile.OpenRead(template); using var exported = ZipFile.OpenRead(output);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement Load(ZipArchive archive, string name) { using var stream = archive.GetEntry(name)!.Open(); return XElement.Load(stream); }
        Assert(XNode.DeepEquals(Load(original, "xl/worksheets/sheet2.xml").Element(ns+"mergeCells"), Load(exported, "xl/worksheets/sheet2.xml").Element(ns+"mergeCells")), "Detail merges changed.");
        Assert(XNode.DeepEquals(Load(original, "xl/worksheets/sheet2.xml").Element(ns+"pageMargins"), Load(exported, "xl/worksheets/sheet2.xml").Element(ns+"pageMargins")), "Detail margins changed.");
        Assert(!exported.Entries.Any(e => e.FullName.Contains("externalLink")), "Historical external link retained.");
    });
    Test("literal text is not exported as formula, existing files and originals never overwritten", () =>
    {
        var literal = Temp("literal.xlsx");
        XlsxWriter.Write(literal, [new("文本", [["=1+1"]])]);
        Assert(XlsxReader.Read(literal)[0].Rows[0].At(1) == "=1+1", "Text interpreted as formula.");
        Reject(() => XlsxWriter.Write(inputPath, [new("x", [["overwrite"]])]), "Input overwritten.");
        Assert(DataFiles.Sha256(inputPath) == inputHash, "Input file changed.");
    });
    Console.WriteLine($"All {passed} checks passed. Test files are marked TEST-ONLY and will be deleted.");
    Directory.Delete(temporary, true);
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"FAILED after {passed} checks. {ex}\nDiagnostic fixtures (TEST-ONLY): {temporary}");
    return 1;
}
