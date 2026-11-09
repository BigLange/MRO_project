using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Mro.Core;

var temporary = Path.Combine(Path.GetTempPath(), "Mro-Legacy-Specs-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(temporary);
var passed = 0;
void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
void Reject(Action action, string message)
{
    try { action(); } catch (Exception ex) when (ex is InvalidDataException or IOException) { return; }
    throw new Exception(message);
}
void Test(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
string Temp(string name) => Path.Combine(temporary, name);
string Content(string name) => Path.Combine(AppContext.BaseDirectory, "SharedContent", name);
XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
XElement Xml(string path, string entry) { using var z = ZipFile.OpenRead(path); using var stream = z.GetEntry(entry)!.Open(); return XElement.Load(stream); }
try
{
    var standard = Content("检测标准表.xlsx");
    var originalHash = DataFiles.Sha256(standard);
    var catalog = StandardImporter.Read(standard);
    var product = catalog.Resolve("0011010116");
    var reportTemplate = Content(Path.Combine("ReferenceReports", "检测报告2.xlsx"));
    var detailTemplate = Content(Path.Combine("ReferenceReports", "检测报告1.xlsx"));
    List<Measurement> Records(params decimal[] weights) => weights.Select((w, i) => new Measurement(product.Barcode, w, "2026.11.09", $"12:00:{i:00}", 0m, w,
        "TEST-ONLY", "原始数据", i + 2, "TEST-" + i)).ToList();
    Dictionary<string, string> Batches(int batch) => new() { [product.Barcode] = batch.ToString() };
    var measurements = Records(91m, 89m, 90m);
    string RatioWorkbook(string label, string value, bool numeric = false, bool percentage = false, string? formula = null)
    {
        var path = Temp(label + ".xlsx");
        File.Copy(standard, path);
        var sheets = XlsxReader.Read(standard);
        var index = sheets.FindIndex(s => s.Rows.Any(r => r.Cells.Values.Any(c => c.Text == "浮动比例")));
        var sheet = sheets[index];
        var column = sheet.Rows.SelectMany(r => r.Cells).First(c => c.Value.Text == "浮动比例").Key;
        var row = sheet.Rows.First(r => r.Cells.Values.Any(c => c.Text == product.Barcode)).Number;
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        XElement Read(string part) { using var stream = zip.GetEntry(part)!.Open(); return XElement.Load(stream); }
        void Save(string part, XElement xml) { zip.GetEntry(part)!.Delete(); using var stream = zip.CreateEntry(part).Open(); xml.Save(stream); }
        var part = $"xl/worksheets/sheet{index + 1}.xml";
        var xml = Read(part);
        var cell = xml.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == XlsxWriter.ColumnName(column) + row);
        cell.RemoveNodes(); cell.Attribute("s")?.Remove(); cell.Attribute("t")?.Remove();
        if (percentage)
        {
            var styles = Read("xl/styles.xml"); var xfs = styles.Element(ns + "cellXfs")!;
            cell.SetAttributeValue("s", xfs.Elements().Count());
            xfs.Add(new XElement(ns + "xf", new XAttribute("numFmtId", 9), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0), new XAttribute("applyNumberFormat", 1)));
            xfs.SetAttributeValue("count", xfs.Elements().Count()); Save("xl/styles.xml", styles);
        }
        if (formula is not null) cell.Add(new XElement(ns + "f", formula));
        if (numeric) { if (value.Length > 0) cell.Add(new XElement(ns + "v", value)); }
        else { cell.SetAttributeValue("t", "inlineStr"); cell.Add(new XElement(ns + "is", new XElement(ns + "t", value))); }
        Save(part, xml);
        return path;
    }
    Test("standard import uses the original first-barcode rule and J-column shortage", () =>
    {
        Assert(catalog.Products.Count == 104 && catalog.Products.Select(p => p.Barcode).Distinct().Count() == 104, "Original duplicate-barcode handling changed.");
        Assert(catalog.Resolve("0011010178").Barcode == "0011010178", "First duplicate was not retained.");
        Assert(product.Barcode == "0011010116" && product.FloatingRatioPercent == 110m, "Identifiers or ratio changed.");
        Assert(catalog.Resolve("4801000353").AllowedShortageGrams == 8.1m, "J-column value replaced by calculated shortage.");
        Reject(() => catalog.Resolve("UNKNOWN"), "Unknown barcode resolved.");
    });
    Test("sampling boundaries retain the original 21-31 to 32 behavior", () =>
    {
        (int Batch, int Count)[] values = [(1,1),(20,20),(21,32),(30,32),(40,32),(41,35),(60,35),(80,47),(100,49),(200,64),(300,67),(400,81),(500,81),(100000,98)];
        foreach (var (batch,count) in values) Assert(catalog.RequiredSamples(batch) == count, "Sampling boundary changed.");
        Reject(() => catalog.RequiredSamples(100001), "Unsupported batch accepted.");
    });
    Test("standard batch-size button mapping is restored", () =>
    {
        Assert(LegacyInspection.StandardBatchSize(catalog, 3) == 3 && LegacyInspection.StandardBatchSize(catalog, 32) == 40
            && LegacyInspection.StandardBatchSize(catalog, 98) == 100000 && LegacyInspection.StandardBatchSize(catalog, 33) == 33, "Original button mapping changed.");
    });
    Test("floating ratios accept numeric points, Excel percentages, original text and saved formula results", () =>
    {
        string[] paths = [RatioWorkbook("ratio-text", "110%"), RatioWorkbook("ratio-fullwidth", "110％"),
            RatioWorkbook("ratio-number", "110", true), RatioWorkbook("ratio-number-text", "110"),
            RatioWorkbook("ratio-percentage", "1.1", true, true),
            RatioWorkbook("ratio-formula", "110", true, false, "RANDBETWEEN(105,115)"),
            RatioWorkbook("ratio-formula-percentage", "1.1", true, true, "RANDBETWEEN(105,115)/100")];
        foreach (var path in paths)
        {
            var imported = StandardImporter.Read(path).Resolve(product.Barcode);
            Assert(imported.FloatingRatioPercent == 110m && Generation.MaximumWeight(imported) == 99m, "Equivalent ratio format changed the maximum: " + path);
        }
    });
    Test("ratio formulas require a saved result and other input formulas remain rejected", () =>
    {
        Reject(() => StandardImporter.Read(RatioWorkbook("ratio-no-cache", "", true, false, "RANDBETWEEN(105,115)")), "Formula without saved calculation accepted.");
        Reject(() => StandardImporter.Read(RatioWorkbook("ratio-error", "#VALUE!", false, false, "1/0")), "Formula error accepted.");
        Assert(StandardImporter.Read(RatioWorkbook("ratio-too-low", "99", true)).Resolve(product.Barcode).FloatingRatioPercent is null, "Below-100 ratio accepted.");
        var rows = new SheetRow(2, new() { [1] = new("90", true, true) });
        Reject(() => rows.At(1), "Formula restrictions changed outside the ratio importer.");
    });
    Test("80 percent of each product falls in outer quarters with balanced sides and shuffled order", () =>
    {
        foreach (var seed in new[] { 1, 7, 29, 1234 })
        foreach (var count in new[] { 1, 2, 3, 7, 20, 32, 98, 1000 })
        {
            var random = new Random(seed);
            var weights = new PassWeightGenerator(random.NextDouble).Generate(90m, 99m, count);
            var low = weights.Count(w => w <= 92.25m); var high = weights.Count(w => w >= 96.75m);
            Assert(weights.Count == count && weights.All(w => w >= 90m && w <= 99m && w * 10m == decimal.Truncate(w * 10m)), "Weight count, precision or bounds changed.");
            Assert(low + high == (int)decimal.Round(count * 0.8m, 0, MidpointRounding.AwayFromZero) && Math.Abs(low - high) <= 1, "80/20 quota or side balance failed.");
            if (count == 1000) Assert(!weights.SequenceEqual(weights.OrderBy(w => w)), "Generated rows were not shuffled.");
        }
    });
    Test("one-decimal band boundaries and narrow ranges stay inside requested limits", () =>
    {
        foreach (var next in new Func<double>[] { () => 0, () => double.BitDecrement(1d) })
        {
            var weights = new PassWeightGenerator(next).Generate(90.04m, 99.04m, 100);
            Assert(weights.All(w => w >= 90.04m && w <= 99.04m) && weights.Count(w => w <= 92.29m || w >= 96.79m) == 80, "Rounding moved a sample outside its range or band.");
        }
        Assert(new PassWeightGenerator().Generate(90m, 90m, 3).All(w => w == 90m), "Equal limits failed.");
        Assert(new PassWeightGenerator().Generate(90.04m, 90.16m, 3).All(w => w == 90.1m), "Narrow range failed.");
        Reject(() => new PassWeightGenerator().Generate(90.01m, 90.09m, 3), "Unrepresentable one-decimal range accepted.");
        Reject(() => new PassWeightGenerator(() => double.NaN).Generate(90m, 99m, 3), "Invalid random source accepted.");
    });
    Test("helper A/B import skips unknown rows and leaves invalid quantities for editing", () =>
    {
        var path = Temp("商品表.xlsx");
        XlsxWriter.Write(path, [new("商品", [["条码", "数量"], [product.Barcode, 3], ["UNKNOWN", 99], [product.Barcode, ""], [product.Barcode, "not-a-number"]])]);
        var rows = ProductionImporter.Read(path, catalog);
        Assert(rows.Count == 3 && rows.Select(r => r.BatchSize).SequenceEqual(new[] { 3, 0, 0 }), "Legacy helper import behavior changed.");
    });
    Test("helper maximum follows floating ratio or manual override", () =>
    {
        Assert(Generation.MaximumWeight(product) == 99m && Generation.MaximumWeight(product, 92m) == 92m, "Maximum rule changed.");
        Reject(() => Generation.MaximumWeight(product, 89m), "Below-standard maximum accepted.");
        Reject(() => Generation.MaximumWeight(product with { FloatingRatioPercent = null }), "Missing ratio defaulted.");
    });
    Test("helper generation accepts original small batches and retains date/time behavior", () =>
    {
        var entry = new ProductionEntry(product.Barcode, 30, "", "", "TEST-ONLY", "", 0, MaximumWeightGrams: 92m);
        var generated = Generation.Build(catalog, [entry], new(2026,11,9,23,59,59), new PassWeightGenerator(() => 0.5));
        Assert(generated.Count == 32 && generated.All(m => m.Grams >= 90m && m.Grams <= 92m), "Legacy generation count or cap changed.");
        Assert(generated[1].Date == "2026.11.09" && generated[1].Time == "00:00:00", "Original fixed-date sequence changed.");
        Assert(Generation.Build(catalog, [entry with { BatchSize = 1 }, entry with { BatchSize = 1 }], DateTime.Today).Count == 2, "Repeated helper rows were blocked.");
    });
    Test("helper exports only one workbook with the original eight columns", () =>
    {
        var directory = Path.Combine(temporary, "helper");
        var exported = Generation.Export(Path.Combine(directory, "W261109.xlsx"), catalog,
            [new(product.Barcode, 3, "", "", "TEST-ONLY", "", 0)], new(2026,11,9,12,0,0));
        Assert(Directory.GetFiles(directory, "*.xlsx").Length == 1 && exported.Count == 3, "Helper exported additional reports.");
        var sheet = XlsxReader.Read(exported.WeighingPath).Single();
        Assert(Enumerable.Range(1,8).Select(i => sheet.Rows[0].At(i)).SequenceEqual(new[] { "序号","日期","时间","皮重","重量","毛重","单位","扫描" }), "Eight-column header changed.");
        Assert(sheet.Rows.All(r => r.Cells.Keys.All(c => c <= 8)) && sheet.Name.StartsWith("W"), "Added visible metadata columns or source labels.");
        Assert(MeasurementImporter.Read(exported.WeighingPath).Count == 3, "Generated workbook cannot be imported by main.");
        Assert(XNode.DeepEquals(Xml(Content("WeighingTemplate.xlsx"),"xl/worksheets/sheet1.xml").Element(ns+"cols"), Xml(exported.WeighingPath,"xl/worksheets/sheet1.xml").Element(ns+"cols")), "Original weighing widths changed.");
    });
    Test("weighing round trip preserves eight-column values without added metadata", () =>
    {
        var path = Temp("称重.xlsx"); WorkbookExports.Measurements(path, catalog, measurements, Batches(3));
        var read = MeasurementImporter.Read(path);
        Assert(read.Select(m => m.Grams).SequenceEqual(measurements.Select(m => m.Grams)) && read.All(m => m.Barcode == product.Barcode), "Weights or barcode changed.");
        using var zip = ZipFile.OpenRead(path);
        Assert(!zip.Entries.Any(e => e.FullName.Contains("customXml")), "Unrequested hidden data part added.");
    });
    Test("original supplementation retains existing samples and uses six-draw centered weights", () =>
    {
        var adjusted = LegacyInspection.Prepare(catalog, measurements, Batches(40), () => 0.5).Single();
        Assert(adjusted.Measurements.Count == 32 && adjusted.Measurements.Take(3).SequenceEqual(measurements), "Supplement changed original records.");
        Assert(adjusted.Measurements.Skip(3).All(m => m.Grams == 90.5m), "Legacy supplementation range changed.");
        Assert(measurements.Count == 3, "Input list mutated.");
    });
    Test("original overflow trimming removes Pass then T1 then T2 from the end", () =>
    {
        var raw = Records(91m,89m,82m,92m,93m);
        var adjusted = LegacyInspection.Prepare(catalog, raw, Batches(3)).Single();
        Assert(adjusted.Measurements.Select(m => m.Grams).SequenceEqual(new[] { 91m,89m,82m }) && raw.Count == 5, "Legacy trimming order changed.");
    });
    Test("legacy count conclusion does not add an average-weight gate", () =>
    {
        var custom = catalog with { Products = [product with { StandardGrams = 100m, AllowedShortageGrams = 5m }] };
        var raw = Records(Enumerable.Repeat(100m,31).Append(99m).ToArray());
        var metrics = RecoveredJudgement.Evaluate(LegacyInspection.Prepare(custom, raw, Batches(40)).Single(), custom.SamplingRuleFor(40));
        Assert(metrics.LegacyCountPass == true && !metrics.AverageMeetsNominal, "Count-only legacy conclusion changed.");
        var report = WorkbookExports.Report(Temp("旧判定.xlsx"), custom, raw, Batches(40), new(2026,11,9), reportTemplate);
        Assert(XlsxReader.Read(report.SummaryPath).Single().Rows.Single(r => r.Number == 7).At(15) == "符合规定", "Added judgement text or mean gate in report.");
    });
    Test("reports have only original template sheets and original report wording", () =>
    {
        var generated = measurements.Select(m => m with { DataOrigin = "规则生成" }).ToList();
        var result = WorkbookExports.Report(Temp("输出.xlsx"), catalog, generated, Batches(3), new(2026,11,9), reportTemplate);
        var summary = XlsxReader.Read(result.SummaryPath).Single(); var detail = XlsxReader.Read(result.DetailPaths.Single()).Single();
        Assert(summary.Rows.Single(r => r.Number == 2).At(1) == "产品检测报告", "Added summary title marker.");
        Assert(detail.Rows.Single(r => r.Number == 1).At(1) == XlsxReader.Read(detailTemplate).Single().Rows.Single(r => r.Number == 1).At(1), "Detail title differs from original template.");
        var row = summary.Rows.Single(r => r.Number == 7);
        Assert(row.At(9) == "正常" && row.At(15) == "不合格", "Original default sensory or conclusion wording changed.");
        Assert(row.At(14) == product.InspectionBasis && summary.Rows.Single(r => r.Number == 4).At(14) == "检验依据", "Original template basis column changed.");
        var values = summary.Rows.Concat(detail.Rows).SelectMany(r => r.Cells.Values).Select(c => c.Text).ToArray();
        Assert(values.All(v => !new[] { "规则生成", "待复核", "未录入", "未提供", "旧版口径", "测量追溯", "报告说明" }.Any(v.Contains)), "Unrequested report marker remains.");
        Assert(detail.Rows.Single(r => r.Number == 35).At(3) == product.Code + "20261109", "Added record-number suffix.");
        Assert(detail.Rows.Single(r => r.Number == 33).At(3) == "2026年11月09日", "Original production-date default changed.");
    });
    Test("report geometry, merges and page setup come from the original workbooks", () =>
    {
        var result = WorkbookExports.Report(Temp("版式.xlsx"), catalog, measurements, Batches(3), new(2026,11,9), reportTemplate);
        foreach (var (source, target) in new[] { (reportTemplate,result.SummaryPath),(detailTemplate,result.DetailPaths.Single()) })
        {
            var before = Xml(source,"xl/worksheets/sheet1.xml"); var after = Xml(target,"xl/worksheets/sheet1.xml");
            foreach (var part in new[] { "cols","mergeCells","pageMargins","pageSetup" })
                Assert(XNode.DeepEquals(before.Element(ns+part),after.Element(ns+part)), "Original layout changed: " + part);
            Assert(before.Element(ns+"sheetData")!.Elements(ns+"row").Select(r=>string.Join("|",r.Attributes())).SequenceEqual(
                after.Element(ns+"sheetData")!.Elements(ns+"row").Select(r=>string.Join("|",r.Attributes()))), "Original row geometry changed.");
            using var zip = ZipFile.OpenRead(target);
            Assert(!zip.Entries.Any(e => e.FullName.Contains("externalLink")), "Historical broken external link retained.");
        }
    });
    Test("separate product files and summary retain original filename suffixes", () =>
    {
        var second = catalog.Products.First(p => p.Barcode != product.Barcode);
        var raw = measurements.Concat(new[] { measurements[0] with { Barcode = second.Barcode, RecordId = "SECOND" } }).ToList();
        var batches = Batches(3); batches.Add(second.Barcode,"1");
        var report = WorkbookExports.Report(Temp("分文件.xlsx"), catalog, raw, batches, new(2026,11,9), reportTemplate);
        Assert(report.DetailPaths.Count == 2 && Path.GetFileName(report.SummaryPath) == "分文件_批次检测报告结果.xlsx", "Separate filenames changed.");
        foreach (var path in report.DetailPaths) Assert(path.EndsWith("_批次检测报告.xlsx") && XlsxReader.Read(path).Count == 1, "Product report is not independent.");
        Assert(!File.Exists(Temp("分文件.xlsx")), "Combined workbook still exported.");
    });
    Test("natural product-code sorting and 15-row summary pagination match legacy behavior", () =>
    {
        Assert(new[] { "M10","M2","M1","Z2" }.OrderBy(c=>c,ProductCodeComparer.Instance).SequenceEqual(new[] { "M1","M2","M10","Z2" }), "Natural code sorting changed.");
        var products = catalog.Products.Take(16).ToList();
        var raw = products.Select((p,i)=>measurements[0] with { Barcode=p.Barcode, RecordId="PAGE-"+i }).ToList();
        var batches = products.ToDictionary(p=>p.Barcode,p=>"1");
        var report = WorkbookExports.Report(Temp("分页.xlsx"),catalog,raw,batches,new(2026,11,9),reportTemplate);
        var sheets = XlsxReader.Read(report.SummaryPath);
        Assert(sheets.Count==2 && report.DetailPaths.Count==16, "Unrequested pages or missing summary pagination.");
        var sorted=products.OrderBy(p=>p.Code,ProductCodeComparer.Instance).ToArray();
        Assert(sheets[0].Rows.Single(r=>r.Number==21).At(2)==sorted[14].Code && sheets[1].Rows.Single(r=>r.Number==7).At(2)==sorted[15].Code, "Pagination dropped or reordered products.");
    });
    Test("existing outputs are replaced as in the original program", () =>
    {
        var result=WorkbookExports.Report(Temp("覆盖.xlsx"),catalog,measurements,Batches(3),new(2026,11,9),reportTemplate);
        File.WriteAllText(result.SummaryPath,"OLD-CUSTOMER-FILE");
        WorkbookExports.Report(Temp("覆盖.xlsx"),catalog,Records(91m,91m,91m),Batches(3),new(2026,11,9),reportTemplate);
        Assert(XlsxReader.Read(result.SummaryPath).Single().Rows.Single(r=>r.Number==7).At(15)=="符合规定", "Legacy replacement failed.");
        var hash=DataFiles.Sha256(result.SummaryPath);
        Reject(()=>WorkbookExports.Report(Temp("覆盖.xlsx"),catalog,measurements,Batches(3),new(2026,11,9),Temp("missing.xlsx")),"Missing template accepted.");
        Assert(DataFiles.Sha256(result.SummaryPath)==hash && !Directory.GetDirectories(temporary,".mro-*").Any(),"Failed export damaged prior outputs.");
    });
    Test("input and source templates stay unchanged", () =>
    {
        Assert(DataFiles.Sha256(standard)==originalHash,"Original standard file changed.");
        Reject(()=>InspectionValidator.Validate(catalog,measurements,new Dictionary<string,string>{{product.Barcode,"0"}}),"Zero quantity accepted.");
        Reject(()=>InspectionValidator.Validate(catalog,[measurements[0],measurements[0]],Batches(2)),"Duplicate record ID accepted.");
    });
    Console.WriteLine($"All {passed} checks passed.");
    if(args.Contains("--keep-fixtures")) Console.WriteLine("TEST-ONLY fixtures: "+temporary);
    else Directory.Delete(temporary,true);
    return 0;
}
catch(Exception ex)
{
    Console.Error.WriteLine($"FAILED after {passed} checks. {ex}\nFixtures: {temporary}");return 1;
}
