using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace Mro.Core;

/// <summary>Fill the empty recovered template; never copy historical result values.</summary>
public static class TemplateReportExporter
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/package/2006/content-types";
    private sealed record Page(string Name, XElement Sheet, string? PrintArea = null);

    public static void Write(string path, string templatePath, StandardCatalog catalog,
        IReadOnlyList<Measurement> measurements, IReadOnlyDictionary<string, string> batchSizes,
        DateTime reportDate, IReadOnlyList<ProductionEntry>? production = null)
    {
        var groups = InspectionValidator.Validate(catalog, measurements, batchSizes);
        if (production is not null && production.Any(p => !batchSizes.ContainsKey(p.Barcode)))
            throw new InvalidDataException("生产表中还有未纳入报告的产品，不能省略缺测产品。");
        using var template = ZipFile.OpenRead(templatePath);
        XElement Xml(string entry)
        {
            using var stream = (template.GetEntry(entry) ?? throw new InvalidDataException("模板缺少 " + entry)).Open();
            return XElement.Load(stream);
        }
        var summary = Xml("xl/worksheets/sheet1.xml");
        var detail = Xml("xl/worksheets/sheet2.xml");
        var trace = Xml("xl/worksheets/sheet3.xml");
        var notes = Xml("xl/worksheets/sheet4.xml");
        if (new[] { summary, detail, trace, notes }.Any(s => s.Descendants(S + "f").Any()))
            throw new InvalidDataException("空白报告模板不能含计算或历史结果公式。");
        var styles = Xml("xl/styles.xml");
        var pages = new List<Page>();
        var redStyles = new Dictionary<string, string>();
        var metrics = groups.Select(g => RecoveredJudgement.Evaluate(g, catalog.SamplingRuleFor(g.BatchSize))).ToArray();
        var productionByBarcode = (production ?? []).ToDictionary(e => e.Barcode);
        for (var pageIndex = 0; pageIndex * 15 < groups.Count; pageIndex++)
        {
            var sheet = new XElement(summary);
            var pageGroups = groups.Skip(pageIndex * 15).Take(15).ToArray();
            var dates = pageGroups.Select(g => productionByBarcode.GetValueOrDefault(g.Product.Barcode)?.ProductionDate ?? "")
                .Where(d => d.Length > 0).Distinct().ToArray();
            var samplingDates = pageGroups.SelectMany(g => g.Measurements).Select(m => m.Date).Where(d => d.Length > 0).Distinct().ToArray();
            SetCell(sheet, "C3", samplingDates.Length == 1 ? samplingDates[0] : "见测量追溯");
            SetCell(sheet, "G3", dates.Length == 1 && pageGroups.All(g => !string.IsNullOrEmpty(productionByBarcode.GetValueOrDefault(g.Product.Barcode)?.ProductionDate)) ? dates[0] : "见生产表/未提供");
            SetCell(sheet, "K3", reportDate.ToString("yyyy年MM月dd日", CultureInfo.InvariantCulture));
            for (var j = 0; j < pageGroups.Length; j++)
            {
                var group = pageGroups[j]; var m = metrics[pageIndex * 15 + j]; var row = 7 + j;
                object?[] values = [pageIndex * 15 + j + 1, group.Product.Code, group.Product.ProductName,
                    group.Product.Category, group.Product.StandardGrams.ToString("0.0", CultureInfo.InvariantCulture),
                    group.BatchSize, CountText(group), group.Product.SensoryRequirement, "未录入",
                    group.Product.Requirement, $"T1：{m.AllowedT1Count}，T2：{m.AllowedT2Count}",
                    m.AverageGrams.ToString("0.##", CultureInfo.InvariantCulture) + "g",
                    m.LegacyCountPass is null ? "缺少短缺量" : $"T1：{m.T1Count}，T2：{m.T2Count}", m.ReviewConclusion];
                for (var column = 0; column < values.Length; column++) SetCell(sheet, XlsxWriter.ColumnName(column + 1) + row, values[column]);
            }
            pages.Add(new($"汇总{pageIndex + 1:000}", sheet, "$A$1:$N$23"));
        }
        var traceRows = new List<object?[]>();
        for (var i = 0; i < groups.Count; i++)
        {
            var group = groups[i]; var m = metrics[i];
            for (var offset = 0; offset < group.Measurements.Count; offset += 100)
            {
                var sheet = new XElement(detail);
                SetCell(sheet, "C2", group.Product.ProductName);
                SetCell(sheet, "G2", group.Product.StandardGrams.ToString("0.0", CultureInfo.InvariantCulture) + " g");
                for (var j = 0; j < Math.Min(100, group.Measurements.Count - offset); j++)
                {
                    var measurement = group.Measurements[offset + j];
                    var cell = SetCell(sheet, DetailSlot(j + 1), $"样品{offset + j + 1}. {measurement.Grams.ToString("0.0", CultureInfo.InvariantCulture)} g");
                    if (group.Product.AllowedShortageGrams is decimal shortage && RecoveredJudgement.Classify(group.Product.StandardGrams, shortage, measurement.Grams) != LegacySampleJudgement.Pass)
                        cell.SetAttributeValue("s", RedStyle((string?)cell.Attribute("s") ?? "0", styles, redStyles));
                }
                SetCell(sheet, "C28", m.AverageGrams.ToString("0.0", CultureInfo.InvariantCulture) + " g");
                SetCell(sheet, "G28", group.Product.AllowedShortageGrams is decimal allowed ? allowed.ToString("0.0", CultureInfo.InvariantCulture) + " g" : "未提供");
                SetCell(sheet, "C29", m.LegacyCountPass is null ? "待复核" : m.T1Count.ToString());
                SetCell(sheet, "G29", m.LegacyCountPass is null ? "待复核" : m.T2Count.ToString());
                SetCell(sheet, "C31", group.Product.Barcode); SetCell(sheet, "G31", group.Product.SourceLocation);
                SetCell(sheet, "C32", reportDate.ToString("yyyy年MM月dd日", CultureInfo.InvariantCulture));
                SetCell(sheet, "G32", group.BatchSize + " 袋");
                SetCell(sheet, "C33", productionByBarcode.GetValueOrDefault(group.Product.Barcode)?.ProductionDate is { Length: > 0 } date ? date : "未提供");
                SetCell(sheet, "G33", group.Product.Category); SetCell(sheet, "C34", group.Product.Code);
                SetCell(sheet, "G34", CountText(group) + " 袋");
                SetCell(sheet, "C35", group.Product.Code + reportDate.ToString("yyyyMMdd") + $"-{offset / 100 + 1}");
                SetCell(sheet, "G35", group.Product.InspectionBasis);
                pages.Add(new($"明细{i + 1:000}_{offset / 100 + 1:00}", sheet, "$A$1:$H$35"));
            }
            for (var j = 0; j < group.Measurements.Count; j++)
            {
                var measurement = group.Measurements[j];
                traceRows.Add([group.Product.Barcode, group.Product.Code, group.Product.ProductName, j + 1, measurement.Grams,
                    group.Product.AllowedShortageGrams is decimal shortage ? RecoveredJudgement.Classify(group.Product.StandardGrams, shortage, measurement.Grams).ToString() : "未判定",
                    measurement.SourceFile, measurement.SourceSheet, measurement.SourceRow, measurement.RecordId]);
            }
        }
        FillRows(trace, traceRows);
        pages.Add(new("测量追溯", trace));
        var noteRows = new List<object?[]> {
            ["报告范围", "按原程序口径计算的实测统计；法规未独立核验，感官结果未录入，不自动出具总体合格证明。"],
            ["标准来源", catalog.SourceFile], ["标准SHA-256", catalog.SourceSha256],
            ["模板SHA-256", DataFiles.Sha256(templatePath)],
            ["T1/T2口径", "旧版：实测≥规格为Pass；规格-短缺量≤实测<规格为T1；低于规格-短缺量为T2。不是已核实的法规定义。"],
            ["旧版批次口径", "原函数只检查T1数量及是否有T2；本版额外显示平均值和数量偏差，不能把计数通过当作完整法规合格。"],
            ["短缺量来源", "读取标准表J列；H列文字与J列可能冲突，需审核。"],
            ["测量范围", "所有导入记录均保留；不截取、补值或改变不利记录。"],
            ["显示精度", "表单按旧版显示1或2位小数；测量追溯页保留原始精度。"],
            ["分页", "汇总每页15个产品，明细每页100条；统计按该产品全部记录计算。"],
            ["来源字段", "仅用于回查输入，不代表经过仪器签名认证。"],
            ["生产数量来源", production is { Count: > 0 } ? "已导入生产表；下方逐条列出来源与是否被手工改动。" : "界面手工输入"]
        };
        for (var i = 0; i < groups.Count; i++)
        {
            var g = groups[i]; var m = metrics[i];
            noteRows.Add([$"{g.Product.Barcode} 精确平均值", m.AverageGrams]);
            noteRows.Add([$"{g.Product.Barcode} 平均值达到规格", m.AverageMeetsNominal ? "是" : "否"]);
            noteRows.Add([$"{g.Product.Barcode} 原版计数函数", m.LegacyCountPass is null ? "无法判定" : m.LegacyCountPass.Value ? "计数通过（不等于法规合格）" : "计数不通过"]);
            if (productionByBarcode.TryGetValue(g.Product.Barcode, out var entry))
                noteRows.Add([$"{g.Product.Barcode} 生产表来源", $"{entry.SourceFile} | {entry.SourceSheet}:{entry.SourceRow} | 原数量{entry.BatchSize}，当前数量{g.BatchSize}"]);
        }
        FillRows(notes, noteRows); pages.Add(new("报告说明", notes));
        WritePackage(path, pages, styles, Xml("xl/theme/theme1.xml"));
    }

    private static string CountText(InspectionGroup group) => group.Measurements.Count == group.RequiredSamples
        ? group.RequiredSamples.ToString() : $"实测{group.Measurements.Count}/应{group.RequiredSamples}";
    public static string DetailSlot(int oneBasedIndex)
    {
        if (oneBasedIndex is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(oneBasedIndex));
        var zero = oneBasedIndex - 1;
        return XlsxWriter.ColumnName(1 + zero / 25 * 2) + (3 + zero % 25);
    }
    private static XElement SetCell(XElement sheet, string reference, object? value)
    {
        var data = sheet.Element(S + "sheetData") ?? throw new InvalidDataException("模板缺少sheetData。");
        var number = int.Parse(new string(reference.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
        var row = data.Elements(S + "row").FirstOrDefault(r => (int?)r.Attribute("r") == number);
        if (row is null) { row = new XElement(S + "row", new XAttribute("r", number)); data.Add(row); }
        var cell = row.Elements(S + "c").FirstOrDefault(c => (string?)c.Attribute("r") == reference);
        if (cell is null) { cell = new XElement(S + "c", new XAttribute("r", reference)); row.Add(cell); }
        cell.RemoveNodes(); cell.Attribute("t")?.Remove();
        if (value is decimal or int or long) cell.Add(new XElement(S + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
        else { cell.SetAttributeValue("t", "inlineStr"); cell.Add(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""))); }
        return cell;
    }
    private static void FillRows(XElement sheet, List<object?[]> rows)
    {
        if (rows.Count > 1_048_574) throw new InvalidDataException("追溯记录超过单页Excel上限，请分批导出。");
        var data = sheet.Element(S + "sheetData")!;
        var prototype = new XElement(data.Elements(S + "row").First(r => (int?)r.Attribute("r") == 3));
        data.Elements(S + "row").Where(r => (int?)r.Attribute("r") >= 3).Remove();
        for (var i = 0; i < rows.Count; i++)
        {
            var row = new XElement(prototype); row.SetAttributeValue("r", i + 3);
            foreach (var cell in row.Elements(S + "c"))
            {
                var column = new string(((string)cell.Attribute("r")!).TakeWhile(char.IsAsciiLetter).ToArray());
                cell.SetAttributeValue("r", column + (i + 3)); cell.RemoveNodes(); cell.Attribute("t")?.Remove();
            }
            data.Add(row);
            for (var j = 0; j < rows[i].Length; j++) SetCell(sheet, XlsxWriter.ColumnName(j + 1) + (i + 3), rows[i][j]);
        }
        sheet.Element(S + "dimension")?.Remove();
    }
    private static string RedStyle(string original, XElement styles, Dictionary<string, string> cache)
    {
        if (cache.TryGetValue(original, out var cached)) return cached;
        var fonts = styles.Element(S + "fonts")!; var xfs = styles.Element(S + "cellXfs")!;
        var xf = new XElement(xfs.Elements().ElementAt(int.Parse(original)));
        var font = new XElement(fonts.Elements().ElementAt((int?)xf.Attribute("fontId") ?? 0));
        font.Element(S + "color")?.Remove(); font.Add(new XElement(S + "color", new XAttribute("rgb", "FFFF0000")));
        var fontId = fonts.Elements().Count(); fonts.Add(font); fonts.SetAttributeValue("count", fontId + 1);
        xf.SetAttributeValue("fontId", fontId); xf.SetAttributeValue("applyFont", 1);
        var id = xfs.Elements().Count().ToString(CultureInfo.InvariantCulture); xfs.Add(xf); xfs.SetAttributeValue("count", xfs.Elements().Count());
        cache[original] = id; return id;
    }

    private static void WritePackage(string path, List<Page> pages, XElement styles, XElement theme)
    {
        path = Path.GetFullPath(path);
        if (File.Exists(path)) throw new IOException("输出文件已存在，请使用新文件名。");
        if (pages.Any(p => p.Name.Length > 31) || pages.Select(p => p.Name).Distinct().Count() != pages.Count)
            throw new InvalidDataException("报告工作表名称重复或过长。");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var zip = ZipFile.Open(temporary, ZipArchiveMode.Create))
            {
                void Add(string name, XElement element) { using var stream = zip.CreateEntry(name).Open(); new XDocument(element).Save(stream); }
                XElement Relation(string id, string kind, string target) => new(P + "Relationship", new XAttribute("Id", id), new XAttribute("Type", R.NamespaceName + "/" + kind), new XAttribute("Target", target));
                XElement Override(string part, string kind) => new(C + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml." + kind + "+xml"));
                var types = new XElement(C + "Types", new XElement(C + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(C + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                    Override("xl/workbook.xml", "sheet.main"), Override("xl/styles.xml", "styles"),
                    new XElement(C + "Override", new XAttribute("PartName", "/xl/theme/theme1.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")));
                var workbook = new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R),
                    new XElement(S + "bookViews", new XElement(S + "workbookView")),
                    new XElement(S + "sheets", pages.Select((p, i) => new XElement(S + "sheet", new XAttribute("name", p.Name), new XAttribute("sheetId", i + 1), new XAttribute(R + "id", "rId" + (i + 1))))));
                var names = new XElement(S + "definedNames");
                var relations = new XElement(P + "Relationships", Relation("styles", "styles", "styles.xml"), Relation("theme", "theme", "theme/theme1.xml"));
                for (var i = 0; i < pages.Count; i++)
                {
                    var p = pages[i]; var entry = $"xl/worksheets/sheet{i + 1}.xml";
                    types.Add(Override(entry, "worksheet")); relations.Add(Relation("rId" + (i + 1), "worksheet", $"worksheets/sheet{i + 1}.xml"));
                    if (p.PrintArea is not null) names.Add(new XElement(S + "definedName", new XAttribute("name", "_xlnm.Print_Area"), new XAttribute("localSheetId", i), $"'{p.Name}'!{p.PrintArea}"));
                    if (p.Name.StartsWith("汇总", StringComparison.Ordinal)) names.Add(new XElement(S + "definedName", new XAttribute("name", "_xlnm.Print_Titles"), new XAttribute("localSheetId", i), $"'{p.Name}'!$1:$6"));
                    Add(entry, p.Sheet);
                }
                workbook.Add(names);
                Add("[Content_Types].xml", types); Add("_rels/.rels", new XElement(P + "Relationships", Relation("rId1", "officeDocument", "xl/workbook.xml")));
                Add("xl/workbook.xml", workbook); Add("xl/_rels/workbook.xml.rels", relations); Add("xl/styles.xml", styles); Add("xl/theme/theme1.xml", theme);
            }
            File.Move(temporary, path, false);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
