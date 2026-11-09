using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace Mro.Core;

/// <summary>Fill the original report forms without adding report text, columns or worksheets.</summary>
public static class TemplateReportExporter
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/package/2006/content-types";
    private sealed record Page(string Name, XElement Sheet, string? PrintArea = null);
    private sealed record Template(XElement Sheet, XElement Styles, XElement Theme);

    private static Template Load(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        XElement Read(string name) { using var stream = (zip.GetEntry(name) ?? throw new InvalidDataException("模板缺少 " + name)).Open(); return XElement.Load(stream); }
        var sheet = Read("xl/worksheets/sheet1.xml");
        var strings = zip.GetEntry("xl/sharedStrings.xml") is null ? [] : Read("xl/sharedStrings.xml").Elements(S + "si").ToArray();
        foreach (var cell in sheet.Descendants(S + "c").Where(c => (string?)c.Attribute("t") == "s"))
        {
            var index = int.Parse(cell.Element(S + "v")!.Value, CultureInfo.InvariantCulture);
            cell.RemoveNodes(); cell.SetAttributeValue("t", "inlineStr");
            cell.Add(new XElement(S + "is", strings[index].Nodes().Select(n => n is XElement e ? new XElement(e) : n)));
        }
        return new(sheet, Read("xl/styles.xml"), Read("xl/theme/theme1.xml"));
    }

    private static string ReferenceFolder(string? templatePath)
    {
        if (templatePath is null) return Path.Combine(AppContext.BaseDirectory, "SharedContent", "ReferenceReports");
        if (!File.Exists(templatePath)) throw new FileNotFoundException("报告模板不存在。", templatePath);
        var directory = Path.GetDirectoryName(Path.GetFullPath(templatePath))!;
        return Path.GetFileName(templatePath) == "MRO-ReportTemplate.xlsx" ? Path.Combine(directory, "..", "ReferenceReports") : directory;
    }

    public static ReportExportResult Write(string path, string? templatePath, StandardCatalog catalog,
        IReadOnlyList<Measurement> measurements, IReadOnlyDictionary<string, string> batchSizes,
        DateTime reportDate, IReadOnlyList<ProductionEntry>? production = null)
    {
        var groups = LegacyInspection.Prepare(catalog, measurements, batchSizes);
        var folder = ReferenceFolder(templatePath);
        var summary = Load(Path.Combine(folder, "检测报告2.xlsx"));
        var detail = Load(Path.Combine(folder, "检测报告1.xlsx"));
        var result = ReportFileNames.Create(path, groups);
        var summaryPages = new List<Page>();
        var outputs = new List<PendingWorkbook>();
        var redStyles = new Dictionary<string, string>();
        var metrics = groups.Select(g => RecoveredJudgement.Evaluate(g, catalog.SamplingRuleFor(g.BatchSize))).ToArray();
        var date = reportDate.ToString("yyyy年MM月dd日", CultureInfo.InvariantCulture);
        var stamp = DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture);
        for (var pageIndex = 0; pageIndex * 15 < groups.Count; pageIndex++)
        {
            var sheet = new XElement(summary.Sheet);
            foreach (var cell in sheet.Descendants(S + "c").Where(c => RowNumber(c) is >= 7 and <= 21)) Clear(cell);
            SetCell(sheet, "C3", date); SetCell(sheet, "G3", date); SetCell(sheet, "K3", date);
            var pageGroups = groups.Skip(pageIndex * 15).Take(15).ToArray();
            for (var j = 0; j < pageGroups.Length; j++)
            {
                var group = pageGroups[j]; var m = metrics[pageIndex * 15 + j]; var row = 7 + j;
                object?[] values = [pageIndex * 15 + j + 1, group.Product.Code, group.Product.ProductName,
                    group.Product.Category, group.Product.StandardGrams, group.BatchSize, group.RequiredSamples,
                    group.Product.SensoryRequirement, "正常", group.Product.Requirement,
                    $"T1：{m.AllowedT1Count}，T2：{m.AllowedT2Count}",
                    LegacyInspection.Average(group).ToString("0.##", CultureInfo.InvariantCulture) + "g",
                    $"T1：{m.T1Count}，T2：{m.T2Count}", group.Product.InspectionBasis,
                    m.LegacyCountPass == true ? "符合规定" : "不合格"];
                for (var column = 0; column < values.Length; column++) SetCell(sheet, XlsxWriter.ColumnName(column + 1) + row, values[column]);
            }
            summaryPages.Add(new(SheetName("批次检测报告结果" + (groups.Count > 15 ? (pageIndex + 1).ToString() : "") + "_" + stamp), sheet, "$A$1:$O$23"));
        }
        for (var i = 0; i < groups.Count; i++)
        {
            var group = groups[i]; var m = metrics[i];
            var sheet = new XElement(detail.Sheet);
            for (var j = 1; j <= 100; j++) SetCell(sheet, DetailSlot(j), "");
            SetCell(sheet, "C2", group.Product.ProductName);
            SetCell(sheet, "G2", group.Product.StandardGrams.ToString("0.0", CultureInfo.InvariantCulture) + " g");
            for (var j = 0; j < group.Measurements.Count; j++)
            {
                var measurement = group.Measurements[j];
                var cell = SetCell(sheet, DetailSlot(j + 1), $"样品{j + 1}. {measurement.Grams.ToString("0.0", CultureInfo.InvariantCulture)} g");
                if (RecoveredJudgement.Classify(group.Product.StandardGrams, group.Product.AllowedShortageGrams!.Value, measurement.Grams) != LegacySampleJudgement.Pass)
                    cell.SetAttributeValue("s", RedStyle((string?)cell.Attribute("s") ?? "0", detail.Styles, redStyles));
            }
            SetCell(sheet, "C28", LegacyInspection.Average(group).ToString("0.0", CultureInfo.InvariantCulture) + " g");
            SetCell(sheet, "G28", group.Product.AllowedShortageGrams!.Value.ToString("0.0", CultureInfo.InvariantCulture) + " g");
            SetCell(sheet, "C29", m.T1Count); SetCell(sheet, "G29", m.T2Count);
            SetCell(sheet, "C31", group.Product.Barcode); SetCell(sheet, "G31", group.Product.SourceLocation);
            SetCell(sheet, "C32", date); SetCell(sheet, "G32", group.BatchSize + " 袋");
            SetCell(sheet, "C33", date); SetCell(sheet, "G33", group.Product.Category);
            SetCell(sheet, "C34", group.Product.Code); SetCell(sheet, "G34", group.RequiredSamples + " 袋");
            SetCell(sheet, "C35", group.Product.Code + reportDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
            SetCell(sheet, "G35", group.Product.InspectionBasis);
            var page = new Page(SheetName(group.Product.ProductName + " 批次检测报告_" + stamp), sheet, "$A$1:$H$35");
            outputs.Add(new(result.DetailPaths[i], target => WritePackage(target, [page], detail.Styles, detail.Theme)));
        }
        outputs.Add(new(result.SummaryPath, target => WritePackage(target, summaryPages, summary.Styles, summary.Theme)));
        WorkbookExportSet.Write(outputs.GroupBy(o => o.Path, StringComparer.OrdinalIgnoreCase).Select(g => g.Last()).ToList());
        return result;
    }

    internal static void WriteWeighing(string path, IReadOnlyList<Measurement> measurements)
    {
        if (measurements.Count == 0) throw new InvalidDataException("没有称重记录。");
        var template = Load(Path.Combine(AppContext.BaseDirectory, "SharedContent", "WeighingTemplate.xlsx"));
        var sheet = template.Sheet;
        var data = sheet.Element(S + "sheetData")!;
        var prototype = new XElement(data.Elements(S + "row").First(r => (int?)r.Attribute("r") == 2));
        data.Elements(S + "row").Where(r => (int?)r.Attribute("r") >= 2).Remove();
        for (var i = 0; i < Math.Max(6, measurements.Count); i++)
        {
            var row = new XElement(prototype); row.SetAttributeValue("r", i + 2);
            foreach (var cell in row.Elements(S + "c"))
            {
                cell.SetAttributeValue("r", new string(((string)cell.Attribute("r")!).TakeWhile(char.IsAsciiLetter).ToArray()) + (i + 2)); Clear(cell);
            }
            data.Add(row);
            if (i >= measurements.Count) continue;
            var m = measurements[i];
            object?[] values = [(i + 1).ToString("D5"), m.Date, m.Time, m.TareGrams ?? 0m,
                m.Grams.ToString("0.0", CultureInfo.InvariantCulture), (m.GrossGrams ?? m.Grams).ToString("0.0", CultureInfo.InvariantCulture), "g", m.Barcode];
            for (var c = 0; c < values.Length; c++) SetCell(sheet, XlsxWriter.ColumnName(c + 1) + (i + 2), values[c]);
        }
        sheet.Element(S + "dimension")?.SetAttributeValue("ref", "A1:H" + Math.Max(7, measurements.Count + 1));
        WorkbookExportSet.Write([new(path, target => WritePackage(target,
            [new("W" + DateTime.Now.ToString("yyMMdd", CultureInfo.InvariantCulture), sheet)], template.Styles, template.Theme))]);
    }

    private static void Clear(XElement cell) { cell.RemoveNodes(); cell.Attribute("t")?.Remove(); }
    private static int RowNumber(XElement cell) => int.Parse(new string(((string)cell.Attribute("r")!).Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
    private static string SheetName(string value) => new string(value.Select(c => "[]:*?/\\".Contains(c) ? '_' : c).ToArray())[..Math.Min(31, value.Length)];
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
        if (cell is null)
        {
            cell = new XElement(S + "c", new XAttribute("r", reference));
            if (row.Elements(S + "c").LastOrDefault()?.Attribute("s") is XAttribute style) cell.SetAttributeValue("s", style.Value);
            row.Add(cell);
            var sorted = row.Elements(S + "c").OrderBy(c => ColumnIndex((string)c.Attribute("r")!)).ToArray();
            row.RemoveNodes(); row.Add(sorted);
        }
        cell.RemoveNodes(); cell.Attribute("t")?.Remove();
        if (value is decimal or int or long) cell.Add(new XElement(S + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
        else { cell.SetAttributeValue("t", "inlineStr"); cell.Add(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""))); }
        return cell;
    }

    private static int ColumnIndex(string reference)
    {
        var value = 0;
        foreach (var ch in reference.TakeWhile(char.IsAsciiLetter)) value = value * 26 + ch - 'A' + 1;
        return value;
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
