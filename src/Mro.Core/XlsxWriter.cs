using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace Mro.Core;

public sealed record WorkbookPage(string Name, List<object?[]> Rows, List<string>? Merges = null);

public static class XlsxWriter
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageR = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Content = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static void Write(string path, IReadOnlyList<WorkbookPage> pages)
    {
        if (pages.Count == 0 || pages.Select(p => p.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != pages.Count)
            throw new InvalidDataException("工作表名称为空或重复。");
        foreach (var page in pages)
            if (string.IsNullOrWhiteSpace(page.Name) || page.Name.Length > 31 || page.Name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0
                || page.Rows.Count > 1_048_576 || page.Rows.Any(r => r.Length > 16384))
                throw new InvalidDataException("工作表名称或数据规模无效。");
        path = Path.GetFullPath(path);
        if (File.Exists(path)) throw new IOException("输出文件已存在，请选择新文件名；不会覆盖输入或原报告。");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var archive = ZipFile.Open(temporary, ZipArchiveMode.Create))
            {
                void Add(string name, XElement xml)
                {
                    using var stream = archive.CreateEntry(name).Open();
                    new XDocument(new XDeclaration("1.0", "utf-8", "yes"), xml).Save(stream);
                }
                var types = new XElement(Content + "Types",
                    new XElement(Content + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(Content + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                    new XElement(Content + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                    new XElement(Content + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
                for (var i = 0; i < pages.Count; i++) types.Add(new XElement(Content + "Override", new XAttribute("PartName", $"/xl/worksheets/sheet{i + 1}.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
                Add("[Content_Types].xml", types);
                Add("_rels/.rels", new XElement(PackageR + "Relationships", new XElement(PackageR + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", R.NamespaceName + "/officeDocument"), new XAttribute("Target", "xl/workbook.xml"))));
                Add("xl/workbook.xml", new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R),
                    new XElement(S + "sheets", pages.Select((page, i) => new XElement(S + "sheet", new XAttribute("name", page.Name), new XAttribute("sheetId", i + 1), new XAttribute(R + "id", "rId" + (i + 1)))))));
                var relationships = new XElement(PackageR + "Relationships", pages.Select((_, i) =>
                    new XElement(PackageR + "Relationship", new XAttribute("Id", "rId" + (i + 1)), new XAttribute("Type", R.NamespaceName + "/worksheet"), new XAttribute("Target", $"worksheets/sheet{i + 1}.xml"))));
                relationships.Add(new XElement(PackageR + "Relationship", new XAttribute("Id", "rIdStyles"), new XAttribute("Type", R.NamespaceName + "/styles"), new XAttribute("Target", "styles.xml")));
                Add("xl/_rels/workbook.xml.rels", relationships);
                Add("xl/styles.xml", Styles());
                for (var i = 0; i < pages.Count; i++) Add($"xl/worksheets/sheet{i + 1}.xml", Sheet(pages[i]));
            }
            File.Move(temporary, path, false);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private static XElement Styles()
    {
        XElement Font(bool bold) => new(S + "font", bold ? new XElement(S + "b") : null,
            new XElement(S + "sz", new XAttribute("val", 11)), new XElement(S + "name", new XAttribute("val", "Microsoft YaHei")));
        XElement Xf(int font, int fill, int numberFormat) => new(S + "xf", new XAttribute("numFmtId", numberFormat),
            new XAttribute("fontId", font), new XAttribute("fillId", fill), new XAttribute("borderId", 1),
            new XAttribute("xfId", 0), new XAttribute("applyAlignment", 1), new XAttribute("applyNumberFormat", 1),
            new XElement(S + "alignment", new XAttribute("vertical", "center"), new XAttribute("wrapText", 1)));
        return new XElement(S + "styleSheet",
            new XElement(S + "numFmts", new XAttribute("count", 1), new XElement(S + "numFmt", new XAttribute("numFmtId", 164), new XAttribute("formatCode", "0.###"))),
            new XElement(S + "fonts", new XAttribute("count", 2), Font(false), Font(true)),
            new XElement(S + "fills", new XAttribute("count", 3),
                new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "none"))),
                new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "gray125"))),
                new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "solid"), new XElement(S + "fgColor", new XAttribute("rgb", "FFF3F1EA")), new XElement(S + "bgColor", new XAttribute("indexed", 64))))),
            new XElement(S + "borders", new XAttribute("count", 2), new XElement(S + "border"),
                new XElement(S + "border", new[] { "left", "right", "top", "bottom" }.Select(side => new XElement(S + side, new XAttribute("style", "thin"), new XElement(S + "color", new XAttribute("rgb", "FFD5CCC2")))))),
            new XElement(S + "cellStyleXfs", new XAttribute("count", 1), new XElement(S + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))),
            new XElement(S + "cellXfs", new XAttribute("count", 3), Xf(0, 0, 0), Xf(1, 2, 0), Xf(0, 0, 164)),
            new XElement(S + "cellStyles", new XAttribute("count", 1), new XElement(S + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0))));
    }

    private static XElement Sheet(WorkbookPage page)
    {
        var columnCount = page.Rows.Count == 0 ? 1 : Math.Max(1, page.Rows.Max(r => r.Length));
        var data = new XElement(S + "sheetData");
        for (var r = 0; r < page.Rows.Count; r++)
        {
            var row = new XElement(S + "row", new XAttribute("r", r + 1), new XAttribute("ht", r < 2 ? 30 : 24), new XAttribute("customHeight", 1));
            for (var c = 0; c < page.Rows[r].Length; c++)
            {
                var value = page.Rows[r][c];
                var numeric = value is decimal or double or float or int or long;
                var cell = new XElement(S + "c", new XAttribute("r", ColumnName(c + 1) + (r + 1)), new XAttribute("s", r < 2 ? 1 : numeric ? 2 : 0));
                if (numeric) cell.Add(new XElement(S + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
                else cell.Add(new XAttribute("t", "inlineStr"), new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), Convert.ToString(value, CultureInfo.InvariantCulture) ?? "")));
                row.Add(cell);
            }
            data.Add(row);
        }
        var sheet = new XElement(S + "worksheet",
            new XElement(S + "sheetViews", new XElement(S + "sheetView", new XAttribute("workbookViewId", 0), new XElement(S + "pane", new XAttribute("ySplit", 2), new XAttribute("topLeftCell", "A3"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")))),
            new XElement(S + "cols", Enumerable.Range(1, columnCount).Select(c => new XElement(S + "col", new XAttribute("min", c), new XAttribute("max", c), new XAttribute("width", c >= 7 ? 28 : 19), new XAttribute("customWidth", 1)))), data);
        if (page.Merges is { Count: > 0 }) sheet.Add(new XElement(S + "mergeCells", new XAttribute("count", page.Merges.Count), page.Merges.Select(m => new XElement(S + "mergeCell", new XAttribute("ref", m)))));
        sheet.Add(new XElement(S + "pageMargins", new XAttribute("left", 0.25), new XAttribute("right", 0.25), new XAttribute("top", 0.4), new XAttribute("bottom", 0.4), new XAttribute("header", 0.2), new XAttribute("footer", 0.2)),
            new XElement(S + "pageSetup", new XAttribute("paperSize", 9), new XAttribute("orientation", "landscape")));
        return sheet;
    }

    public static string ColumnName(int column)
    {
        var name = "";
        while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; }
        return name;
    }
}
