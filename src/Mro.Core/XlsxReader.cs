using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Mro.Core;

public sealed record SheetCell(string Text, bool IsFormula, bool IsNumeric);
public sealed record SheetRow(int Number, Dictionary<int, SheetCell> Cells)
{
    public string At(int column)
    {
        if (!Cells.TryGetValue(column, out var value)) return "";
        if (value.IsFormula) throw new InvalidDataException($"第 {Number} 行第 {column} 列含公式，请提供已确认的原始值。");
        return value.Text.Trim();
    }
}
public sealed record WorkbookSheet(string Name, List<SheetRow> Rows, bool Date1904 = false);

public static class XlsxReader
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static XDocument Xml(ZipArchiveEntry entry)
    {
        if (entry.Length > 64L * 1024 * 1024) throw new InvalidDataException("工作表过大，请拆分后导入。");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024 * 1024 });
        return XDocument.Load(reader);
    }

    public static List<WorkbookSheet> Read(string path)
    {
        if (!Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("恢复版导入支持 .xlsx；旧 .xls 请先在 Excel 中另存为 .xlsx。");
        using var archive = ZipFile.OpenRead(path);
        ZipArchiveEntry Entry(string name) => archive.GetEntry(name) ?? throw new InvalidDataException($"Excel 缺少 {name}");
        var workbook = Xml(Entry("xl/workbook.xml"));
        var date1904 = ((string?)workbook.Root?.Element(S + "workbookPr")?.Attribute("date1904")) is "1" or "true";
        var relations = Xml(Entry("xl/_rels/workbook.xml.rels")).Root!.Elements()
            .Where(e => (string?)e.Attribute("TargetMode") != "External")
            .ToDictionary(e => (string)e.Attribute("Id")!, e => (string)e.Attribute("Target")!);
        var stringsEntry = archive.GetEntry("xl/sharedStrings.xml");
        string[] strings = stringsEntry is null ? [] : Xml(stringsEntry).Root!.Elements(S + "si")
            .Select(e => string.Concat(e.Descendants(S + "t").Select(t => t.Value))).ToArray();
        var result = new List<WorkbookSheet>();
        foreach (var sheet in workbook.Descendants(S + "sheet"))
        {
            var id = (string?)sheet.Attribute(R + "id") ?? throw new InvalidDataException("工作表没有关系 ID。");
            if (!relations.TryGetValue(id, out var target)) throw new InvalidDataException("工作表引用无效。");
            var part = new Uri(new Uri("https://xlsx.invalid/xl/workbook.xml"), target.Replace('\\', '/')).AbsolutePath.TrimStart('/');
            if (!part.StartsWith("xl/", StringComparison.Ordinal)) throw new InvalidDataException("工作表路径无效。");
            var rows = new List<SheetRow>();
            foreach (var row in Xml(Entry(part)).Descendants(S + "sheetData").Elements(S + "row"))
            {
                var cells = new Dictionary<int, SheetCell>();
                foreach (var cell in row.Elements(S + "c"))
                {
                    var reference = (string?)cell.Attribute("r") ?? throw new InvalidDataException("单元格缺少位置。");
                    var column = 0;
                    foreach (var ch in reference.TakeWhile(char.IsAsciiLetter)) column = checked(column * 26 + char.ToUpperInvariant(ch) - 'A' + 1);
                    if (column < 1 || column > 16384) throw new InvalidDataException("Excel 列超出范围。");
                    var kind = (string?)cell.Attribute("t");
                    var value = cell.Element(S + "v")?.Value ?? "";
                    string text;
                    if (kind == "s")
                    {
                        if (!int.TryParse(value, out var index) || index < 0 || index >= strings.Length) throw new InvalidDataException("共享文本引用无效。");
                        text = strings[index];
                    }
                    else if (kind == "inlineStr") text = string.Concat(cell.Descendants(S + "t").Select(t => t.Value));
                    else text = value;
                    cells[column] = new(text, cell.Element(S + "f") is not null, kind is null or "n");
                }
                rows.Add(new((int?)row.Attribute("r") ?? rows.Count + 1, cells));
            }
            result.Add(new((string?)sheet.Attribute("name") ?? "Sheet", rows, date1904));
        }
        if (result.Count == 0) throw new InvalidDataException("Excel 没有工作表。");
        return result;
    }
}
