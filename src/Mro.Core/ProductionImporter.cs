using System.Globalization;

namespace Mro.Core;

public sealed record ProductionEntry(string Barcode, int BatchSize, string ProductName, string ProductionDate,
    string SourceFile, string SourceSheet, int SourceRow);

public static class ProductionImporter
{
    public static List<ProductionEntry> Read(string path, StandardCatalog catalog)
    {
        // Original HelperEntryWorkbookImporter reads first sheet, column A barcode / B quantity.
        // Also accept named barcode/code headers, resolving codes explicitly and uniquely.
        var sheet = XlsxReader.Read(path)[0];
        var header = sheet.Rows.FirstOrDefault(r => r.Cells.Values.Any(c =>
            c.Text.Trim() is "条码" or "产品条码" or "编码" or "产品编号"));
        int barcodeColumn = 1, batchColumn = 2, codeColumn = 0, nameColumn = 0, dateColumn = 0;
        if (header is not null)
        {
            int Find(params string[] names) => header.Cells.FirstOrDefault(c => names.Contains(c.Value.Text.Trim())).Key;
            barcodeColumn = Find("条码", "产品条码");
            codeColumn = Find("编码", "产品编号");
            batchColumn = Find("生产数量", "抽样基数", "批量", "数量");
            nameColumn = Find("名称", "产品名称", "样品名称");
            dateColumn = Find("生产日期", "生产日期/批号");
            if (batchColumn == 0) throw new InvalidDataException("生产表缺少生产数量/抽样基数列。");
        }
        var result = new List<ProductionEntry>();
        foreach (var row in sheet.Rows.Where(r => header is null || r.Number > header.Number))
        {
            if (row.Cells.Values.All(c => string.IsNullOrWhiteSpace(c.Text) && !c.IsFormula)) continue;
            ProductStandard product;
            var barcode = row.At(barcodeColumn);
            if (barcodeColumn > 0)
            {
                if (barcode.Length == 0) throw new InvalidDataException($"生产表第 {row.Number} 行缺少条码。");
                product = catalog.Resolve(barcode);
                var code = row.At(codeColumn);
                if (code.Length > 0 && code != product.Code) throw new InvalidDataException($"生产表第 {row.Number} 行条码与编码不一致。");
            }
            else
            {
                var code = row.At(codeColumn);
                var matches = catalog.Products.Where(p => p.Code == code && code.Length > 0).ToList();
                if (matches.Count != 1) throw new InvalidDataException($"生产表第 {row.Number} 行编码没有唯一标准产品：{code}");
                product = matches[0]; barcode = product.Barcode;
                catalog.Resolve(barcode); // A code must not bypass a duplicate-barcode conflict in the weighing path.
            }
            var name = row.At(nameColumn);
            if (name.Length > 0 && name != product.ProductName) throw new InvalidDataException($"生产表第 {row.Number} 行名称与标准表不一致。");
            var batch = DataFiles.PositiveInteger(row.At(batchColumn), $"生产表第 {row.Number} 行生产数量");
            var date = row.At(dateColumn);
            if (dateColumn > 0 && row.Cells.TryGetValue(dateColumn, out var dateCell) && dateCell.IsNumeric && date.Length > 0)
            {
                if (DateTime.TryParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var ymd)) date = ymd.ToString("yyyy-MM-dd");
                else date = DateTime.FromOADate(double.Parse(date, CultureInfo.InvariantCulture) + (sheet.Date1904 ? 1462 : 0)).ToString("yyyy-MM-dd");
            }
            result.Add(new(barcode, batch, product.ProductName, date, Path.GetFullPath(path), sheet.Name, row.Number));
        }
        if (result.Count == 0) throw new InvalidDataException("生产表没有有效条目。");
        if (result.GroupBy(r => r.Barcode).Any(g => g.Count() > 1))
            throw new InvalidDataException("生产表存在重复产品/批次；请按实际批次拆分，不能静默合并生产数量。");
        return result;
    }
}
