using System.Globalization;

namespace Mro.Core;

public sealed record ProductionEntry(string Barcode, int BatchSize, string ProductName, string ProductionDate,
    string SourceFile, string SourceSheet, int SourceRow, string BatchId = "", decimal? MaximumWeightGrams = null, string SensoryResult = "")
{
    public string Key => InspectionKey.For(Barcode, BatchId);

    public static List<ProductionEntry> FromMeasurements(IEnumerable<Measurement> measurements) => measurements
        .GroupBy(m => m.Key).Where(g => g.Any(m => m.BatchSize.HasValue))
        .Select(g =>
        {
            var first = g.First(m => m.BatchSize.HasValue);
            if (g.Where(m => m.BatchSize.HasValue).Select(m => m.BatchSize).Distinct().Count() != 1)
                throw new InvalidDataException($"条码 {first.Barcode} 同一批次的生产数量不一致。");
            return new ProductionEntry(first.Barcode, first.BatchSize!.Value, "", first.ProductionDate,
                first.SourceFile, first.SourceSheet, first.SourceRow, first.BatchId, first.MaximumWeightGrams, first.SensoryResult);
        }).ToList();
}

public static class ProductionImporter
{
    public static List<ProductionEntry> Read(string path, StandardCatalog catalog)
    {
        var sheet = XlsxReader.Read(path)[0];
        var barcodes = catalog.Products.Select(p => p.Barcode).ToHashSet(StringComparer.Ordinal);
        var entries = new List<ProductionEntry>();
        foreach (var row in sheet.Rows)
        {
            var barcode = row.At(1).Trim();
            if (!barcodes.Contains(barcode)) continue;
            var quantity = int.TryParse(row.At(2).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) && n > 0 ? n : 0;
            entries.Add(new(barcode, quantity, catalog.Resolve(barcode).ProductName, "", Path.GetFullPath(path), sheet.Name, row.Number));
        }
        return entries;
    }
}
