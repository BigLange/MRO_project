using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;

namespace Mro.Core;

// New implementation for measured data; these are not decompiled business models.
public sealed record ProductStandard(string Barcode, string Code, string ProductName, decimal StandardGrams,
    string Category, string Requirement, string InspectionBasis, string SourceLocation,
    decimal? AllowedShortageGrams = null, string SensoryRequirement = "", decimal? FloatingRatioPercent = null);

public sealed record SamplingRule(int MinimumBatch, int MaximumBatch, int SampleCount, bool SampleEntireBatch = false,
    int AllowedT1Count = 0, int AllowedT2Count = 0);

public sealed record StandardCatalog(List<ProductStandard> Products, List<SamplingRule> SamplingRules,
    string SourceFile, string SourceSha256, List<string> ImportWarnings)
{
    public ProductStandard Resolve(string barcode)
    {
        var matches = Products.Where(p => p.Barcode == barcode.Trim()).ToList();
        if (matches.Count == 0) throw new InvalidDataException($"未知条码：{barcode}。不会自动按名称或编码猜测匹配。");
        if (matches.Count != 1) throw new InvalidDataException($"条码 {barcode} 对应多个标准产品，请先解决标准表歧义。");
        return matches[0];
    }

    public int RequiredSamples(int batchSize)
    {
        var rule = SamplingRuleFor(batchSize);
        return rule.SampleEntireBatch ? batchSize : rule.SampleCount;
    }

    public SamplingRule SamplingRuleFor(int batchSize)
    {
        if (batchSize <= 0) throw new InvalidDataException("抽样基数必须是正整数。");
        var matches = SamplingRules.Where(r => r.MinimumBatch <= batchSize && r.MaximumBatch >= batchSize).ToList();
        if (matches.Count != 1)
            throw new InvalidDataException($"抽样基数 {batchSize} 没有唯一规则。原程序支持 1–100000，其他范围请核实 sampling-plan.json。");
        return matches[0];
    }
}

public sealed record Measurement(string Barcode, decimal Grams, string Date, string Time,
    decimal? TareGrams, decimal? GrossGrams, string SourceFile, string SourceSheet, int SourceRow, string RecordId,
    string DataOrigin = "导入记录", string BatchId = "", int? BatchSize = null, string ProductionDate = "", string SensoryResult = "", decimal? MaximumWeightGrams = null)
{
    public string Key => InspectionKey.For(Barcode, BatchId);
}

public static class InspectionKey
{
    public static string For(string barcode, string batchId = "")
    {
        barcode = barcode.Trim(); batchId = batchId.Trim();
        if (barcode.Contains('\u001f') || batchId.Contains('\u001f')) throw new InvalidDataException("条码或批号包含无效分隔符。");
        return batchId.Length == 0 ? barcode : barcode + '\u001f' + batchId;
    }
}

public sealed record InspectionGroup(ProductStandard Product, int BatchSize, int RequiredSamples, List<Measurement> Measurements, string BatchId = "")
{
    public decimal AverageGrams => Measurements.Average(m => m.Grams);
    public string Key => InspectionKey.For(Product.Barcode, BatchId);
}

public static class DataFiles
{
    public static readonly JsonSerializerOptions Json = new() { WriteIndented = true, PropertyNameCaseInsensitive = true };
    public static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
    public static string Identifier(string value) => value.Trim(); // Never parse barcodes as numbers.
    public static decimal Decimal(string value, string description)
    {
        if (!decimal.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
            throw new InvalidDataException($"{description} 不是有效数值：{value}");
        return n;
    }
    public static int PositiveInteger(string value, string description)
    {
        if (!int.TryParse(value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var n) || n <= 0)
            throw new InvalidDataException($"{description} 必须是正整数。");
        return n;
    }
    public static void SaveJson<T>(string path, T value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        // New catalog format is deliberately stored separately from the old installation.
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try { File.WriteAllText(temporary, JsonSerializer.Serialize(value, Json)); File.Move(temporary, path, true); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}

public static class InspectionValidator
{
    public static List<InspectionGroup> Validate(StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes)
    {
        if (measurements.Count == 0) throw new InvalidDataException("没有称重记录。");
        if (batchSizes.Keys.Except(measurements.Select(m => m.Key)).Any())
            throw new InvalidDataException("存在没有实测记录的生产条目，请补充测量记录或移除空条目。");
        if (measurements.Select(m => m.RecordId).Distinct().Count() != measurements.Count)
            throw new InvalidDataException("发现重复测量记录 ID，不能重复计算同一次测量。");
        var groups = new List<InspectionGroup>();
        var errors = new List<string>();
        foreach (var group in measurements.GroupBy(m => m.Key))
        {
            try
            {
                var first = group.First();
                var product = catalog.Resolve(first.Barcode);
                if (!batchSizes.TryGetValue(group.Key, out var text)) throw new InvalidDataException($"条码 {group.Key} 缺少抽样基数。");
                var batch = DataFiles.PositiveInteger(text, "抽样基数");
                var required = catalog.RequiredSamples(batch);
                var actual = group.ToList();
                if (actual.Any(m => m.Grams <= 0)) throw new InvalidDataException($"条码 {group.Key} 存在非正数测量值。");
                foreach (var field in new[] { actual.Select(m => m.ProductionDate), actual.Select(m => m.SensoryResult), actual.Select(m => m.BatchSize?.ToString() ?? "") })
                    if (field.Where(v => v.Length > 0).Distinct().Count() > 1)
                        throw new InvalidDataException($"条码 {first.Barcode} 同一批次的数量、日期或感官结果不一致。");
                groups.Add(new(product, batch, required, actual, first.BatchId));
            }
            catch (InvalidDataException ex) { errors.Add(ex.Message); }
        }
        if (errors.Count != 0) throw new InvalidDataException(string.Join(Environment.NewLine, errors));
        return groups;
    }
}
