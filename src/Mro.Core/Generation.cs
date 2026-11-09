using System.Globalization;

namespace Mro.Core;

/// <summary>Generate 80% of values in the outer quarters and 20% in the middle, then shuffle.</summary>
public sealed class PassWeightGenerator(Func<double>? nextDouble = null)
{
    private double Next()
    {
        var value = (nextDouble ?? Random.Shared.NextDouble)();
        if (!double.IsFinite(value) || value < 0 || value >= 1) throw new InvalidDataException("随机数必须位于[0,1)。");
        return value;
    }

    public List<decimal> Generate(decimal minimum, decimal maximum, int count)
    {
        if (minimum <= 0 || maximum < minimum || count <= 0) throw new InvalidDataException("生成范围或样品数无效。");
        var first = decimal.Ceiling(minimum * 10m);
        var last = decimal.Floor(maximum * 10m);
        if (first > last) throw new InvalidDataException("生成范围内没有可保留1位小数的重量，请调整浮动比例或最大重量。");
        decimal Draw(decimal low, decimal high) => Math.Min(high, low + decimal.Floor((decimal)Next() * (high - low + 1m))) / 10m;
        var lowLast = decimal.Floor((minimum + (maximum - minimum) * 0.25m) * 10m);
        var highFirst = decimal.Ceiling((minimum + (maximum - minimum) * 0.75m) * 10m);
        // Very narrow intervals may not contain a tenth of a gram in all three bands.
        if (lowLast < first || highFirst > last || lowLast + 1m > highFirst - 1m)
            return Enumerable.Range(0, count).Select(_ => Draw(first, last)).ToList();
        var edgeCount = (int)decimal.Round(count * 0.8m, 0, MidpointRounding.AwayFromZero);
        var lowCount = edgeCount / 2 + (edgeCount % 2 == 1 && Next() < 0.5 ? 1 : 0);
        var highCount = edgeCount - lowCount;
        var weights = Enumerable.Range(0, lowCount).Select(_ => Draw(first, lowLast))
            .Concat(Enumerable.Range(0, highCount).Select(_ => Draw(highFirst, last)))
            .Concat(Enumerable.Range(0, count - edgeCount).Select(_ => Draw(lowLast + 1m, highFirst - 1m))).ToList();
        for (var i = weights.Count - 1; i > 0; i--)
        {
            var j = (int)Math.Floor(Next() * (i + 1));
            (weights[i], weights[j]) = (weights[j], weights[i]);
        }
        return weights;
    }
}

public static class Generation
{
    public static decimal MaximumWeight(ProductStandard product, decimal? manualMaximum = null)
    {
        var maximum = manualMaximum ?? (product.FloatingRatioPercent is decimal percent && percent >= 100m
            ? product.StandardGrams * percent / 100m
            : throw new InvalidDataException($"条码 {product.Barcode} 缺少有效浮动比例，请填写最大重量。"));
        if (maximum < product.StandardGrams) throw new InvalidDataException($"条码 {product.Barcode} 最大重量不能小于规格。");
        return maximum;
    }

    public static List<Measurement> Build(StandardCatalog catalog, IReadOnlyList<ProductionEntry> entries,
        DateTime generatedAt, PassWeightGenerator? generator = null)
    {
        if (entries.Count == 0) throw new InvalidDataException("请先导入商品表或添加生成条目。");
        var plans = entries.Select(entry =>
        {
            var product = catalog.Resolve(entry.Barcode);
            var count = catalog.RequiredSamples(entry.BatchSize);
            return (Entry: entry, Product: product, Count: count, Maximum: MaximumWeight(product, entry.MaximumWeightGrams));
        }).ToList();
        if (plans.Sum(p => (long)p.Count) > 1_048_575) throw new InvalidDataException("称重记录超过Excel单表上限，请分批生成。");
        generator ??= new();
        var runId = Guid.NewGuid().ToString("N");
        var result = new List<Measurement>();
        foreach (var plan in plans)
        {
            foreach (var weight in generator.Generate(plan.Product.StandardGrams, plan.Maximum, plan.Count))
            {
                if (weight < plan.Product.StandardGrams || weight > plan.Maximum) throw new InvalidDataException("生成重量不在规格和最大值范围内。");
                var sequence = result.Count + 1;
                var instant = generatedAt.AddSeconds(sequence - 1);
                result.Add(new(plan.Entry.Barcode, weight, generatedAt.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture),
                    instant.ToString("HH:mm:ss", CultureInfo.InvariantCulture), 0m, weight,
                    "规则生成:" + runId, "称重明细_规则生成", sequence + 1, $"generated:{runId}:{sequence}",
                    "规则生成", plan.Entry.BatchId, plan.Entry.BatchSize, plan.Entry.ProductionDate, plan.Entry.SensoryResult, plan.Maximum));
            }
        }
        return result;
    }

    public static (string WeighingPath, int Count) Export(string weighingPath, StandardCatalog catalog,
        IReadOnlyList<ProductionEntry> entries, DateTime generatedAt)
    {
        weighingPath = Path.GetFullPath(weighingPath);
        if (!Path.GetExtension(weighingPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("输出文件必须为.xlsx。");
        var measurements = Build(catalog, entries, generatedAt);
        WorkbookExports.Measurements(weighingPath, catalog, measurements, new Dictionary<string, string>());
        return (weighingPath, measurements.Count);
    }
}
