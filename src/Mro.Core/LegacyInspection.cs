using System.Globalization;

namespace Mro.Core;

public static class LegacyInspection
{
    public static decimal Average(InspectionGroup group) => decimal.Round(group.AverageGrams, 1, MidpointRounding.AwayFromZero);

    public static int StandardBatchSize(StandardCatalog catalog, int sampleCount)
    {
        foreach (var rule in catalog.SamplingRules.OrderBy(r => r.MaximumBatch))
        {
            if (rule.SampleEntireBatch && sampleCount <= rule.MaximumBatch) return sampleCount;
            if (!rule.SampleEntireBatch && rule.SampleCount == sampleCount) return rule.MaximumBatch;
        }
        return sampleCount;
    }

    public static List<InspectionGroup> Prepare(StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes, Func<double>? nextDouble = null)
    {
        var groups = InspectionValidator.Validate(catalog, measurements, batchSizes);
        return groups.OrderBy(g => g.Product.Code, ProductCodeComparer.Instance).Select(group =>
        {
            if (group.Product.AllowedShortageGrams is not decimal shortage) throw new InvalidDataException("标准表缺少允许短缺量。");
            var samples = group.Measurements.ToList();
            if (samples.Count < group.RequiredSamples)
            {
                var minimum = samples.Min(m => m.Grams);
                var maximum = samples.Max(m => m.Grams);
                if (maximum >= group.Product.StandardGrams) minimum = group.Product.StandardGrams;
                var missing = group.RequiredSamples - samples.Count;
                for (var i = 0; i < missing; i++)
                {
                    var position = Enumerable.Range(0, 6).Average(_ => (decimal)(nextDouble ?? Random.Shared.NextDouble)());
                    var grams = decimal.Round(minimum + (maximum - minimum) * position, 1, MidpointRounding.AwayFromZero);
                    samples.Add(samples[0] with { Grams = grams, GrossGrams = grams, DataOrigin = "规则生成", RecordId = "SUPP-" + (i + 1) });
                }
            }
            foreach (var judgement in new[] { LegacySampleJudgement.Pass, LegacySampleJudgement.T1, LegacySampleJudgement.T2 })
                for (var i = samples.Count - 1; i >= 0 && samples.Count > group.RequiredSamples; i--)
                    if (RecoveredJudgement.Classify(group.Product.StandardGrams, shortage, samples[i].Grams) == judgement) samples.RemoveAt(i);
            return group with { Measurements = samples };
        }).ToList();
    }
}

public sealed class ProductCodeComparer : IComparer<string>
{
    public static ProductCodeComparer Instance { get; } = new();
    public int Compare(string? x, string? y)
    {
        static (string Letter, int? Number) Parse(string? value)
        {
            var text = value?.Trim() ?? "";
            if (text.Length == 0) return ("", null);
            return (text[..1], int.TryParse(text[1..], NumberStyles.None, CultureInfo.InvariantCulture, out var n) ? n : null);
        }
        var a = Parse(x); var b = Parse(y);
        var letter = string.Compare(a.Letter, b.Letter, StringComparison.OrdinalIgnoreCase);
        if (letter != 0) return letter;
        if (a.Number.HasValue && b.Number.HasValue)
        {
            var number = a.Number.Value.CompareTo(b.Number.Value);
            if (number != 0) return number;
        }
        else if (a.Number.HasValue) return -1;
        else if (b.Number.HasValue) return 1;
        return string.Compare(x, y, StringComparison.Ordinal);
    }
}
