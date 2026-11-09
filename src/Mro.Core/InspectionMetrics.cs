namespace Mro.Core;

public enum LegacySampleJudgement { Pass, T1, T2 }

public sealed record InspectionMetrics(int ActualCount, decimal AverageGrams, decimal MinimumGrams, decimal MaximumGrams,
    int PassCount, int T1Count, int T2Count, int AllowedT1Count, int AllowedT2Count,
    bool? LegacyCountPass, bool AverageMeetsNominal, string ReviewConclusion);

public static class RecoveredJudgement
{
    // Recovered from NetContentJudgement.Evaluate IL, not asserted to be a legal definition of T1/T2.
    public static LegacySampleJudgement Classify(decimal nominal, decimal shortage, decimal actual)
    {
        if (nominal <= 0 || shortage < 0 || shortage >= nominal || actual <= 0)
            throw new InvalidDataException("规格、允许短缺量或实测重量无效。");
        return actual >= nominal ? LegacySampleJudgement.Pass
            : actual >= nominal - shortage ? LegacySampleJudgement.T1 : LegacySampleJudgement.T2;
    }

    public static InspectionMetrics Evaluate(InspectionGroup group, SamplingRule rule)
    {
        if (group.Measurements.Count == 0) throw new InvalidDataException("没有测量记录，无法计算统计。");
        var weights = group.Measurements.Select(m => m.Grams).ToArray();
        var average = weights.Average();
        if (group.Product.AllowedShortageGrams is not decimal shortage)
            return new(weights.Length, average, weights.Min(), weights.Max(), 0, 0, 0,
                rule.AllowedT1Count, rule.AllowedT2Count, null, average >= group.Product.StandardGrams, "缺少短缺量，待复核");
        var classes = weights.Select(w => Classify(group.Product.StandardGrams, shortage, w)).ToArray();
        var t1 = classes.Count(c => c == LegacySampleJudgement.T1);
        var t2 = classes.Count(c => c == LegacySampleJudgement.T2);
        // Original BatchConclusionJudge rejects any T2 and checks only allowed T1, not the average.
        var legacyCountPass = t2 == 0 && t1 <= rule.AllowedT1Count;
        var conclusion = !legacyCountPass || average < group.Product.StandardGrams
            ? "未满足净含量检查条件（旧版口径，待法规复核）"
            : weights.Length != group.RequiredSamples ? "实测数与计划数不同，待复核"
            : "满足旧版净含量检查条件（法规及感官待复核）";
        return new(weights.Length, average, weights.Min(), weights.Max(), classes.Length - t1 - t2, t1, t2,
            rule.AllowedT1Count, rule.AllowedT2Count, legacyCountPass, average >= group.Product.StandardGrams, conclusion);
    }
}
