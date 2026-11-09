namespace Mro.Core;

public static class WorkbookExports
{
    public static void Report(string path, StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes, DateTime reportDate, string? templatePath = null,
        IReadOnlyList<ProductionEntry>? production = null)
    {
        if (templatePath is not null)
        {
            TemplateReportExporter.Write(path, templatePath, catalog, measurements, batchSizes, reportDate, production);
            return;
        }
        var groups = InspectionValidator.Validate(catalog, measurements, batchSizes);
        var summary = new WorkbookPage("汇总", [
            ["实测称重数据汇总（未自动判定合格）"],
            ["条码", "编码", "样品名称", "标准规格(g)", "生产数量", "应抽样数", "实测数量", "实测平均值(g)", "报告日期", "明细工作表", "T1(旧版)", "T2(旧版)", "最小值(g)", "最大值(g)", "复核结果"]
        ], ["A1:J1"]);
        var pages = new List<WorkbookPage> { summary };
        for (var i = 0; i < groups.Count; i++)
        {
            var group = groups[i];
            var metrics = RecoveredJudgement.Evaluate(group, catalog.SamplingRuleFor(group.BatchSize));
            var name = $"明细{i + 1:000}"; // Stable and valid even if product names contain invalid sheet characters.
            summary.Rows.Add([group.Product.Barcode, group.Product.Code, group.Product.ProductName,
                group.Product.StandardGrams, group.BatchSize, group.RequiredSamples, group.Measurements.Count,
                decimal.Round(group.AverageGrams, 3, MidpointRounding.AwayFromZero), reportDate.ToString("yyyy-MM-dd"), name,
                metrics.T1Count, metrics.T2Count, metrics.MinimumGrams, metrics.MaximumGrams, metrics.ReviewConclusion]);
            var detail = new WorkbookPage(name, [
                [$"{group.Product.ProductName} — 实测称重记录"],
                ["序号", "日期", "时间", "实测重量(g)", "标准规格(g)", "条码", "来源文件", "来源工作表", "来源行", "测量记录ID"]
            ], ["A1:J1"]);
            detail.Rows.AddRange(group.Measurements.Select((m, j) => new object?[] {
                j + 1, m.Date, m.Time, m.Grams, group.Product.StandardGrams, m.Barcode,
                m.SourceFile, m.SourceSheet, m.SourceRow, m.RecordId }));
            pages.Add(detail);
        }
        pages.Add(new("数据说明", [
            ["真实测量数据说明"], ["项目", "值"],
            ["报告日期", reportDate.ToString("yyyy-MM-dd")], ["标准来源", catalog.SourceFile],
            ["标准文件SHA-256", catalog.SourceSha256],
            ["数据范围", "每条明细对应导入的称重记录；不补测量值、不截取有利结果。"],
            ["平均值", "使用每个产品全部导入记录计算；显示时四舍五入至3位小数。"],
            ["判定范围", "已恢复旧版计数口径并显示平均值；法规尚未独立核验，不能作为自动合格证明。"],
            ["来源字段", "用于回查输入记录，未经仪器签名验证，不构成数据真实性认证。"],
            ["恢复状态", "这是简表格式；使用模板输出可保留原报告表单的样式。"]
        ]));
        XlsxWriter.Write(path, pages);
    }

    public static void Measurements(string path, StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes)
    {
        InspectionValidator.Validate(catalog, measurements, batchSizes);
        var rows = new List<object?[]> {
            ["序号", "日期", "时间", "皮重", "重量", "毛重", "单位", "扫描", "来源文件", "来源工作表", "来源行", "测量记录ID"]
        };
        rows.AddRange(measurements.Select((m, i) => new object?[] {
            (i + 1).ToString("D5"), m.Date, m.Time, m.TareGrams, m.Grams, m.GrossGrams, "g", m.Barcode,
            m.SourceFile, m.SourceSheet, m.SourceRow, m.RecordId }));
        XlsxWriter.Write(path, [new("实测称重", rows)]);
    }
}
