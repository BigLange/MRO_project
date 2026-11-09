namespace Mro.Core;

public static class WorkbookExports
{
    public static ReportExportResult Report(string path, StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes, DateTime reportDate, string? templatePath = null,
        IReadOnlyList<ProductionEntry>? production = null) =>
        TemplateReportExporter.Write(path, templatePath, catalog, measurements, batchSizes, reportDate, production);

    public static void Measurements(string path, StandardCatalog catalog, IReadOnlyList<Measurement> measurements,
        IReadOnlyDictionary<string, string> batchSizes) => TemplateReportExporter.WriteWeighing(path, measurements);
}
