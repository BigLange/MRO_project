namespace Mro.Core;

public sealed record ReportExportResult(IReadOnlyList<string> DetailPaths, string SummaryPath)
{
    public IEnumerable<string> AllPaths => DetailPaths.Append(SummaryPath);
}

internal static class ReportFileNames
{
    public static ReportExportResult Create(string prefixPath, IReadOnlyList<InspectionGroup> groups)
    {
        prefixPath = Path.GetFullPath(prefixPath);
        if (!Path.GetExtension(prefixPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("输出文件必须为.xlsx。");
        var directory = Path.GetDirectoryName(prefixPath)!;
        var prefix = Path.GetFileNameWithoutExtension(prefixPath);
        string Clean(string text) => new(text.Select(c => c < 32 || "<>:\"/\\|?*".Contains(c) ? '_' : c).ToArray());
        string FileName(string suffix) => Path.Combine(directory, Clean(prefix + suffix) + ".xlsx");
        var summary = FileName("_批次检测报告结果");
        var details = new List<string>();
        foreach (var group in groups)
        {
            var product = "_" + group.Product.ProductName;
            var path = FileName(product + "_批次检测报告");
            details.Add(path);
        }
        return new(details, summary);
    }
}

internal sealed record PendingWorkbook(string Path, Action<string> Write);

internal static class WorkbookExportSet
{
    // Preserve legacy replacement behavior, with rollback if publishing the new set fails.
    public static void Write(IReadOnlyList<PendingWorkbook> workbooks)
    {
        var paths = workbooks.Select(w => Path.GetFullPath(w.Path)).ToArray();
        if (paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != paths.Length)
            throw new InvalidDataException("导出文件名重复。");
        if (paths.Any(Directory.Exists)) throw new IOException("输出文件名对应已有文件夹。");
        var folder = Path.GetDirectoryName(paths[0])!;
        if (paths.Any(p => !string.Equals(Path.GetDirectoryName(p), folder, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("同批报告必须导出到同一文件夹。");
        Directory.CreateDirectory(folder);
        var staging = Path.Combine(folder, ".mro-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(staging);
        var moved = new List<string>();
        var backups = new Dictionary<string, string>();
        try
        {
            for (var i = 0; i < workbooks.Count; i++) workbooks[i].Write(Path.Combine(staging, i + ".xlsx"));
            for (var i = 0; i < paths.Length; i++)
                if (File.Exists(paths[i]))
                {
                    var backup = Path.Combine(staging, i + ".backup");
                    File.Move(paths[i], backup); backups.Add(paths[i], backup);
                }
            for (var i = 0; i < workbooks.Count; i++)
            {
                File.Move(Path.Combine(staging, i + ".xlsx"), paths[i]);
                moved.Add(paths[i]);
            }
        }
        catch
        {
            foreach (var path in moved) File.Delete(path);
            foreach (var backup in backups) File.Move(backup.Value, backup.Key);
            throw;
        }
        finally { Directory.Delete(staging, true); }
    }
}
