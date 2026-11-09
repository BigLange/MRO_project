using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Mro.Core;

namespace Mro.Recovered.UI;

public abstract class ObservableModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void Changed([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new(name));
    protected void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    { if (EqualityComparer<T>.Default.Equals(field, value)) return; field = value; Changed(name); }
}

public sealed class ProductRow : ObservableModel
{
    private string barcode = "", batchSizeText = "", requiredSampleCount = "";
    public string Barcode { get => barcode; set => Set(ref barcode, value); }
    public string Code { get; set; } = "";
    public string ProductName { get; set; } = "";
    public int RecordCount { get; set; }
    public string BatchSizeText { get => batchSizeText; set => Set(ref batchSizeText, value); }
    // Kept to match the original binding; this is now a read-only measured-record count.
    public string MaximumWeightText => RecordCount.ToString();
    public string RequiredSampleCount { get => requiredSampleCount; set => Set(ref requiredSampleCount, value); }
}

public sealed record IssueRow(string Title, string Message, string DetailsText = "");
public sealed record LogRow(DateTime Time, string Application, string Operation, string Result, string Message, string Path, string Error)
{
    public string Summary => $"{Time:yyyy-MM-dd HH:mm:ss}  {Application}  {Operation}  {Result}  {Message}";
    public string PathSummary => Path;
    public string ErrorSummary => Error;
}

public static class Paths
{
    public static string State => System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MRO-Recovered");
    public static string Catalog => System.IO.Path.Combine(State, "standard-catalog.v2.json");
    public static string Preferences => System.IO.Path.Combine(State, "folder-preferences.json");
    public static string Content => System.IO.Path.Combine(AppContext.BaseDirectory, "SharedContent");
    public static string Standard => System.IO.Path.Combine(Content, "检测标准表.xlsx");
    public static string Plan => System.IO.Path.Combine(Content, "sampling-plan.json");
    public static string ReportTemplate => System.IO.Path.Combine(Content, "ReportTemplates", "MRO-ReportTemplate.xlsx");
    public static StandardCatalog LoadCatalog() => JsonSerializer.Deserialize<StandardCatalog>(File.ReadAllText(Catalog), DataFiles.Json)
        ?? throw new InvalidDataException("共享标准库为空，请在主程序中重新导入。");
}

public static class StartupErrors
{
    public static void Record(Exception error)
    {
        try { Directory.CreateDirectory(Paths.State); File.AppendAllText(System.IO.Path.Combine(Paths.State, "startup-errors.log"), $"{DateTime.Now:O} {error}\n"); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}

public abstract class WindowState : ObservableModel
{
    private string statusMessage = "恢复版：请先导入标准库及真实称重记录。";
    private string logApplicationFilter = "", logOperationFilter = "", logResultFilter = "";
    protected abstract string ApplicationName { get; }
    public string StatusMessage { get => statusMessage; set => Set(ref statusMessage, value); }
    public string StandardStorePath => Paths.Catalog;
    public string LogApplicationFilter { get => logApplicationFilter; set => Set(ref logApplicationFilter, value); }
    public string LogOperationFilter { get => logOperationFilter; set => Set(ref logOperationFilter, value); }
    public string LogResultFilter { get => logResultFilter; set => Set(ref logResultFilter, value); }
    public ObservableCollection<IssueRow> IssueRows { get; } = [];
    public ObservableCollection<LogRow> OperationLogRows { get; } = [];
    protected StandardCatalog? Catalog { get; set; }
    protected List<Measurement> Measurements { get; set; } = [];
    protected string MeasurementFile { get; set; } = "";
    protected List<ProductionEntry> ProductionEntries { get; set; } = [];
    protected UserFolderPreferenceStore PreferenceStore { get; } = new(Paths.Preferences);
    private string LogFile => System.IO.Path.Combine(Paths.State, "operations.jsonl");

    public void Run(string operation, Action action)
    {
        IssueRows.Clear();
        try { action(); AppendLog(operation, "Success", StatusMessage, ""); }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException or JsonException or InvalidOperationException or System.Xml.XmlException or OverflowException)
        {
            StatusMessage = ex.Message;
            IssueRows.Add(new("操作未完成", ex.Message));
            AppendLog(operation, "Failure", "操作未完成", ex.Message);
        }
        RefreshLogs();
    }

    protected StandardCatalog GetCatalog()
    {
        // Re-read to observe updates made by the other app.
        Catalog = Paths.LoadCatalog();
        StandardImporter.ValidateRules(Catalog.SamplingRules);
        return Catalog;
    }

    protected List<ProductRow> GroupRows(List<Measurement> input, IEnumerable<ProductRow>? previous = null)
    {
        var catalog = GetCatalog();
        var batches = (previous ?? []).GroupBy(r => r.Barcode).ToDictionary(g => g.Key, g => g.First().BatchSizeText);
        var rows = new List<ProductRow>();
        var measurementGroups = input.GroupBy(m => m.Barcode).ToDictionary(g => g.Key, g => g.Count());
        var production = ProductionEntries.ToDictionary(e => e.Barcode);
        foreach (var barcode in measurementGroups.Keys.Union(production.Keys))
        {
            ProductStandard? product = null;
            try { product = catalog.Resolve(barcode); }
            catch (InvalidDataException ex) { IssueRows.Add(new("条码关联失败", ex.Message)); }
            var batch = batches.GetValueOrDefault(barcode, "");
            rows.Add(new() { Barcode = barcode, Code = product?.Code ?? "", ProductName = product?.ProductName ?? "未匹配或条码重复",
                RecordCount = measurementGroups.GetValueOrDefault(barcode), BatchSizeText = batch.Length > 0 ? batch : production.GetValueOrDefault(barcode)?.BatchSize.ToString() ?? "" });
        }
        return rows;
    }

    protected Dictionary<string, string> Batches(IEnumerable<ProductRow> rows)
    {
        var list = rows.ToList();
        if (list.Any(r => string.IsNullOrWhiteSpace(r.Barcode))) throw new InvalidDataException("有条目没有填写条码。");
        if (list.GroupBy(r => r.Barcode.Trim()).Any(g => g.Count() > 1)) throw new InvalidDataException("存在重复条码条目，请合并为一条并核实生产数量。");
        return list.ToDictionary(r => r.Barcode.Trim(), r => r.BatchSizeText);
    }

    protected void Check(IEnumerable<ProductRow> rows)
    {
        var catalog = GetCatalog();
        var rowList = rows.ToList();
        foreach (var row in rowList)
        {
            try
            {
                row.RequiredSampleCount = catalog.RequiredSamples(DataFiles.PositiveInteger(row.BatchSizeText, "抽样基数")).ToString();
            }
            catch (InvalidDataException) { row.RequiredSampleCount = "待核实"; }
        }
        InspectionValidator.Validate(catalog, Measurements, Batches(rowList));
        StatusMessage = $"抽样数量核对通过：{rowList.Count} 个产品，{Measurements.Count} 条真实测量记录。";
    }

    private void AppendLog(string operation, string result, string message, string error)
    {
        try
        {
            Directory.CreateDirectory(Paths.State);
            var entry = new LogRow(DateTime.Now, ApplicationName, operation, result, message, MeasurementFile, error);
            File.AppendAllText(LogFile, JsonSerializer.Serialize(entry) + Environment.NewLine);
        }
        catch (IOException ex) { IssueRows.Add(new("日志保存失败", ex.Message)); }
        catch (UnauthorizedAccessException ex) { IssueRows.Add(new("日志保存失败", ex.Message)); }
    }

    public void RefreshLogs()
    {
        OperationLogRows.Clear();
        if (!File.Exists(LogFile)) return;
        try
        {
            var recent = File.ReadLines(LogFile).TakeLast(1000).ToList();
            foreach (var line in recent.AsEnumerable().Reverse())
            {
                LogRow? row;
                try { row = JsonSerializer.Deserialize<LogRow>(line); } catch (JsonException) { continue; }
                if (row is null || (!string.IsNullOrEmpty(LogApplicationFilter) && row.Application != LogApplicationFilter)
                    || (!string.IsNullOrEmpty(LogResultFilter) && row.Result != LogResultFilter)
                    || (!string.IsNullOrEmpty(LogOperationFilter) && !row.Operation.Contains(LogOperationFilter, StringComparison.OrdinalIgnoreCase))) continue;
                OperationLogRows.Add(row);
                if (OperationLogRows.Count >= 200) break;
            }
        }
        catch (IOException ex) { IssueRows.Add(new("日志读取失败", ex.Message)); }
    }

    public static void CommitGrid(Window window)
    {
        void Commit(DependencyObject root)
        {
            if (root is DataGrid grid)
            {
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true))
                    throw new InvalidDataException("表格编辑尚未通过校验，请先完成当前单元格输入。");
            }
            for (var i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(root); i++)
                Commit(System.Windows.Media.VisualTreeHelper.GetChild(root, i));
        }
        Commit(window);
    }
}
