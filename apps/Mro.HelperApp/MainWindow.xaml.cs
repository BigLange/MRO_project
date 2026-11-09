using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Mro.Core;
using Mro.Recovered.UI;

namespace Mro.HelperApp;

public partial class MainWindow : Window
{
    private readonly HelperState state = new();
    public MainWindow()
    {
        InitializeComponent();
        DataContext = state;
        Loaded += (_, _) => state.Run("启动", state.Initialize);
        Closing += (_, _) => state.Run("保存路径偏好", state.SavePreferences);
    }
    private void ReloadStandard_Click(object sender, RoutedEventArgs e) => state.Run("重新加载标准库", state.Initialize);
    private void ChooseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", FileName = $"实测称重_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx", AddExtension = true, InitialDirectory = state.PreferredOutputDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.OutputWorkbookPath = dialog.FileName;
    }
    private void AddEntry_Click(object sender, RoutedEventArgs e) => state.EntryRows.Add(new());
    private void ImportMeasurements_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredEntryDirectory ?? "" };
        if (dialog.ShowDialog(this) == true)
            state.Run("导入实测称重", () => { WindowState.CommitGrid(this); state.ImportMeasurements(dialog.FileName); });
    }
    private void ImportProduction_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredEntryDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.Run("导入生产表", () => { WindowState.CommitGrid(this); state.ImportProduction(dialog.FileName); });
    }
    private void DeleteEntry_Click(object sender, RoutedEventArgs e) => state.Run("移除空条目", () => { WindowState.CommitGrid(this); state.DeleteEntry(); });
    private void CheckSamples_Click(object sender, RoutedEventArgs e) => state.Run("检查实测数量", () => { WindowState.CommitGrid(this); state.CheckSamples(); });
    private void ExportMeasurements_Click(object sender, RoutedEventArgs e) => state.Run("导出实测称重表", () => { WindowState.CommitGrid(this); state.ExportMeasurements(); });
    private void FilterLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
    private void RefreshLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
}

public sealed class HelperState : WindowState
{
    protected override string ApplicationName => "马里奥产品抽检辅助系统";
    private string outputWorkbookPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"实测称重_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    private ProductRow? selectedEntryRow;
    public string OutputWorkbookPath { get => outputWorkbookPath; set => Set(ref outputWorkbookPath, value); }
    public ProductRow? SelectedEntryRow { get => selectedEntryRow; set => Set(ref selectedEntryRow, value); }
    public ObservableCollection<ProductRow> EntryRows { get; } = [];
    public string? PreferredEntryDirectory { get; private set; }
    public string? PreferredOutputDirectory { get; private set; }
    private bool preferencesLoaded;
    private string productionWorkbookPath = "";
    public void Initialize()
    {
        if (!preferencesLoaded)
        {
            var preferences = PreferenceStore.Load();
            PreferredEntryDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.HelperEntryWorkbookDirectory);
            PreferredOutputDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.HelperOutputDirectory);
            if (PreferredOutputDirectory is not null) OutputWorkbookPath = Path.Combine(PreferredOutputDirectory, Path.GetFileName(OutputWorkbookPath));
            if (PreferenceStore.LoadWarning is string warning) IssueRows.Add(new("路径偏好", warning));
            preferencesLoaded = true;
        }
        var catalog = GetCatalog();
        StatusMessage = $"共享标准库：{catalog.Products.Count} 个产品。请导入真实称重表。";
    }
    public void ImportMeasurements(string path)
    {
        var measurements = MeasurementImporter.Read(path);
        var rows = GroupRows(measurements, EntryRows);
        MeasurementFile = path; Measurements = measurements;
        PreferredEntryDirectory = Path.GetDirectoryName(path);
        EntryRows.Clear();
        foreach (var row in rows) EntryRows.Add(row);
        StatusMessage = $"已导入 {measurements.Count} 条实测记录，请输入抽样基数。";
    }
    public void ImportProduction(string path)
    {
        ProductionEntries = ProductionImporter.Read(path, GetCatalog()); productionWorkbookPath = path;
        PreferredEntryDirectory = Path.GetDirectoryName(path);
        var rows = GroupRows(Measurements, EntryRows);
        var production = ProductionEntries.ToDictionary(e => e.Barcode);
        foreach (var row in rows) if (production.TryGetValue(row.Barcode, out var entry)) row.BatchSizeText = entry.BatchSize.ToString();
        EntryRows.Clear(); foreach (var row in rows) EntryRows.Add(row);
        StatusMessage = $"已导入 {ProductionEntries.Count} 条生产记录；请导入对应的真实称重记录。";
    }
    public void SavePreferences() => PreferenceStore.Update(old => old with
    {
        HelperEntryWorkbookDirectory = PreferredEntryDirectory,
        HelperOutputDirectory = Path.GetDirectoryName(OutputWorkbookPath)
    });
    public void DeleteEntry()
    {
        if (SelectedEntryRow is null) return;
        if (SelectedEntryRow.RecordCount > 0) throw new InvalidDataException("实测条目不能在这里删除；如原始记录有误，请修正原表后重新导入。");
        EntryRows.Remove(SelectedEntryRow);
        StatusMessage = "已移除空条目。";
    }
    public void CheckSamples()
    {
        if (productionWorkbookPath.Length > 0) ProductionEntries = ProductionImporter.Read(productionWorkbookPath, GetCatalog());
        if (MeasurementFile.Length > 0)
        {
            var fresh = MeasurementImporter.Read(MeasurementFile);
            var previous = EntryRows.ToList(); var rows = GroupRows(fresh, previous);
            rows.AddRange(previous.Where(p => rows.All(r => r.Barcode != p.Barcode)));
            EntryRows.Clear(); foreach (var row in rows) EntryRows.Add(row);
            Measurements = fresh;
        }
        Check(EntryRows);
    }
    public void ExportMeasurements()
    {
        if (string.IsNullOrWhiteSpace(MeasurementFile)) throw new InvalidDataException("请先导入真实称重记录。");
        CheckSamples();
        WorkbookExports.Measurements(OutputWorkbookPath, GetCatalog(), Measurements, Batches(EntryRows));
        StatusMessage = "实测称重表已导出：" + OutputWorkbookPath;
    }
}
