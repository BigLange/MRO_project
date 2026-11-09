using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Mro.Core;
using Mro.Recovered.UI;
using UiState = Mro.Recovered.UI.WindowState;

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
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", FileName = $"称重明细_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx", AddExtension = true, InitialDirectory = state.PreferredOutputDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.OutputWorkbookPath = dialog.FileName;
    }
    private void AddEntry_Click(object sender, RoutedEventArgs e) => state.EntryRows.Add(new());
    private void ImportMeasurements_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredEntryDirectory ?? "" };
        if (dialog.ShowDialog(this) == true)
            state.Run("导入实测称重", () => { UiState.CommitGrid(this); state.ImportMeasurements(dialog.FileName); });
    }
    private void ImportProduction_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredEntryDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.Run("导入生产表", () => { UiState.CommitGrid(this); state.ImportProduction(dialog.FileName); });
    }
    private void DeleteEntry_Click(object sender, RoutedEventArgs e) => state.Run("移除空条目", () => { UiState.CommitGrid(this); state.DeleteEntry(); });
    private void CheckSamples_Click(object sender, RoutedEventArgs e) => state.Run("预览生成样品数", () => { UiState.CommitGrid(this); state.PreviewSamples(); });
    private void Generate_Click(object sender, RoutedEventArgs e) => state.Run("生成表3", () => { UiState.CommitGrid(this); state.GenerateOutputs(); });
    private void ExportMeasurements_Click(object sender, RoutedEventArgs e) => state.Run("生成标准表", () => { UiState.CommitGrid(this); state.GenerateOutputs(); });
    private void FilterLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
    private void RefreshLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
}

public sealed class HelperState : UiState
{
    protected override string ApplicationName => "马里奥产品抽检辅助系统";
    private string outputWorkbookPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"W{DateTime.Now:yyMMdd}.xlsx");
    private ProductRow? selectedEntryRow;
    public string OutputWorkbookPath { get => outputWorkbookPath; set => Set(ref outputWorkbookPath, value); }
    public ProductRow? SelectedEntryRow { get => selectedEntryRow; set => Set(ref selectedEntryRow, value); }
    public ObservableCollection<ProductRow> EntryRows { get; } = [];
    public string? PreferredEntryDirectory { get; private set; }
    public string? PreferredOutputDirectory { get; private set; }
    private bool preferencesLoaded;
    public void Initialize()
    {
        if (!preferencesLoaded)
        {
            var preferences = PreferenceStore.Load();
            PreferredEntryDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.HelperEntryWorkbookDirectory);
            PreferredOutputDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.HelperOutputDirectory);
            if (PreferredOutputDirectory is not null) OutputWorkbookPath = Path.Combine(PreferredOutputDirectory, Path.GetFileName(OutputWorkbookPath));
            preferencesLoaded = true;
        }
        var catalog = GetCatalog();
        StatusMessage = $"共享标准库：{catalog.Products.Count} 个产品。请填写或导入生成条目。";
    }
    public void ImportMeasurements(string path) => ImportProduction(path);
    public void ImportProduction(string path)
    {
        var catalog = GetCatalog();
        var entries = ProductionImporter.Read(path, catalog);
        PreferredEntryDirectory = Path.GetDirectoryName(path);
        foreach (var entry in entries)
            EntryRows.Add(new() { Barcode = entry.Barcode, ProductName = catalog.Resolve(entry.Barcode).ProductName,
                BatchSizeText = entry.BatchSize > 0 ? entry.BatchSize.ToString() : "" });
        StatusMessage = $"已导入 {entries.Count} 条生成条目，请核对抽样基数。";
    }
    public void SavePreferences() => PreferenceStore.Update(old => old with
    {
        HelperEntryWorkbookDirectory = PreferredEntryDirectory,
        HelperOutputDirectory = Path.GetDirectoryName(OutputWorkbookPath)
    });
    public void DeleteEntry()
    {
        if (SelectedEntryRow is not null) EntryRows.Remove(SelectedEntryRow);
        StatusMessage = "已删除选中条目。";
    }
    public void PreviewSamples() => Preview(EntryRows);
    public void GenerateOutputs()
    {
        PreviewSamples();
        var entries = EntryRows.Select(row => new ProductionEntry(row.Barcode.Trim(),
            DataFiles.PositiveInteger(row.BatchSizeText, "抽样基数"), row.ProductName, "", "界面录入", "", 0,
            MaximumWeightGrams: string.IsNullOrWhiteSpace(row.MaximumWeightText) ? null : DataFiles.Decimal(row.MaximumWeightText, "最大重量"))).ToList();
        var exported = Generation.Export(OutputWorkbookPath, GetCatalog(), entries, DateTime.Now);
        OperationOutputPaths = exported.WeighingPath;
        StatusMessage = $"标准表已生成：{exported.Count} 条称重记录。文件路径见运行日志。";
    }
}
