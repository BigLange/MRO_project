using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Mro.Core;
using Mro.Recovered.UI;
using UiState = Mro.Recovered.UI.WindowState;

namespace Mro.MainApp;

public partial class MainWindow : Window
{
    private readonly MainState state = new();
    public MainWindow()
    {
        InitializeComponent();
        DataContext = state;
        Loaded += (_, _) => state.Run("启动", state.Initialize);
        Closing += (_, _) => state.Run("保存路径偏好", state.SavePreferences);
    }
    private void BrowseStandard_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredStandardDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.StandardWorkbookPath = dialog.FileName;
    }
    private void ImportStandard_Click(object sender, RoutedEventArgs e) => state.Run("导入标准库", state.ImportStandard);
    private void BrowseWeighing_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredWeighingDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.WeighingWorkbookPath = dialog.FileName;
    }
    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "选择报告输出文件夹" };
        if (dialog.ShowDialog(this) == true) state.OutputDirectoryPath = dialog.FolderName;
    }
    private void Analyze_Click(object sender, RoutedEventArgs e) => state.Run("分析称重表", () => { UiState.CommitGrid(this); state.Analyze(); });
    private void ImportProduction_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", CheckFileExists = true, InitialDirectory = state.PreferredProductionDirectory ?? "" };
        if (dialog.ShowDialog(this) == true) state.Run("导入生产表", () => { UiState.CommitGrid(this); state.ImportProduction(dialog.FileName); });
    }
    private void CheckSamples_Click(object sender, RoutedEventArgs e) => state.Run("一键生成标准抽样基数", () => { UiState.CommitGrid(this); state.FillStandardBatchSizes(); });
    private void ExportReport_Click(object sender, RoutedEventArgs e) => state.Run("导出检测报告", () => { UiState.CommitGrid(this); state.ExportReport(); });
    private void FilterLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
    private void RefreshLogs_Click(object sender, RoutedEventArgs e) => state.RefreshLogs();
}

public sealed class MainState : UiState
{
    protected override string ApplicationName => "马里奥产品抽检分类系统";
    private string standardWorkbookPath = Paths.Standard, weighingWorkbookPath = "";
    private string outputDirectoryPath = Path.Combine(AppContext.BaseDirectory, "Exports");
    private DateTime? reportDate = DateTime.Today;
    public string StandardWorkbookPath { get => standardWorkbookPath; set => Set(ref standardWorkbookPath, value); }
    public string WeighingWorkbookPath { get => weighingWorkbookPath; set => Set(ref weighingWorkbookPath, value); }
    public string OutputDirectoryPath { get => outputDirectoryPath; set => Set(ref outputDirectoryPath, value); }
    public DateTime? ReportDate { get => reportDate; set => Set(ref reportDate, value); }
    public ObservableCollection<ProductRow> ProductRows { get; } = [];
    public string? PreferredStandardDirectory { get; private set; }
    public string? PreferredWeighingDirectory { get; private set; }
    public string? PreferredProductionDirectory { get; private set; }
    private string productionWorkbookPath = "";

    public void Initialize()
    {
        var preferences = PreferenceStore.Load();
        StandardWorkbookPath = UserFolderPreferenceStore.GetExistingFile(preferences.MainStandardWorkbookPath) ?? Paths.Standard;
        WeighingWorkbookPath = UserFolderPreferenceStore.GetExistingFile(preferences.MainWeighingWorkbookPath) ?? "";
        OutputDirectoryPath = UserFolderPreferenceStore.GetExistingDirectory(preferences.MainOutputDirectory) ?? OutputDirectoryPath;
        PreferredStandardDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.MainStandardWorkbookDirectory);
        PreferredWeighingDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.MainWeighingWorkbookDirectory);
        PreferredProductionDirectory = UserFolderPreferenceStore.GetExistingDirectory(preferences.MainProductionWorkbookDirectory);
        if (PreferenceStore.LoadWarning is string warning) IssueRows.Add(new("路径偏好", warning));
        if (File.Exists(Paths.Catalog)) { var catalog = GetCatalog(); StatusMessage = $"已加载共享标准库：{catalog.Products.Count} 个产品。请选择称重表。"; }
        else StatusMessage = "请先导入标准库，再选择称重记录。";
    }
    public void ImportStandard()
    {
        var catalog = StandardImporter.Read(StandardWorkbookPath, Paths.Plan);
        DataFiles.SaveJson(Paths.Catalog, catalog);
        Catalog = catalog;
        ProductRows.Clear(); Measurements = []; ProductionEntries = []; productionWorkbookPath = "";
        StatusMessage = $"共享标准库已更新：{catalog.Products.Count} 个产品，{catalog.ImportWarnings.Count} 条提示。";
        foreach (var warning in catalog.ImportWarnings) IssueRows.Add(new("标准表提示", warning));
    }
    public void Analyze()
    {
        ProductionEntries = [];
        var path = Path.GetFullPath(WeighingWorkbookPath);
        var previous = string.Equals(MeasurementFile, path, StringComparison.OrdinalIgnoreCase) ? ProductRows : null;
        var input = MeasurementImporter.Read(WeighingWorkbookPath);

        var rows = GroupRows(input, previous);
        MeasurementFile = path;
        Measurements = input;
        ProductRows.Clear();
        foreach (var row in rows) ProductRows.Add(row);
        StatusMessage = $"已读取 {input.Count} 条称重记录、{rows.Count} 个产品，请录入抽样基数。";
    }
    public void ImportProduction(string path)
    {
        ProductionEntries = ProductionImporter.Read(path, GetCatalog());
        productionWorkbookPath = path;
        PreferredProductionDirectory = Path.GetDirectoryName(path);
        var rows = GroupRows(Measurements, null);
        ProductRows.Clear(); foreach (var row in rows) ProductRows.Add(row);
        StatusMessage = $"已导入 {ProductionEntries.Count} 条生产记录；缺少称重记录的产品会提示缺测。";
    }
    public void SavePreferences() => PreferenceStore.Update(old => old with
    {
        MainStandardWorkbookPath = StandardWorkbookPath, MainStandardWorkbookDirectory = Path.GetDirectoryName(StandardWorkbookPath),
        MainWeighingWorkbookPath = WeighingWorkbookPath, MainWeighingWorkbookDirectory = Path.GetDirectoryName(WeighingWorkbookPath),
        MainOutputDirectory = OutputDirectoryPath, MainProductionWorkbookDirectory = PreferredProductionDirectory
    });
    public void FillStandardBatchSizes()
    {
        if (ProductRows.Count == 0) Analyze();
        var catalog = GetCatalog();
        foreach (var row in ProductRows) row.BatchSizeText = LegacyInspection.StandardBatchSize(catalog, row.RecordCount).ToString();
        StatusMessage = $"已为 {ProductRows.Count} 个产品生成标准抽样基数。";
    }
    public void CheckSamples() => Check(ProductRows);
    public void ExportReport()
    {
        Analyze(); // Re-read current inputs and catalog before exporting; never export stale state.
        CheckSamples();
        if (ReportDate is null) throw new InvalidDataException("请选择报告日期。");
        if (string.IsNullOrWhiteSpace(OutputDirectoryPath)) throw new InvalidDataException("请选择输出文件夹。");
        var path = Path.Combine(OutputDirectoryPath, $"批次检测报告_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        var reports = WorkbookExports.Report(path, GetCatalog(), Measurements, Batches(ProductRows), ReportDate.Value, Paths.ReportTemplate, Entries(ProductRows));
        OperationOutputPaths = string.Join(Environment.NewLine, reports.AllPaths);
        StatusMessage = $"已分别导出 {reports.DetailPaths.Count} 份商品明细及1份汇总，文件路径见运行日志。";
    }
}
