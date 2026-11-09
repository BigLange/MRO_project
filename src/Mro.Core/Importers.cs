using System.Text.Json;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Mro.Core;

public static class StandardImporter
{
    public static StandardCatalog Read(string path, string? samplingPlanPath = null)
    {
        var candidates = XlsxReader.Read(path).SelectMany(sheet => sheet.Rows
            .Where(row => row.Cells.Values.Any(c => c.Text.Trim() == "条码") && row.Cells.Values.Any(c => c.Text.Trim() == "样品名称"))
            .Select(row => (Sheet: sheet, Header: row))).ToList();
        if (candidates.Count != 1) throw new InvalidDataException("标准表必须有唯一的产品表头，至少包含条码、样品名称、编码、规格（g）。");
        var (sheet, header) = candidates[0];
        int Column(string name) => header.Cells.SingleOrDefault(p => p.Value.Text.Trim() == name).Key is var c && c > 0
            ? c : throw new InvalidDataException($"标准表缺少列：{name}");
        var barcodeColumn = Column("条码"); var codeColumn = Column("编码");
        var nameColumn = Column("样品名称"); var weightColumn = Column("规格（g）");
        int Optional(string name) => header.Cells.FirstOrDefault(p => p.Value.Text.Trim() == name).Key;
        var products = new List<ProductStandard>(); var warnings = new List<string>();
        foreach (var row in sheet.Rows.Where(r => r.Number > header.Number))
        {
            // Product rows must have a serial number. The footer contains separate rules.
            if (!int.TryParse(row.At(1), out var serial) || serial <= 0) continue;
            var barcode = DataFiles.Identifier(row.At(barcodeColumn));
            var name = row.At(nameColumn); var code = row.At(codeColumn);
            if (barcode.Length == 0 || name.Length == 0 || code.Length == 0) throw new InvalidDataException($"标准表第 {row.Number} 行缺少条码、编码或名称。");
            var weight = DataFiles.Decimal(row.At(weightColumn), $"标准表第 {row.Number} 行规格");
            if (weight <= 0) throw new InvalidDataException($"标准表第 {row.Number} 行规格必须大于零。");
            if (row.Cells[barcodeColumn].IsNumeric) warnings.Add($"第 {row.Number} 行条码为数字单元格，原前导零无法推断，请核对。");
            var shortage = DataFiles.Decimal(row.At(10), $"标准表第 {row.Number} 行允许短缺量");
            if (shortage < 0 || shortage >= weight) throw new InvalidDataException($"标准表第 {row.Number} 行允许短缺量必须非负且小于规格。");
            products.Add(new(barcode, code, name, weight, row.At(Optional("样品分类")), row.At(8), row.At(Optional("检验依据")), row.At(Optional("取样地点")), shortage, row.At(7)));
        }
        if (products.Count == 0) throw new InvalidDataException("标准表中没有有效产品。");
        foreach (var duplicate in products.GroupBy(p => p.Barcode).Where(g => g.Count() > 1))
            warnings.Add($"重复条码 {duplicate.Key}，关联此条码时会阻止导出。");
        if (products.GroupBy(p => p.Code).Any(g => g.Count() > 1)) warnings.Add("存在重复产品编码，请核对。");
        // Recovered IL: each row defines an upper batch bound, not an isolated typical value.
        var rules = samplingPlanPath is not null && File.Exists(samplingPlanPath)
            ? JsonSerializer.Deserialize<List<SamplingRule>>(File.ReadAllText(samplingPlanPath), DataFiles.Json)
                ?? throw new InvalidDataException("sampling-plan.json 为空。")
            : ReadRecoveredRules(sheet);
        ValidateRules(rules);
        warnings.Add("抽样区间和 T1/T2 采用原程序口径，尚未独立核实法规；部分小批量的应抽数可能超过批量，会阻止导出。");
        warnings.Add("允许短缺量按原程序读取 J 列，净含量文字要求读取 H 列；两列可能不一致，报告会显示待复核说明。");
        return new(products, rules, Path.GetFullPath(path), DataFiles.Sha256(path), warnings);
    }

    private static List<SamplingRule> ReadRecoveredRules(WorkbookSheet sheet)
    {
        var rules = new List<SamplingRule>();
        var header = sheet.Rows.FirstOrDefault(r => r.At(7).Contains("检验批量", StringComparison.Ordinal))
            ?? throw new InvalidDataException("未找到原格式抽样规则表头，请配置 sampling-plan.json。");
        var previousUpper = 0;
        foreach (var row in sheet.Rows.Where(r => r.Number > header.Number))
        {
            var batch = row.At(7); var count = row.At(9);
            if (batch.Length == 0) continue; // Merged subranges are not independent rules.
            var numbers = Regex.Matches(batch, @"\d+");
            if (numbers.Count == 0 || count.Length == 0) { if (rules.Count > 0) break; else continue; }
            var upper = int.Parse(numbers[^1].Value, CultureInfo.InvariantCulture);
            var entire = count.Equals("N", StringComparison.OrdinalIgnoreCase);
            var samples = entire ? 0 : DataFiles.PositiveInteger(count, "应抽样数");
            if (!int.TryParse(row.At(10), out var t1) || !int.TryParse(row.At(11), out var t2))
                throw new InvalidDataException($"规则第 {row.Number} 行缺少 T1/T2 允许数量。");
            rules.Add(new(previousUpper + 1, upper, samples, entire, t1, t2));
            previousUpper = upper;
        }
        if (rules.Count == 0) throw new InvalidDataException("没有识别到抽样规则，请提供已核实的 sampling-plan.json。");
        return rules;
    }

    public static void ValidateRules(List<SamplingRule> rules)
    {
        if (rules.Count == 0) throw new InvalidDataException("抽样规则不能为空。");
        var ordered = rules.OrderBy(r => r.MinimumBatch).ToArray();
        for (var i = 0; i < ordered.Length; i++)
        {
            var r = ordered[i];
            if (r.MinimumBatch <= 0 || r.MaximumBatch < r.MinimumBatch || r.AllowedT1Count < 0 || r.AllowedT2Count < 0
                || (!r.SampleEntireBatch && (r.SampleCount <= 0 || r.SampleCount > r.MaximumBatch)))
                throw new InvalidDataException("抽样规则范围或样本数无效。");
            if (i > 0 && ordered[i - 1].MaximumBatch >= r.MinimumBatch) throw new InvalidDataException("抽样规则区间重叠。");
        }
    }
}

public static class MeasurementImporter
{
    public static List<Measurement> Read(string path)
    {
        var hash = DataFiles.Sha256(path);
        var result = new List<Measurement>();
        var matchedSheets = 0;
        foreach (var sheet in XlsxReader.Read(path))
        {
            var headers = sheet.Rows.Where(r => r.Cells.Values.Any(c => c.Text.Trim() == "扫描") && r.Cells.Values.Any(c => c.Text.Trim() == "重量")).ToList();
            if (headers.Count == 0) continue;
            if (headers.Count != 1) throw new InvalidDataException($"工作表 {sheet.Name} 有多个称重表头。");
            matchedSheets++;
            var header = headers[0];
            int Col(string name) => header.Cells.FirstOrDefault(p => p.Value.Text.Trim() == name).Key;
            var barcodeColumn = Col("扫描"); var weightColumn = Col("重量"); var unitColumn = Col("单位");
            if (unitColumn == 0) throw new InvalidDataException("称重表缺少单位列，不能猜测 kg/g。");
            foreach (var row in sheet.Rows.Where(r => r.Number > header.Number))
            {
                if (row.Cells.Values.All(c => string.IsNullOrWhiteSpace(c.Text) && !c.IsFormula)) continue;
                var barcode = DataFiles.Identifier(row.At(barcodeColumn));
                if (barcode.Length == 0) throw new InvalidDataException($"{sheet.Name} 第 {row.Number} 行缺少扫描条码。");
                var unit = row.At(unitColumn).ToLowerInvariant();
                var factor = unit switch { "kg" or "千克" or "公斤" => 1000m, "g" or "克" => 1m, _ => throw new InvalidDataException($"{sheet.Name} 第 {row.Number} 行单位无法识别：{unit}") };
                var weight = DataFiles.Decimal(row.At(weightColumn), $"{sheet.Name} 第 {row.Number} 行重量") * factor;
                if (weight <= 0) throw new InvalidDataException($"{sheet.Name} 第 {row.Number} 行重量必须大于零。");
                decimal? OptionalWeight(string name)
                {
                    var text = row.At(Col(name));
                    if (text.Length == 0) return null;
                    var n = DataFiles.Decimal(text, name) * factor;
                    return n >= 0 ? n : throw new InvalidDataException($"第 {row.Number} 行{name}不能是负数。");
                }
                var originFile = row.At(Col("来源文件")); var originSheet = row.At(Col("来源工作表"));
                var originRow = row.At(Col("来源行")); var recordId = row.At(Col("测量记录ID"));
                var hasOrigin = originFile.Length > 0 && originSheet.Length > 0 && int.TryParse(originRow, out var parsedRow) && parsedRow > 0;
                if (!hasOrigin && (originFile.Length > 0 || originSheet.Length > 0 || originRow.Length > 0 || recordId.Length > 0))
                    throw new InvalidDataException($"第 {row.Number} 行来源字段不完整。");
                string Temporal(string field, bool date)
                {
                    var column = Col(field); var text = row.At(column);
                    if (text.Length == 0 || !row.Cells.TryGetValue(column, out var cell) || !cell.IsNumeric) return text;
                    if (date && DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var ymd))
                        return ymd.ToString("yyyy-MM-dd");
                    if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial)) return text;
                    var instant = DateTime.FromOADate(serial + (sheet.Date1904 ? 1462 : 0));
                    return instant.ToString(date ? "yyyy-MM-dd" : "HH:mm:ss.fff", CultureInfo.InvariantCulture);
                }
                result.Add(new(barcode, weight, Temporal("日期", true), Temporal("时间", false), OptionalWeight("皮重"), OptionalWeight("毛重"),
                    hasOrigin ? originFile : Path.GetFullPath(path), hasOrigin ? originSheet : sheet.Name,
                    hasOrigin ? int.Parse(originRow) : row.Number,
                    hasOrigin && recordId.Length > 0 ? recordId : $"{hash}:{sheet.Name}:{row.Number}"));
            }
        }
        if (matchedSheets == 0) throw new InvalidDataException("未找到包含扫描、重量和单位列的称重工作表。");
        if (result.Count == 0) throw new InvalidDataException("称重表没有记录。");
        if (result.Select(m => m.RecordId).Distinct().Count() != result.Count) throw new InvalidDataException("发现重复测量记录 ID。");
        return result;
    }
}
