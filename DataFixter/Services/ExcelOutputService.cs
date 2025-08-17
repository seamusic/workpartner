using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataFixter.Models;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace DataFixter.Services
{
    /// <summary>
    /// Excel输出服务
    /// 实现生成修正后的Excel文件、生成详细的修正报告和统计信息功能
    /// </summary>
    public class ExcelOutputService
    {
        private readonly ILogger<ExcelOutputService> _logger;
        private readonly OutputOptions _options;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">输出选项</param>
        public ExcelOutputService(ILogger<ExcelOutputService> logger, OutputOptions? options = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _options = options ?? new OutputOptions();
        }

        /// <summary>
        /// 生成修正后的Excel文件
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="originalDirectory">原始文件目录</param>
        /// <returns>输出结果</returns>
        public OutputResult GenerateCorrectedExcelFiles(
            List<MonitoringPoint> monitoringPoints, 
            string outputDirectory, 
            string originalDirectory)
        {
            var result = new OutputResult();
            var totalFiles = 0;
            var processedFiles = 0;

            try
            {
                _logger.LogInformation("开始生成修正后的Excel文件，输出目录: {OutputDirectory}", outputDirectory);

                // 确保输出目录存在
                Directory.CreateDirectory(outputDirectory);

                // 按文件分组监测点数据
                var dataByFile = GroupDataByFile(monitoringPoints);
                totalFiles = dataByFile.Count;

                foreach (var fileGroup in dataByFile)
                {
                    try
                    {
                        var fileName = fileGroup.Key;
                        var fileData = fileGroup.Value;

                        // 生成修正后的Excel文件
                        var outputPath = Path.Combine(outputDirectory, GetCorrectedFileName(fileName));
                        var fileResult = GenerateSingleExcelFile(fileName, fileData, outputPath, originalDirectory);
                        
                        result.AddFileResult(fileResult);
                        processedFiles++;

                        if (processedFiles % 5 == 0)
                        {
                            _logger.LogInformation("已处理 {ProcessedFiles}/{TotalFiles} 个文件", processedFiles, totalFiles);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "生成文件 {FileName} 时发生异常", fileGroup.Key);
                        
                        result.AddFileResult(new FileOutputResult
                        {
                            FileName = fileGroup.Key,
                            Status = OutputStatus.Error,
                            Message = $"生成文件时发生异常: {ex.Message}"
                        });
                    }
                }

                _logger.LogInformation("Excel文件生成完成: 总计 {TotalFiles} 个文件, 成功 {SuccessCount} 个", 
                    totalFiles, result.FileResults.Count(r => r.Status == OutputStatus.Success));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成修正后的Excel文件时发生异常");
                result.Status = OutputStatus.Error;
                result.Message = $"生成过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 生成单个Excel文件
        /// </summary>
        /// <param name="originalFileName">原始文件名</param>
        /// <param name="fileData">文件数据</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="originalDirectory">原始文件目录</param>
        /// <returns>文件输出结果</returns>
        private FileOutputResult GenerateSingleExcelFile(
            string originalFileName, 
            List<PeriodData> fileData, 
            string outputPath, 
            string originalDirectory)
        {
            var result = new FileOutputResult
            {
                FileName = originalFileName,
                Status = OutputStatus.Success,
                Message = "文件生成成功"
            };

            try
            {
                // 读取原始Excel文件
                var originalFilePath = Path.Combine(originalDirectory, originalFileName);
                if (!File.Exists(originalFilePath))
                {
                    result.Status = OutputStatus.Error;
                    result.Message = $"原始文件不存在: {originalFilePath}";
                    return result;
                }

                // 根据文件扩展名选择合适的工作簿类型
                IWorkbook workbook;
                using (var originalStream = new FileStream(originalFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (Path.GetExtension(originalFileName).ToLower() == ".xlsx")
                    {
                        workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(originalStream);
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(originalStream);
                    }
                }

                // 获取第一个工作表
                var sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    result.Status = OutputStatus.Error;
                    result.Message = "原始文件中没有找到工作表";
                    return result;
                }

                // 修正数据：只修改需要修正的单元格，保持原始格式
                var correctedRows = CorrectDataInSheet(sheet, fileData);

                // 保存修正后的文件
                using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(fileStream);

                result.OutputPath = outputPath;
                result.RowCount = correctedRows;
                result.Message = $"成功修正 {correctedRows} 个数据项，保持原始格式";
            }
            catch (Exception ex)
            {
                result.Status = OutputStatus.Error;
                result.Message = $"修正文件时发生异常: {ex.Message}";
                _logger.LogError(ex, "修正文件 {FileName} 时发生异常", originalFileName);
            }

            return result;
        }

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="sheet">工作表</param>
        private void SetColumnHeaders(ISheet sheet)
        {
            var headerRow = sheet.CreateRow(3); // 第4行（索引3）
            var headers = new[]
            {
                "点名", "里程", "本期变化量X", "本期变化量Y", "本期变化量Z",
                "累计变化量X", "累计变化量Y", "累计变化量Z",
                "日变化量X", "日变化量Y", "日变化量Z"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                
                // 设置标题样式
                var style = sheet.Workbook.CreateCellStyle();
                var font = sheet.Workbook.CreateFont();
                font.IsBold = true;
                style.SetFont(font);
                cell.CellStyle = style;
            }
        }

        /// <summary>
        /// 写入数据行
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="data">数据</param>
        /// <param name="rowNumber">行号</param>
        private void WriteDataRow(IRow row, PeriodData data, int rowNumber)
        {
            var colIndex = 0;

            // 点名
            row.CreateCell(colIndex++).SetCellValue(data.PointName ?? "");

            // 里程
            row.CreateCell(colIndex++).SetCellValue(data.Mileage);

            // 本期变化量
            row.CreateCell(colIndex++).SetCellValue(data.CurrentPeriodX);
            row.CreateCell(colIndex++).SetCellValue(data.CurrentPeriodY);
            row.CreateCell(colIndex++).SetCellValue(data.CurrentPeriodZ);

            // 累计变化量
            row.CreateCell(colIndex++).SetCellValue(data.CumulativeX);
            row.CreateCell(colIndex++).SetCellValue(data.CumulativeY);
            row.CreateCell(colIndex++).SetCellValue(data.CumulativeZ);

            // 日变化量
            row.CreateCell(colIndex++).SetCellValue(data.DailyX);
            row.CreateCell(colIndex++).SetCellValue(data.DailyY);
            row.CreateCell(colIndex++).SetCellValue(data.DailyZ);
        }

        /// <summary>
        /// 修正数据：只修改需要修正的单元格，保持原始格式
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="fileData">文件数据</param>
        /// <returns>修正的行数</returns>
        private int CorrectDataInSheet(ISheet sheet, List<PeriodData> fileData)
        {
            var correctedRows = 0;
            
            // 首先尝试自动检测列位置
            var columnMapping = DetectColumnMapping(sheet);
            
            foreach (var data in fileData)
            {
                if (data.RowNumber > 0 && data.RowNumber <= sheet.LastRowNum + 1)
                {
                    var rowIndex = data.RowNumber - 1; // 转换为0基索引
                    var row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        // 修正累计变化量列
                        if (columnMapping.CumulativeX >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeX) ?? row.CreateCell(columnMapping.CumulativeX);
                            cell.SetCellValue(data.CumulativeX);
                            correctedRows++;
                        }
                        if (columnMapping.CumulativeY >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeY) ?? row.CreateCell(columnMapping.CumulativeY);
                            cell.SetCellValue(data.CumulativeY);
                            correctedRows++;
                        }
                        if (columnMapping.CumulativeZ >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeZ) ?? row.CreateCell(columnMapping.CumulativeZ);
                            cell.SetCellValue(data.CumulativeZ);
                            correctedRows++;
                        }
                    }
                }
            }
            
            return correctedRows;
        }

        /// <summary>
        /// 自动检测列映射
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <returns>列映射信息</returns>
        private ColumnMapping DetectColumnMapping(ISheet sheet)
        {
            var mapping = new ColumnMapping();
            
            // 查找标题行（通常是第4行，索引3）
            var headerRow = sheet.GetRow(3);
            if (headerRow == null) return mapping;
            
            // 遍历标题行，查找相关列
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell == null) continue;
                
                var cellValue = cell.StringCellValue?.ToLower() ?? "";
                
                // 检测累计变化量列
                if (cellValue.Contains("累计") && cellValue.Contains("x"))
                    mapping.CumulativeX = i;
                else if (cellValue.Contains("累计") && cellValue.Contains("y"))
                    mapping.CumulativeY = i;
                else if (cellValue.Contains("累计") && cellValue.Contains("z"))
                    mapping.CumulativeZ = i;
                // 支持英文列名
                else if (cellValue.Contains("cumulative") && cellValue.Contains("x"))
                    mapping.CumulativeX = i;
                else if (cellValue.Contains("cumulative") && cellValue.Contains("y"))
                    mapping.CumulativeY = i;
                else if (cellValue.Contains("cumulative") && cellValue.Contains("z"))
                    mapping.CumulativeZ = i;
            }
            
            return mapping;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="sheet">工作表</param>
        private void SetColumnWidths(ISheet sheet)
        {
            var columnWidths = new[] { 15, 12, 15, 15, 15, 15, 15, 15, 15, 15, 15 };
            
            for (int i = 0; i < columnWidths.Length; i++)
            {
                sheet.SetColumnWidth(i, columnWidths[i] * 256); // 转换为Excel单位
            }
        }

        /// <summary>
        /// 按文件分组数据
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <returns>按文件分组的数据</returns>
        private Dictionary<string, List<PeriodData>> GroupDataByFile(List<MonitoringPoint> monitoringPoints)
        {
            var result = new Dictionary<string, List<PeriodData>>();

            foreach (var point in monitoringPoints)
            {
                foreach (var periodData in point.PeriodDataList)
                {
                    if (periodData.FileInfo?.OriginalFileName == null) continue;

                    var fileName = periodData.FileInfo.OriginalFileName;
                    if (!result.ContainsKey(fileName))
                    {
                        result[fileName] = new List<PeriodData>();
                    }
                    result[fileName].Add(periodData);
                }
            }

            return result;
        }

        /// <summary>
        /// 获取修正后的文件名
        /// </summary>
        /// <param name="originalFileName">原始文件名</param>
        /// <returns>修正后的文件名</returns>
        private string GetCorrectedFileName(string originalFileName)
        {
            var nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
            var extension = Path.GetExtension(originalFileName);
            // 保持原始扩展名，因为我们使用的是HSSFWorkbook
            return $"{nameWithoutExt}_修正后{extension}";
        }

        /// <summary>
        /// 生成修正报告
        /// </summary>
        /// <param name="correctionResult">修正结果</param>
        /// <param name="validationResults">验证结果</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>报告生成结果</returns>
        public ReportResult GenerateCorrectionReport(
            CorrectionResult correctionResult, 
            List<ValidationResult> validationResults, 
            string outputDirectory)
        {
            var result = new ReportResult();

            try
            {
                _logger.LogInformation("开始生成修正报告");

                // 生成详细报告
                var detailedReport = GenerateDetailedReport(correctionResult, validationResults);
                var detailedPath = Path.Combine(outputDirectory, "修正详细报告.txt");
                File.WriteAllText(detailedPath, detailedReport, System.Text.Encoding.UTF8);

                // 生成统计报告
                var statisticsReport = GenerateStatisticsReport(correctionResult, validationResults);
                var statisticsPath = Path.Combine(outputDirectory, "修正统计报告.txt");
                File.WriteAllText(statisticsPath, statisticsReport, System.Text.Encoding.UTF8);

                // 生成Excel格式的修正记录
                var excelReportPath = Path.Combine(outputDirectory, "修正记录.xlsx");
                GenerateExcelReport(correctionResult, excelReportPath);

                result.Status = ReportStatus.Success;
                result.Message = "报告生成成功";
                result.DetailedReportPath = detailedPath;
                result.StatisticsReportPath = statisticsPath;
                result.ExcelReportPath = excelReportPath;

                _logger.LogInformation("修正报告生成完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成修正报告时发生异常");
                result.Status = ReportStatus.Error;
                result.Message = $"生成报告时发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 生成详细报告
        /// </summary>
        /// <param name="correctionResult">修正结果</param>
        /// <param name="validationResults">验证结果</param>
        /// <returns>详细报告内容</returns>
        private string GenerateDetailedReport(CorrectionResult correctionResult, List<ValidationResult> validationResults)
        {
            var report = new System.Text.StringBuilder();
            
            report.AppendLine("=== DataFixter 数据修正详细报告 ===");
            report.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine();

            // 总体统计
            report.AppendLine("【总体统计】");
            report.AppendLine(correctionResult.GetSummary());
            report.AppendLine();

            // 验证结果统计
            var validCount = validationResults.Count(v => v.Status == ValidationStatus.Valid);
            var invalidCount = validationResults.Count(v => v.Status == ValidationStatus.Invalid);
            var needsAdjustmentCount = validationResults.Count(v => v.Status == ValidationStatus.NeedsAdjustment);
            
            report.AppendLine("【验证结果统计】");
            report.AppendLine($"验证通过: {validCount} 条");
            report.AppendLine($"验证失败: {invalidCount} 条");
            report.AppendLine($"需要修正: {needsAdjustmentCount} 条");
            report.AppendLine();

            // 修正详情
            report.AppendLine("【修正详情】");
            foreach (var pointResult in correctionResult.PointResults)
            {
                report.AppendLine($"监测点: {pointResult.PointName}");
                report.AppendLine($"  状态: {pointResult.Status}");
                report.AppendLine($"  消息: {pointResult.Message}");
                
                if (pointResult.Corrections.Any())
                {
                    report.AppendLine($"  修正记录:");
                    foreach (var correction in pointResult.Corrections)
                    {
                        report.AppendLine($"    {correction.Direction}方向 {correction.CorrectionType}: " +
                                        $"{correction.OriginalValue:F6} -> {correction.CorrectedValue:F6}");
                        report.AppendLine($"    原因: {correction.Reason}");
                    }
                }
                report.AppendLine();
            }

            return report.ToString();
        }

        /// <summary>
        /// 生成统计报告
        /// </summary>
        /// <param name="correctionResult">修正结果</param>
        /// <param name="validationResults">验证结果</param>
        /// <returns>统计报告内容</returns>
        private string GenerateStatisticsReport(CorrectionResult correctionResult, List<ValidationResult> validationResults)
        {
            var report = new System.Text.StringBuilder();
            
            report.AppendLine("=== DataFixter 数据修正统计报告 ===");
            report.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine();

            // 修正统计
            var statistics = correctionResult.AdjustmentRecords;
            var totalAdjustments = statistics.Count;
            var totalPoints = statistics.Select(r => r.PointName).Distinct().Count();
            var totalFiles = statistics.Select(r => r.FileName).Where(f => !string.IsNullOrEmpty(f)).Distinct().Count();

            report.AppendLine("【修正统计】");
            report.AppendLine($"总修正次数: {totalAdjustments}");
            report.AppendLine($"涉及监测点数: {totalPoints}");
            report.AppendLine($"涉及文件数: {totalFiles}");
            report.AppendLine();

            // 按调整类型统计
            var adjustmentTypeStats = statistics
                .GroupBy(r => r.AdjustmentType)
                .ToDictionary(g => g.Key.ToString(), g => g.Count());

            report.AppendLine("【按调整类型统计】");
            foreach (var stat in adjustmentTypeStats)
            {
                report.AppendLine($"{stat.Key}: {stat.Value} 次");
            }
            report.AppendLine();

            // 按数据方向统计
            var directionStats = statistics
                .GroupBy(r => r.DataDirection)
                .ToDictionary(g => g.Key.ToString(), g => g.Count());

            report.AppendLine("【按数据方向统计】");
            foreach (var stat in directionStats)
            {
                report.AppendLine($"{stat.Key}方向: {stat.Value} 次");
            }
            report.AppendLine();

            // 按点名统计
            var pointNameStats = statistics
                .GroupBy(r => r.PointName)
                .OrderByDescending(g => g.Count())
                .Take(20); // 只显示前20个

            report.AppendLine("【按点名统计（前20名）】");
            foreach (var stat in pointNameStats)
            {
                var pointName = stat.Key ?? "未知点名";
                var count = stat.Count();
                report.AppendLine($"{pointName}: {count} 次");
            }

            return report.ToString();
        }

        /// <summary>
        /// 生成Excel格式的修正记录
        /// </summary>
        /// <param name="correctionResult">修正结果</param>
        /// <param name="outputPath">输出路径</param>
        private void GenerateExcelReport(CorrectionResult correctionResult, string outputPath)
        {
            using var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("修正记录");

            // 设置列标题
            var headerRow = sheet.CreateRow(0);
            var headers = new[]
            {
                "点名", "文件名", "行号", "调整类型", "数据方向", 
                "原始值", "修正后值", "调整量", "修正原因", "时间"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                
                var style = sheet.Workbook.CreateCellStyle();
                var font = sheet.Workbook.CreateFont();
                font.IsBold = true;
                style.SetFont(font);
                cell.CellStyle = style;
            }

            // 写入数据
            var rowIndex = 1;
            foreach (var record in correctionResult.AdjustmentRecords)
            {
                var row = sheet.CreateRow(rowIndex);
                var colIndex = 0;

                row.CreateCell(colIndex++).SetCellValue(record.PointName ?? "");
                row.CreateCell(colIndex++).SetCellValue(record.FileName ?? "");
                row.CreateCell(colIndex++).SetCellValue(record.RowNumber);
                row.CreateCell(colIndex++).SetCellValue(record.AdjustmentType.ToString());
                row.CreateCell(colIndex++).SetCellValue(record.DataDirection.ToString());
                row.CreateCell(colIndex++).SetCellValue(record.OriginalValue);
                row.CreateCell(colIndex++).SetCellValue(record.AdjustedValue);
                row.CreateCell(colIndex++).SetCellValue(record.AdjustmentAmount);
                row.CreateCell(colIndex++).SetCellValue(record.Reason ?? "");
                row.CreateCell(colIndex++).SetCellValue(record.AdjustmentTime.ToString("yyyy-MM-dd HH:mm:ss"));

                rowIndex++;
            }

            // 设置列宽
            var columnWidths = new[] { 15, 20, 8, 15, 12, 15, 15, 15, 30, 20 };
            for (int i = 0; i < columnWidths.Length; i++)
            {
                sheet.SetColumnWidth(i, columnWidths[i] * 256);
            }

            // 保存文件
            using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
        }
    }

    /// <summary>
    /// 输出选项
    /// </summary>
    public class OutputOptions
    {
        /// <summary>
        /// 是否保留原始格式
        /// </summary>
        public bool PreserveOriginalFormat { get; set; } = true;

        /// <summary>
        /// 是否添加修正标记
        /// </summary>
        public bool AddCorrectionMarks { get; set; } = true;

        /// <summary>
        /// 输出文件编码
        /// </summary>
        public string OutputEncoding { get; set; } = "UTF-8";
    }

    // 输出相关模型类已移动到 Models/OutputModels.cs
}

/// <summary>
/// 列映射信息
/// </summary>
public class ColumnMapping
{
    /// <summary>
    /// 累计变化量X列索引
    /// </summary>
    public int CumulativeX { get; set; } = -1;
    
    /// <summary>
    /// 累计变化量Y列索引
    /// </summary>
    public int CumulativeY { get; set; } = -1;
    
    /// <summary>
    /// 累计变化量Z列索引
    /// </summary>
    public int CumulativeZ { get; set; } = -1;
}
