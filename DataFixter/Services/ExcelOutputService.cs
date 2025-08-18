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

                        if (processedFiles % 50 == 0)
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
            
            // 调试信息：输出列映射结果
            //_logger.LogInformation("列映射检测结果: X本期={CurrentPeriodX}, Y本期={CurrentPeriodY}, Z本期={CurrentPeriodZ}, X累计={CumulativeX}, Y累计={CumulativeY}, Z累计={CumulativeZ}", 
            //    columnMapping.CurrentPeriodX, columnMapping.CurrentPeriodY, columnMapping.CurrentPeriodZ,
            //    columnMapping.CumulativeX, columnMapping.CumulativeY, columnMapping.CumulativeZ);
            
            foreach (var data in fileData)
            {
                if (data.RowNumber > 0 && data.RowNumber <= sheet.LastRowNum + 1)
                {
                    var rowIndex = data.RowNumber - 1; // 转换为0基索引
                    var row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        // 修正本期变化量列
                        if (columnMapping.CurrentPeriodX >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CurrentPeriodX) ?? row.CreateCell(columnMapping.CurrentPeriodX);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CurrentPeriodX);
                            _logger.LogDebug("修正X本期变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CurrentPeriodX, originalValue, data.CurrentPeriodX);
                            correctedRows++;
                        }
                        if (columnMapping.CurrentPeriodY >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CurrentPeriodY) ?? row.CreateCell(columnMapping.CurrentPeriodY);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CurrentPeriodY);
                            _logger.LogDebug("修正Y本期变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CurrentPeriodY, originalValue, data.CurrentPeriodY);
                            correctedRows++;
                        }
                        if (columnMapping.CurrentPeriodZ >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CurrentPeriodZ) ?? row.CreateCell(columnMapping.CurrentPeriodZ);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CurrentPeriodZ);
                            _logger.LogDebug("修正Z本期变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CurrentPeriodZ, originalValue, data.CurrentPeriodZ);
                            correctedRows++;
                        }
                        
                        // 修正累计变化量列
                        if (columnMapping.CumulativeX >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeX) ?? row.CreateCell(columnMapping.CumulativeX);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CumulativeX);
                            _logger.LogDebug("修正X累计变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CumulativeX, originalValue, data.CumulativeX);
                            correctedRows++;
                        }
                        if (columnMapping.CumulativeY >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeY) ?? row.CreateCell(columnMapping.CumulativeY);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CumulativeY);
                            _logger.LogDebug("修正Y累计变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CumulativeY, originalValue, data.CumulativeY);
                            correctedRows++;
                        }
                        if (columnMapping.CumulativeZ >= 0)
                        {
                            var cell = row.GetCell(columnMapping.CumulativeZ) ?? row.CreateCell(columnMapping.CumulativeZ);
                            var originalValue = cell.NumericCellValue;
                            cell.SetCellValue(data.CumulativeZ);
                            _logger.LogDebug("修正Z累计变化量: 行{RowNumber}, 列{ColumnIndex}, {OriginalValue} -> {CorrectedValue}", 
                                data.RowNumber, columnMapping.CumulativeZ, originalValue, data.CumulativeZ);
                            correctedRows++;
                        }
                    }
                }
            }
            
            //_logger.LogInformation("数据修正完成: 修正了 {CorrectedRows} 个数据项", correctedRows);
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
            if (headerRow == null) 
            {
                _logger.LogWarning("未找到标题行（第4行）");
                return mapping;
            }
            
            //_logger.LogInformation("检测到标题行，列数: {ColumnCount}", headerRow.LastCellNum);
            
            // 存储所有X轴、Y轴、Z轴列的位置
            var xAxisColumns = new List<int>();
            var yAxisColumns = new List<int>();
            var zAxisColumns = new List<int>();
            
            // 遍历标题行，查找相关列
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell == null) continue;
                
                var cellValue = cell.StringCellValue?.ToLower() ?? "";
                var originalValue = cell.StringCellValue ?? "";
                //_logger.LogInformation("列{ColumnIndex}: 原始值='{OriginalValue}', 处理后='{CellValue}'", i, originalValue, cellValue);
                
                // 检测本期变化量列
                if (cellValue.Contains("本期") && cellValue.Contains("x"))
                {
                    mapping.CurrentPeriodX = i;
                    //_logger.LogInformation("检测到X本期变化量列: 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("本期") && cellValue.Contains("y"))
                {
                    mapping.CurrentPeriodY = i;
                    //_logger.LogInformation("检测到Y本期变化量列: 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("本期") && cellValue.Contains("z"))
                {
                    mapping.CurrentPeriodZ = i;
                    //_logger.LogInformation("检测到Z本期变化量列: 列{ColumnIndex}", i);
                }
                // 支持英文列名
                else if (cellValue.Contains("current") && cellValue.Contains("x"))
                {
                    mapping.CurrentPeriodX = i;
                    //_logger.LogInformation("检测到X本期变化量列(英文): 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("current") && cellValue.Contains("y"))
                {
                    mapping.CurrentPeriodY = i;
                    //_logger.LogInformation("检测到Y本期变化量列(英文): 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("current") && cellValue.Contains("z"))
                {
                    mapping.CurrentPeriodZ = i;
                    //_logger.LogInformation("检测到Z本期变化量列(英文): 列{ColumnIndex}", i);
                }
                // 检测累计变化量列
                else if (cellValue.Contains("累计") && cellValue.Contains("x"))
                {
                    mapping.CumulativeX = i;
                    //_logger.LogInformation("检测到X累计变化量列: 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("累计") && cellValue.Contains("y"))
                {
                    mapping.CumulativeY = i;
                    //_logger.LogInformation("检测到Y累计变化量列: 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("累计") && cellValue.Contains("z"))
                {
                    mapping.CumulativeZ = i;
                    //_logger.LogInformation("检测到Z累计变化量列: 列{ColumnIndex}", i);
                }
                // 支持英文列名
                else if (cellValue.Contains("cumulative") && cellValue.Contains("x"))
                {
                    mapping.CumulativeX = i;
                    //_logger.LogInformation("检测到X累计变化量列(英文): 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("cumulative") && cellValue.Contains("y"))
                {
                    mapping.CumulativeY = i;
                    //_logger.LogInformation("检测到Y累计变化量列(英文): 列{ColumnIndex}", i);
                }
                else if (cellValue.Contains("cumulative") && cellValue.Contains("z"))
                {
                    mapping.CumulativeZ = i;
                    //_logger.LogInformation("检测到Z累计变化量列(英文): 列{ColumnIndex}", i);
                }
                // 检测"X轴"、"Y轴"、"Z轴"格式的列
                else if (cellValue == "x轴")
                {
                    xAxisColumns.Add(i);
                    //_logger.LogInformation("检测到X轴列: 列{ColumnIndex}", i);
                }
                else if (cellValue == "y轴")
                {
                    yAxisColumns.Add(i);
                    //_logger.LogInformation("检测到Y轴列: 列{ColumnIndex}", i);
                }
                else if (cellValue == "z轴")
                {
                    zAxisColumns.Add(i);
                    //_logger.LogInformation("检测到Z轴列: 列{ColumnIndex}", i);
                }
            }
            
            // 智能列分组识别：基于列位置和数量推断列类型
            if (xAxisColumns.Count >= 2 && yAxisColumns.Count >= 2 && zAxisColumns.Count >= 2)
            {
                //_logger.LogInformation("检测到多组X轴、Y轴、Z轴列，进行智能分组识别");
                
                // 假设：第一组是本期变化量，第二组是累计变化量，第三组是日变化量
                if (xAxisColumns.Count >= 1)
                {
                    mapping.CurrentPeriodX = xAxisColumns[0];
                    //_logger.LogInformation("推断X本期变化量列: 列{ColumnIndex}", xAxisColumns[0]);
                }
                if (yAxisColumns.Count >= 1)
                {
                    mapping.CurrentPeriodY = yAxisColumns[0];
                    //_logger.LogInformation("推断Y本期变化量列: 列{ColumnIndex}", yAxisColumns[0]);
                }
                if (zAxisColumns.Count >= 1)
                {
                    mapping.CurrentPeriodZ = zAxisColumns[0];
                    //_logger.LogInformation("推断Z本期变化量列: 列{ColumnIndex}", zAxisColumns[0]);
                }
                
                if (xAxisColumns.Count >= 2)
                {
                    mapping.CumulativeX = xAxisColumns[1];
                    //_logger.LogInformation("推断X累计变化量列: 列{ColumnIndex}", xAxisColumns[1]);
                }
                if (yAxisColumns.Count >= 2)
                {
                    mapping.CumulativeY = yAxisColumns[1];
                    //_logger.LogInformation("推断Y累计变化量列: 列{ColumnIndex}", yAxisColumns[1]);
                }
                if (zAxisColumns.Count >= 2)
                {
                    mapping.CumulativeZ = zAxisColumns[1];
                    //_logger.LogInformation("推断Z累计变化量列: 列{ColumnIndex}", zAxisColumns[1]);
                }
            }
            
            //_logger.LogInformation("列映射检测完成: X本期={CurrentPeriodX}, Y本期={CurrentPeriodY}, Z本期={CurrentPeriodZ}, X累计={CumulativeX}, Y累计={CumulativeY}, Z累计={CumulativeZ}", 
            //    mapping.CurrentPeriodX, mapping.CurrentPeriodY, mapping.CurrentPeriodZ,
            //    mapping.CumulativeX, mapping.CumulativeY, mapping.CumulativeZ);
            
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
            return $"{nameWithoutExt}{extension}";
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
                 var excelReportPath = Path.Combine(outputDirectory, "修正记录.xls");
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
            var canAdjustmentCount = validationResults.Count(v => v.CanAdjustment);

            report.AppendLine("【验证结果统计】");
            report.AppendLine($"验证通过: {validCount} 条");
            report.AppendLine($"验证失败: {invalidCount} 条");
            report.AppendLine($"需要修正: {needsAdjustmentCount} 条");
            report.AppendLine($"可以修正: {canAdjustmentCount} 条");
            report.AppendLine();

                         // 修正详情
             report.AppendLine("【修正详情】");
             foreach (var pointResult in correctionResult.PointResults)
             {
                 report.AppendLine($"点名: {pointResult.PointName}");
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
             report.AppendLine($"涉及点名数: {totalPoints}");
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
            // 检查数据量，如果超过Excel限制则分批处理
            const int maxRowsPerSheet = 65000; // 留一些余量
            var allCorrections = correctionResult.PointResults
                .SelectMany(p => p.Corrections)
                .ToList();
            
            if (allCorrections.Count == 0)
            {
                _logger.LogWarning("没有修正记录需要生成Excel报告");
                return;
            }
            
            _logger.LogInformation("开始生成Excel修正报告，总修正记录数: {Count}", allCorrections.Count);
            
            // 如果数据量超过限制，使用多个工作表或文件
            if (allCorrections.Count > maxRowsPerSheet)
            {
                GenerateLargeExcelReport(allCorrections, outputPath, maxRowsPerSheet);
            }
            else
            {
                GenerateStandardExcelReport(allCorrections, outputPath);
            }
        }
        
        /// <summary>
        /// 生成标准Excel报告（数据量较小）
        /// </summary>
        private void GenerateStandardExcelReport(List<DataCorrection> corrections, string outputPath)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("修正记录");
            
            // 创建标题行
            var headers = new[] { "点名", "文件名", "行号", "修正类型", "数据方向", "原始值", "修正后值", "修正原因", "修正时间" };
            var headerRow = sheet.CreateRow(0);
            
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
            foreach (var correction in corrections)
            {
                var row = sheet.CreateRow(rowIndex);
                var colIndex = 0;

                row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.PointName ?? "");
                row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.FileInfo?.OriginalFileName ?? "");
                row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.RowNumber ?? 0);
                row.CreateCell(colIndex++).SetCellValue(correction.CorrectionType.ToString());
                row.CreateCell(colIndex++).SetCellValue(correction.Direction.ToString());
                row.CreateCell(colIndex++).SetCellValue(correction.OriginalValue);
                row.CreateCell(colIndex++).SetCellValue(correction.CorrectedValue);
                row.CreateCell(colIndex++).SetCellValue(correction.Reason ?? "");
                row.CreateCell(colIndex++).SetCellValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                rowIndex++;
            }

            // 设置列宽
            var columnWidths = new[] { 15, 20, 8, 15, 12, 15, 15, 30, 20 };
            for (int i = 0; i < columnWidths.Length; i++)
            {
                sheet.SetColumnWidth(i, columnWidths[i] * 256);
            }

            // 保存文件
            using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
            
            _logger.LogInformation("标准Excel报告生成完成，共 {Count} 条记录", corrections.Count);
        }
        
        /// <summary>
        /// 生成大型Excel报告（数据量超过限制时使用）
        /// </summary>
        private void GenerateLargeExcelReport(List<DataCorrection> allCorrections, string baseOutputPath, int maxRowsPerSheet)
        {
            var totalSheets = (int)Math.Ceiling((double)allCorrections.Count / maxRowsPerSheet);
            _logger.LogInformation("数据量较大，将生成 {TotalSheets} 个工作表", totalSheets);
            
            var workbook = new HSSFWorkbook();
            
            for (int sheetIndex = 0; sheetIndex < totalSheets; sheetIndex++)
            {
                var startIndex = sheetIndex * maxRowsPerSheet;
                var endIndex = Math.Min(startIndex + maxRowsPerSheet, allCorrections.Count);
                var sheetCorrections = allCorrections.Skip(startIndex).Take(maxRowsPerSheet).ToList();
                
                var sheetName = sheetIndex == 0 ? "修正记录" : $"修正记录_{sheetIndex + 1}";
                var sheet = workbook.CreateSheet(sheetName);
                
                // 创建标题行
                var headers = new[] { "点名", "文件名", "行号", "修正类型", "数据方向", "原始值", "修正后值", "修正原因", "修正时间" };
                var headerRow = sheet.CreateRow(0);
                
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
                foreach (var correction in sheetCorrections)
                {
                    var row = sheet.CreateRow(rowIndex);
                    var colIndex = 0;

                    row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.PointName ?? "");
                    row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.FileInfo?.OriginalFileName ?? "");
                    row.CreateCell(colIndex++).SetCellValue(correction.PeriodData?.RowNumber ?? 0);
                    row.CreateCell(colIndex++).SetCellValue(correction.CorrectionType.ToString());
                    row.CreateCell(colIndex++).SetCellValue(correction.Direction.ToString());
                    row.CreateCell(colIndex++).SetCellValue(correction.OriginalValue);
                    row.CreateCell(colIndex++).SetCellValue(correction.CorrectedValue);
                    row.CreateCell(colIndex++).SetCellValue(correction.Reason ?? "");
                    row.CreateCell(colIndex++).SetCellValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    rowIndex++;
                }

                // 设置列宽
                var columnWidths = new[] { 15, 20, 8, 15, 12, 15, 15, 30, 20 };
                for (int i = 0; i < columnWidths.Length; i++)
                {
                    sheet.SetColumnWidth(i, columnWidths[i] * 256);
                }
                
                _logger.LogInformation("工作表 {SheetName} 生成完成，包含 {Count} 条记录", sheetName, sheetCorrections.Count);
            }

            // 保存文件
            using var fileStream = new FileStream(baseOutputPath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
            
            _logger.LogInformation("大型Excel报告生成完成，共 {TotalSheets} 个工作表，总计 {TotalRecords} 条记录", 
                totalSheets, allCorrections.Count);
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
    /// 本期变化量X列索引
    /// </summary>
    public int CurrentPeriodX { get; set; } = -1;
    
    /// <summary>
    /// 本期变化量Y列索引
    /// </summary>
    public int CurrentPeriodY { get; set; } = -1;
    
    /// <summary>
    /// 本期变化量Z列索引
    /// </summary>
    public int CurrentPeriodZ { get; set; } = -1;
    
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
