using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    public class BatchExportService
    {
        private readonly ILogger<BatchExportService> _logger;
        private readonly DataExportService _exportService;
        private readonly ExcelMergeService _mergeService;
        private readonly ExportConfig _config;

        public BatchExportService(ILogger<BatchExportService> logger, DataExportService exportService, ExcelMergeService mergeService, ExportConfig config)
        {
            _logger = logger;
            _exportService = exportService;
            _mergeService = mergeService;
            _config = config;
        }

        /// <summary>
        /// 执行批量导出
        /// </summary>
        public async Task ExecuteBatchExportAsync()
        {
            await ExecuteBatchExportAsync(false);
        }

        /// <summary>
        /// 执行批量导出（带Excel合并选项）
        /// </summary>
        /// <param name="mergeExcelFiles">是否合并Excel文件</param>
        public async Task ExecuteBatchExportAsync(bool mergeExcelFiles)
        {
            _logger.LogInformation("开始执行批量导出任务");

            try
            {
                // 确保输出目录存在
                Directory.CreateDirectory(_config.ExportSettings.OutputDirectory);

                var totalExports = _config.Projects.Count * 
                                 _config.ExportSettings.MonthlyExport.Months.Count * 
                                 _config.Projects.Sum(p => p.DataTypes.Count);
                
                _logger.LogInformation($"总计需要导出 {totalExports} 个文件");

                var currentExport = 0;
                var exportResults = new List<ExportResult>();

                // 循环每个项目
                foreach (var project in _config.Projects)
                {
                    _logger.LogInformation($"开始处理项目: {project.ProjectName} (ID: {project.ProjectId})");

                    // 循环每个数据类型
                    foreach (var dataType in project.DataTypes)
                    {
                        _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode})");

                        // 循环每个月
                        foreach (var month in _config.ExportSettings.MonthlyExport.Months)
                        {
                            currentExport++;
                            _logger.LogInformation($"    导出 {month.Name}: {month.StartTime} 至 {month.EndTime} ({currentExport}/{totalExports})");

                            try
                            {
                                var result = await ExportSingleDataAsync(project, dataType, month);
                                exportResults.Add(result);

                                if (result.Success)
                                {
                                    _logger.LogInformation($"    ✓ 导出成功: {result.FileName}");
                                }
                                else
                                {
                                    _logger.LogError($"    ✗ 导出失败: {result.ErrorMessage}");
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, $"    导出异常: {ex.Message}");
                                exportResults.Add(new ExportResult
                                {
                                    Success = false,
                                    ErrorMessage = ex.Message,
                                    ProjectName = project.ProjectName,
                                    DataName = dataType.DataName,
                                    MonthName = month.Name
                                });
                            }

                            // 添加延迟避免请求过于频繁
                            await Task.Delay(1000);
                        }
                    }
                }

                // 如果启用了Excel合并功能
                List<MergeResult>? mergeResults = null;
                if (mergeExcelFiles)
                {
                    _logger.LogInformation("开始Excel文件合并");
                    mergeResults = await _mergeService.MergeExcelFilesAsync(exportResults);
                    
                    var successfulMerges = mergeResults.Count(m => m.Success);
                    _logger.LogInformation($"Excel合并完成: {successfulMerges}/{mergeResults.Count} 个组合并成功");
                }

                // 生成导出报告
                await GenerateExportReportAsync(exportResults, mergeResults);

                _logger.LogInformation("批量导出任务完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "批量导出任务执行失败");
                throw;
            }
        }

        /// <summary>
        /// 导出单个数据
        /// </summary>
        private async Task<ExportResult> ExportSingleDataAsync(ProjectConfig project, DataTypeConfig dataType, MonthConfig month)
        {
            var result = new ExportResult
            {
                ProjectName = project.ProjectName,
                DataName = dataType.DataName,
                MonthName = month.Name,
                StartTime = month.StartTime,
                EndTime = month.EndTime
            };

            try
            {
                // 构建导出参数
                var exportParams = new ExportParameters
                {
                    ProjectId = project.ProjectId,
                    ProjectName = project.ProjectName,
                    DataCode = dataType.DataCode,
                    DataName = dataType.DataName,
                    StartTime = month.StartTime,
                    EndTime = month.EndTime,
                    WithDetail = _config.ExportSettings.WithDetail,
                    PointCodes = _config.ExportSettings.PointCodes
                };

                // 执行导出
                var exportResult = await _exportService.ExportDataAsync(exportParams);
                
                if (exportResult.Success)
                {
                    result.Success = true;
                    result.FileName = exportResult.FileName;
                    result.FilePath = exportResult.FilePath;
                }
                else
                {
                    result.Success = false;
                    result.ErrorMessage = exportResult.ErrorMessage;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 生成导出报告
        /// </summary>
        private async Task GenerateExportReportAsync(List<ExportResult> results, List<MergeResult>? mergeResults = null)
        {
            var reportPath = Path.Combine(_config.ExportSettings.OutputDirectory, $"导出报告_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
            
            var report = new List<string>
            {
                new string('=', 80),
                "数据导出报告",
                new string('=', 80),
                $"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                $"总导出数量: {results.Count}",
                $"成功数量: {results.Count(r => r.Success)}",
                $"失败数量: {results.Count(r => !r.Success)}",
                ""
            };

            // 添加合并结果统计
            if (mergeResults != null && mergeResults.Any())
            {
                report.AddRange(new[]
                {
                    "Excel合并结果:",
                    $"总合并组数: {mergeResults.Count}",
                    $"合并成功: {mergeResults.Count(m => m.Success)}",
                    $"合并失败: {mergeResults.Count(m => !m.Success)}",
                    ""
                });
            }

            report.AddRange(new[]
            {
                "详细结果:",
                new string('-', 80)
            });

            // 按项目分组显示结果
            var projectGroups = results.GroupBy(r => r.ProjectName);
            foreach (var projectGroup in projectGroups)
            {
                report.Add($"项目: {projectGroup.Key}");
                
                var dataTypeGroups = projectGroup.GroupBy(r => r.DataName);
                foreach (var dataTypeGroup in dataTypeGroups)
                {
                    report.Add($"  数据类型: {dataTypeGroup.Key}");
                    
                    foreach (var result in dataTypeGroup)
                    {
                        var status = result.Success ? "✓" : "✗";
                        var info = result.Success ? result.FileName : result.ErrorMessage;
                        report.Add($"    {status} {result.MonthName}: {info}");
                    }
                }
                report.Add("");
            }

            // 失败项目汇总
            var failedResults = results.Where(r => !r.Success).ToList();
            if (failedResults.Any())
            {
                report.Add("失败项目汇总:");
                report.Add(new string('-', 40));
                foreach (var failed in failedResults)
                {
                    report.Add($"✗ {failed.ProjectName} - {failed.DataName} - {failed.MonthName}: {failed.ErrorMessage}");
                }
                report.Add("");
            }

            // 合并结果详情
            if (mergeResults != null && mergeResults.Any())
            {
                report.Add("Excel合并详情:");
                report.Add(new string('-', 40));
                
                foreach (var merge in mergeResults)
                {
                    var status = merge.Success ? "✓" : "✗";
                    var info = merge.Success 
                        ? $"{merge.MergedFileName} (源文件: {merge.SourceFileCount}, 数据行: {merge.DataRowCount})"
                        : merge.ErrorMessage;
                    
                    report.Add($"{status} {merge.ProjectName} - {merge.DataName}: {info}");
                    
                    if (merge.Success && merge.ValidationResult != null)
                    {
                        report.Add($"    验证: {merge.ValidationResult.Message}");
                    }
                }
                report.Add("");
                
                // 合并失败汇总
                var failedMerges = mergeResults.Where(m => !m.Success).ToList();
                if (failedMerges.Any())
                {
                    report.Add("合并失败汇总:");
                    report.Add(new string('-', 30));
                    foreach (var failed in failedMerges)
                    {
                        report.Add($"✗ {failed.ProjectName} - {failed.DataName}: {failed.ErrorMessage}");
                    }
                }
            }

            await File.WriteAllLinesAsync(reportPath, report);
            _logger.LogInformation($"导出报告已生成: {reportPath}");
        }
    }

    // ExportResult类已移动到DataExportService中

    /// <summary>
    /// 导出参数
    /// </summary>
    public class ExportParameters
    {
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public string DataCode { get; set; } = string.Empty;
        public string DataName { get; set; } = string.Empty;
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public int WithDetail { get; set; } = 1;
        public string PointCodes { get; set; } = string.Empty;
    }
}
