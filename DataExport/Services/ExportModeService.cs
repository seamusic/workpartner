using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using DataExport.Services;

namespace DataExport.Services
{
    /// <summary>
    /// 导出模式服务
    /// </summary>
    public class ExportModeService
    {
        private readonly ILogger<ExportModeService> _logger;
        private readonly DataExportService _exportService;
        private readonly ExcelMergeService _mergeService;
        private readonly ExportConfig _config;

        public ExportModeService(ILogger<ExportModeService> logger, DataExportService exportService, ExcelMergeService mergeService, ExportConfig config)
        {
            _logger = logger;
            _exportService = exportService;
            _mergeService = mergeService;
            _config = config;
        }

        /// <summary>
        /// 执行单个项目导出
        /// </summary>
        public async Task<ExportModeResult> ExecuteSingleProjectExportAsync(SingleProjectExportConfig config)
        {
            _logger.LogInformation($"开始执行单个项目导出: {config.ProjectName}");

            try
            {
                // 查找项目配置
                var project = _config.Projects.FirstOrDefault(p => p.ProjectId == config.ProjectId);
                if (project == null)
                {
                    throw new InvalidOperationException($"未找到项目: {config.ProjectName} (ID: {config.ProjectId})");
                }

                // 确定要导出的数据类型
                var dataTypes = config.DataTypeCodes.Any() 
                    ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                    : project.DataTypes;

                if (!dataTypes.Any())
                {
                    throw new InvalidOperationException($"项目 {config.ProjectName} 没有找到指定的数据类型");
                }

                // 确定时间范围
                var timeRanges = config.TimeRange != null 
                    ? new List<TimeRange> { config.TimeRange }
                    : _config.ExportSettings.MonthlyExport.Months.Select(m => new TimeRange 
                        { StartTime = m.StartTime, EndTime = m.EndTime }).ToList();

                var result = new ExportModeResult
                {
                    Mode = ExportMode.SingleProject,
                    ProjectName = config.ProjectName,
                    Success = true,
                    StartTime = DateTime.Now
                };

                var exportResults = new List<ExportResult>();
                var totalExports = dataTypes.Count * timeRanges.Count;
                var currentExport = 0;

                _logger.LogInformation($"项目 {config.ProjectName} 将导出 {dataTypes.Count} 个数据类型，{timeRanges.Count} 个时间范围，总计 {totalExports} 个文件");

                // 循环每个数据类型
                foreach (var dataType in dataTypes)
                {
                    _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode})");

                    // 循环每个时间范围
                    foreach (var timeRange in timeRanges)
                    {
                        currentExport++;
                        var timeDesc = config.TimeRange != null ? "自定义时间范围" : $"时间范围 {currentExport}";
                        _logger.LogInformation($"    导出 {timeDesc}: {timeRange.StartTime} 至 {timeRange.EndTime} ({currentExport}/{totalExports})");

                        try
                        {
                            var exportResult = await ExportSingleDataAsync(project, dataType, timeRange);
                            exportResults.Add(exportResult);

                            if (exportResult.Success)
                            {
                                _logger.LogInformation($"    ✓ 导出成功: {exportResult.FileName}");
                            }
                            else
                            {
                                _logger.LogError($"    ✗ 导出失败: {exportResult.ErrorMessage}");
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
                                DataName = dataType.DataName
                            });
                        }

                        // 添加延迟避免请求过于频繁
                        await Task.Delay(1000);
                    }
                }

                result.ExportResults = exportResults;
                result.SuccessCount = exportResults.Count(r => r.Success);
                result.FailedCount = exportResults.Count(r => !r.Success);
                result.TotalCount = totalExports;

                // 如果启用自动合并，执行Excel合并
                if (_config.ExportSettings.AutoMerge)
                {
                    _logger.LogInformation("开始执行Excel自动合并...");
                    var mergeResults = await ExecuteAutoMergeAsync(project.ProjectName, dataTypes, exportResults);
                    result.MergeResults = mergeResults;
                    result.MergeSuccessCount = mergeResults.Count(r => r.Success);
                    result.MergeFailedCount = mergeResults.Count(r => !r.Success);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"单个项目导出完成: {config.ProjectName}, 成功: {result.SuccessCount}/{result.TotalCount}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"单个项目导出失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = ExportMode.SingleProject,
                    ProjectName = config.ProjectName,
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 执行指定时间范围导出
        /// </summary>
        public async Task<ExportModeResult> ExecuteCustomTimeRangeExportAsync(CustomTimeRangeExportConfig config)
        {
            _logger.LogInformation($"开始执行指定时间范围导出: {config.Description}");

            try
            {
                // 确定要导出的项目
                var projects = config.ProjectIds.Any()
                    ? _config.Projects.Where(p => config.ProjectIds.Contains(p.ProjectId)).ToList()
                    : _config.Projects;

                if (!projects.Any())
                {
                    throw new InvalidOperationException("没有找到指定的项目");
                }

                var result = new ExportModeResult
                {
                    Mode = ExportMode.CustomTimeRange,
                    Description = config.Description,
                    Success = true,
                    StartTime = DateTime.Now
                };

                var exportResults = new List<ExportResult>();
                var totalExports = 0;
                var currentExport = 0;

                // 计算总导出数量
                foreach (var project in projects)
                {
                    var dataTypes = config.DataTypeCodes.Any()
                        ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                        : project.DataTypes;
                    totalExports += dataTypes.Count;
                }

                _logger.LogInformation($"将导出 {projects.Count} 个项目，{totalExports} 个数据类型，时间范围: {config.TimeRange.StartTime} 至 {config.TimeRange.EndTime}");

                // 循环每个项目
                foreach (var project in projects)
                {
                    _logger.LogInformation($"开始处理项目: {project.ProjectName} (ID: {project.ProjectId})");

                    // 确定要导出的数据类型
                    var dataTypes = config.DataTypeCodes.Any()
                        ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                        : project.DataTypes;

                    // 循环每个数据类型
                    foreach (var dataType in dataTypes)
                    {
                        currentExport++;
                        _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode}) ({currentExport}/{totalExports})");

                        try
                        {
                            var exportResult = await ExportSingleDataAsync(project, dataType, config.TimeRange);
                            exportResults.Add(exportResult);

                            if (exportResult.Success)
                            {
                                _logger.LogInformation($"    ✓ 导出成功: {exportResult.FileName}");
                            }
                            else
                            {
                                _logger.LogError($"    ✗ 导出失败: {exportResult.ErrorMessage}");
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
                                DataName = dataType.DataName
                            });
                        }

                        // 添加延迟避免请求过于频繁
                        await Task.Delay(1000);
                    }
                }

                result.ExportResults = exportResults;
                result.SuccessCount = exportResults.Count(r => r.Success);
                result.FailedCount = exportResults.Count(r => !r.Success);
                result.TotalCount = totalExports;

                // 如果启用自动合并，执行Excel合并
                if (_config.ExportSettings.AutoMerge)
                {
                    _logger.LogInformation("开始执行Excel自动合并...");
                    var mergeResults = await ExecuteAutoMergeAsync("自定义时间范围", projects, exportResults);
                    result.MergeResults = mergeResults;
                    result.MergeSuccessCount = mergeResults.Count(r => r.Success);
                    result.MergeFailedCount = mergeResults.Count(r => !r.Success);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"指定时间范围导出完成: {config.Description}, 成功: {result.SuccessCount}/{result.TotalCount}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"指定时间范围导出失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = ExportMode.CustomTimeRange,
                    Description = config.Description,
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 执行导出所有项目（默认模式）
        /// </summary>
        public async Task<ExportModeResult> ExecuteAllProjectsExportAsync()
        {
            _logger.LogInformation("开始执行导出所有项目模式");

            try
            {
                if (!_config.Projects.Any())
                {
                    throw new InvalidOperationException("配置中没有找到任何项目");
                }

                var result = new ExportModeResult
                {
                    Mode = ExportMode.AllProjects,
                    Description = "导出所有项目",
                    Success = true,
                    StartTime = DateTime.Now
                };

                var exportResults = new List<ExportResult>();
                var totalExports = 0;
                var currentExport = 0;

                // 计算总导出数量
                foreach (var project in _config.Projects)
                {
                    totalExports += project.DataTypes.Count * _config.ExportSettings.MonthlyExport.Months.Count;
                }

                _logger.LogInformation($"将导出 {_config.Projects.Count} 个项目，总计 {totalExports} 个文件");

                // 循环每个项目
                foreach (var project in _config.Projects)
                {
                    _logger.LogInformation($"开始处理项目: {project.ProjectName} (ID: {project.ProjectId})");

                    // 循环每个数据类型
                    foreach (var dataType in project.DataTypes)
                    {
                        _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode})");

                        // 循环每个时间范围
                        foreach (var month in _config.ExportSettings.MonthlyExport.Months)
                        {
                            currentExport++;
                            _logger.LogInformation($"    导出 {month.Name}: {month.StartTime} 至 {month.EndTime} ({currentExport}/{totalExports})");

                            try
                            {
                                var timeRange = new TimeRange { StartTime = month.StartTime, EndTime = month.EndTime };
                                var exportResult = await ExportSingleDataAsync(project, dataType, timeRange);
                                exportResults.Add(exportResult);

                                if (exportResult.Success)
                                {
                                    _logger.LogInformation($"      ✓ 导出成功: {exportResult.FileName}");
                                }
                                else
                                {
                                    _logger.LogError($"      ✗ 导出失败: {exportResult.ErrorMessage}");
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, $"      导出异常: {ex.Message}");
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

                result.ExportResults = exportResults;
                result.SuccessCount = exportResults.Count(r => r.Success);
                result.FailedCount = exportResults.Count(r => !r.Success);
                result.TotalCount = totalExports;

                // 如果启用自动合并，执行Excel合并
                if (_config.ExportSettings.AutoMerge)
                {
                    _logger.LogInformation("开始执行Excel自动合并...");
                    var mergeResults = await ExecuteAutoMergeAsync("所有项目", _config.Projects, exportResults);
                    result.MergeResults = mergeResults;
                    result.MergeSuccessCount = mergeResults.Count(r => r.Success);
                    result.MergeFailedCount = mergeResults.Count(r => !r.Success);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"导出所有项目完成，成功: {result.SuccessCount}/{result.TotalCount}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"导出所有项目失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = ExportMode.AllProjects,
                    Description = "导出所有项目",
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 根据配置执行相应的导出模式
        /// </summary>
        public async Task<ExportModeResult> ExecuteExportModeAsync(ExportModeConfig modeConfig)
        {
            _logger.LogInformation($"开始执行导出模式: {modeConfig.Mode}");

            try
            {
                switch (modeConfig.Mode)
                {
                    case ExportMode.AllProjects:
                        return await ExecuteAllProjectsExportAsync();

                    case ExportMode.SingleProject:
                        if (modeConfig.SingleProject == null)
                        {
                            throw new InvalidOperationException("单个项目导出配置不能为空");
                        }
                        return await ExecuteSingleProjectExportAsync(modeConfig.SingleProject);

                    case ExportMode.CustomTimeRange:
                        if (modeConfig.CustomTimeRange == null)
                        {
                            throw new InvalidOperationException("自定义时间范围导出配置不能为空");
                        }
                        return await ExecuteCustomTimeRangeExportAsync(modeConfig.CustomTimeRange);

                    case ExportMode.BatchExport:
                        if (modeConfig.BatchExport == null)
                        {
                            throw new InvalidOperationException("批量导出配置不能为空");
                        }
                        return await ExecuteBatchExportAsync(modeConfig.BatchExport, modeConfig);

                    case ExportMode.IncrementalExport:
                        if (modeConfig.IncrementalExport == null)
                        {
                            throw new InvalidOperationException("增量导出配置不能为空");
                        }
                        return await ExecuteIncrementalExportAsync(modeConfig.IncrementalExport, modeConfig);

                    default:
                        throw new NotSupportedException($"不支持的导出模式: {modeConfig.Mode}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"执行导出模式失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = modeConfig.Mode,
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 执行批量导出
        /// </summary>
        public async Task<ExportModeResult> ExecuteBatchExportAsync(BatchExportConfig config, ExportModeConfig modeConfig)
        {
            _logger.LogInformation($"开始执行批量导出模式");

            try
            {
                var result = new ExportModeResult
                {
                    Mode = ExportMode.BatchExport,
                    Description = "批量导出",
                    Success = true,
                    StartTime = DateTime.Now
                };

                // 确定要导出的项目
                var projects = config.ProjectIds.Any()
                    ? _config.Projects.Where(p => config.ProjectIds.Contains(p.ProjectId)).ToList()
                    : _config.Projects;

                if (!projects.Any())
                {
                    throw new InvalidOperationException("没有找到指定的项目");
                }

                // 确定时间范围
                var timeRanges = config.TimeRanges.Any()
                    ? config.TimeRanges
                    : _config.ExportSettings.MonthlyExport.Months.Select(m => new TimeRange 
                        { StartTime = m.StartTime, EndTime = m.EndTime }).ToList();

                var exportResults = new List<ExportResult>();
                var totalExports = 0;

                // 计算总导出数量
                foreach (var project in projects)
                {
                    var dataTypes = config.DataTypeCodes.Any()
                        ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                        : project.DataTypes;
                    totalExports += dataTypes.Count * timeRanges.Count;
                }

                _logger.LogInformation($"将批量导出 {projects.Count} 个项目，总计 {totalExports} 个文件");

                // 根据配置决定是否并行导出
                if (modeConfig.EnableParallel)
                {
                    exportResults = await ExecuteParallelBatchExportAsync(projects, config, timeRanges, modeConfig, totalExports);
                }
                else
                {
                    exportResults = await ExecuteSequentialBatchExportAsync(projects, config, timeRanges, modeConfig, totalExports);
                }

                result.ExportResults = exportResults;
                result.SuccessCount = exportResults.Count(r => r.Success);
                result.FailedCount = exportResults.Count(r => !r.Success);
                result.TotalCount = totalExports;

                // 如果启用自动合并，执行Excel合并
                if (modeConfig.AutoMerge)
                {
                    _logger.LogInformation("开始执行Excel自动合并...");
                    var mergeResults = await ExecuteAutoMergeAsync("批量导出", projects, exportResults);
                    result.MergeResults = mergeResults;
                    result.MergeSuccessCount = mergeResults.Count(r => r.Success);
                    result.MergeFailedCount = mergeResults.Count(r => !r.Success);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"批量导出完成，成功: {result.SuccessCount}/{result.TotalCount}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"批量导出失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = ExportMode.BatchExport,
                    Description = "批量导出",
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 执行增量导出
        /// </summary>
        public async Task<ExportModeResult> ExecuteIncrementalExportAsync(IncrementalExportConfig config, ExportModeConfig modeConfig)
        {
            _logger.LogInformation($"开始执行增量导出模式，时间范围: {config.StartDate} 至 {config.EndDate}");

            try
            {
                var result = new ExportModeResult
                {
                    Mode = ExportMode.IncrementalExport,
                    Description = "增量导出",
                    Success = true,
                    StartTime = DateTime.Now
                };

                // 计算增量时间范围
                var startTime = config.StartDate;
                var endTime = config.EndDate;

                // 检查增量时间范围是否合理
                var totalDays = (endTime - startTime).TotalDays;
                if (totalDays > 365)
                {
                    endTime = startTime.AddDays(365);
                    _logger.LogWarning($"增量时间范围 {totalDays:F0} 天超过最大限制 365 天，调整为 {startTime:yyyy-MM-dd} 至 {endTime:yyyy-MM-dd}");
                }

                var timeRange = new TimeRange
                {
                    StartTime = startTime.ToString("yyyy-MM-dd HH:mm"),
                    EndTime = endTime.ToString("yyyy-MM-dd HH:mm")
                };

                var exportResults = new List<ExportResult>();
                var totalExports = 0;
                var currentExport = 0;

                // 计算总导出数量
                foreach (var project in _config.Projects)
                {
                    totalExports += project.DataTypes.Count;
                }

                _logger.LogInformation($"将增量导出 {_config.Projects.Count} 个项目，时间范围: {timeRange.StartTime} 至 {timeRange.EndTime}，总计 {totalExports} 个文件");

                // 循环每个项目
                foreach (var project in _config.Projects)
                {
                    _logger.LogInformation($"开始处理项目: {project.ProjectName} (ID: {project.ProjectId})");

                    // 循环每个数据类型
                    foreach (var dataType in project.DataTypes)
                    {
                        currentExport++;
                        _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode}) ({currentExport}/{totalExports})");

                        try
                        {
                            var exportResult = await ExportSingleDataAsync(project, dataType, timeRange);
                            exportResults.Add(exportResult);

                            if (exportResult.Success)
                            {
                                _logger.LogInformation($"    ✓ 导出成功: {exportResult.FileName}");
                            }
                            else
                            {
                                _logger.LogError($"    ✗ 导出失败: {exportResult.ErrorMessage}");
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
                                DataName = dataType.DataName
                            });
                        }

                        // 添加延迟避免请求过于频繁
                        await Task.Delay(modeConfig.ExportInterval);
                    }
                }

                result.ExportResults = exportResults;
                result.SuccessCount = exportResults.Count(r => r.Success);
                result.FailedCount = exportResults.Count(r => !r.Success);
                result.TotalCount = totalExports;

                // 如果启用自动合并，执行Excel合并
                if (modeConfig.AutoMerge)
                {
                    _logger.LogInformation("开始执行Excel自动合并...");
                    var mergeResults = await ExecuteAutoMergeAsync("增量导出", _config.Projects, exportResults);
                    result.MergeResults = mergeResults;
                    result.MergeSuccessCount = mergeResults.Count(r => r.Success);
                    result.MergeFailedCount = mergeResults.Count(r => !r.Success);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"增量导出完成，成功: {result.SuccessCount}/{result.TotalCount}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"增量导出失败: {ex.Message}");
                return new ExportModeResult
                {
                    Mode = ExportMode.IncrementalExport,
                    Description = "增量导出",
                    Success = false,
                    ErrorMessage = ex.Message,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                };
            }
        }

        /// <summary>
        /// 执行并行批量导出
        /// </summary>
        private async Task<List<ExportResult>> ExecuteParallelBatchExportAsync(List<ProjectConfig> projects, BatchExportConfig config, List<TimeRange> timeRanges, ExportModeConfig modeConfig, int totalExports)
        {
            var exportResults = new List<ExportResult>();
            var semaphore = new SemaphoreSlim(modeConfig.MaxParallelCount, modeConfig.MaxParallelCount);

            var tasks = new List<Task>();

            foreach (var project in projects)
            {
                var dataTypes = config.DataTypeCodes.Any()
                    ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                    : project.DataTypes;

                foreach (var dataType in dataTypes)
                {
                    foreach (var timeRange in timeRanges)
                    {
                        var task = Task.Run(async () =>
                        {
                            await semaphore.WaitAsync();
                            try
                            {
                                var exportResult = await ExportSingleDataAsync(project, dataType, timeRange);
                                lock (exportResults)
                                {
                                    exportResults.Add(exportResult);
                                }
                            }
                            finally
                            {
                                semaphore.Release();
                            }
                        });

                        tasks.Add(task);
                    }
                }
            }

            await Task.WhenAll(tasks);
            return exportResults;
        }

        /// <summary>
        /// 执行顺序批量导出
        /// </summary>
        private async Task<List<ExportResult>> ExecuteSequentialBatchExportAsync(List<ProjectConfig> projects, BatchExportConfig config, List<TimeRange> timeRanges, ExportModeConfig modeConfig, int totalExports)
        {
            var exportResults = new List<ExportResult>();
            var currentExport = 0;

            foreach (var project in projects)
            {
                _logger.LogInformation($"开始处理项目: {project.ProjectName} (ID: {project.ProjectId})");

                var dataTypes = config.DataTypeCodes.Any()
                    ? project.DataTypes.Where(dt => config.DataTypeCodes.Contains(dt.DataCode)).ToList()
                    : project.DataTypes;

                foreach (var dataType in dataTypes)
                {
                    _logger.LogInformation($"  处理数据类型: {dataType.DataName} ({dataType.DataCode})");

                    foreach (var timeRange in timeRanges)
                    {
                        currentExport++;
                        _logger.LogInformation($"    导出时间范围: {timeRange.StartTime} 至 {timeRange.EndTime} ({currentExport}/{totalExports})");

                        try
                        {
                            var exportResult = await ExportSingleDataAsync(project, dataType, timeRange);
                            exportResults.Add(exportResult);

                            if (exportResult.Success)
                            {
                                _logger.LogInformation($"      ✓ 导出成功: {exportResult.FileName}");
                            }
                            else
                            {
                                _logger.LogError($"      ✗ 导出失败: {exportResult.ErrorMessage}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"      导出异常: {ex.Message}");
                            exportResults.Add(new ExportResult
                            {
                                Success = false,
                                ErrorMessage = ex.Message,
                                ProjectName = project.ProjectName,
                                DataName = dataType.DataName
                            });
                        }

                        // 添加延迟避免请求过于频繁
                        await Task.Delay(modeConfig.ExportInterval);
                    }
                }
            }

            return exportResults;
        }

        /// <summary>
        /// 导出单个数据
        /// </summary>
        private async Task<ExportResult> ExportSingleDataAsync(ProjectConfig project, DataTypeConfig dataType, TimeRange timeRange)
        {
            var parameters = new ExportParameters
            {
                ProjectId = project.ProjectId,
                ProjectName = project.ProjectName,
                DataCode = dataType.DataCode,
                DataName = dataType.DataName,
                StartTime = timeRange.StartTime,
                EndTime = timeRange.EndTime,
                WithDetail = _config.ExportSettings.WithDetail,
                PointCodes = _config.ExportSettings.PointCodes
            };

            return await _exportService.ExportDataAsync(parameters);
        }

        /// <summary>
        /// 执行自动合并（单个项目）
        /// </summary>
        private async Task<List<MergeResult>> ExecuteAutoMergeAsync(string projectName, List<DataTypeConfig> dataTypes, List<ExportResult> exportResults)
        {
            var mergeResults = new List<MergeResult>();

            foreach (var dataType in dataTypes)
            {
                var dataTypeResults = exportResults.Where(r => r.DataName == dataType.DataName && r.Success).ToList();
                if (dataTypeResults.Count > 1)
                {
                    try
                    {
                        // 为每个数据类型创建单独的合并组
                        var mergeResult = await _mergeService.MergeExcelFilesAsync(dataTypeResults);
                        mergeResults.AddRange(mergeResult);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"合并数据类型 {dataType.DataName} 失败: {ex.Message}");
                        mergeResults.Add(new MergeResult
                        {
                            Success = false,
                            ProjectName = projectName,
                            DataName = dataType.DataName,
                            ErrorMessage = ex.Message
                        });
                    }
                }
            }

            return mergeResults;
        }

        /// <summary>
        /// 执行自动合并（多个项目）
        /// </summary>
        private async Task<List<MergeResult>> ExecuteAutoMergeAsync(string description, List<ProjectConfig> projects, List<ExportResult> ExportResults)
        {
            var mergeResults = new List<MergeResult>();

            foreach (var project in projects)
            {
                var projectResults = ExportResults.Where(r => r.ProjectName == project.ProjectName && r.Success).ToList();
                var dataTypeGroups = projectResults.GroupBy(r => r.DataName).ToList();

                foreach (var dataTypeGroup in dataTypeGroups)
                {
                    if (dataTypeGroup.Count() > 1)
                    {
                        try
                        {
                            // 为每个数据类型创建单独的合并组
                            var mergeResult = await _mergeService.MergeExcelFilesAsync(dataTypeGroup.ToList());
                            mergeResults.AddRange(mergeResult);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"合并项目 {project.ProjectName} 数据类型 {dataTypeGroup.Key} 失败: {ex.Message}");
                            mergeResults.Add(new MergeResult
                            {
                                Success = false,
                                ProjectName = project.ProjectName,
                                DataName = dataTypeGroup.Key,
                                ErrorMessage = ex.Message
                            });
                        }
                    }
                }
            }

            return mergeResults;
        }
    }
}
