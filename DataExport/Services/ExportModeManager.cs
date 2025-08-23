using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// 导出模式管理器
    /// </summary>
    public class ExportModeManager
    {
        private readonly ILogger<ExportModeManager> _logger;
        private readonly ExportModeService _exportModeService;
        private readonly List<ExportModeConfig> _exportModes;
        private readonly GlobalExportSettings _globalSettings;

        public ExportModeManager(ILogger<ExportModeManager> logger, ExportModeService exportModeService, List<ExportModeConfig> exportModes, GlobalExportSettings globalSettings)
        {
            _logger = logger;
            _exportModeService = exportModeService;
            _exportModes = exportModes;
            _globalSettings = globalSettings;
        }

        /// <summary>
        /// 执行所有导出模式
        /// </summary>
        public async Task<List<ExportModeResult>> ExecuteAllModesAsync()
        {
            _logger.LogInformation("开始执行所有导出模式，总计 {Count} 个模式", _exportModes.Count);

            var results = new List<ExportModeResult>();
            var orderedModes = GetOrderedModes();

            if (_globalSettings.EnableModeSwitching)
            {
                results = await ExecuteModesWithSwitchingAsync(orderedModes);
            }
            else
            {
                results = await ExecuteModesSequentiallyAsync(orderedModes);
            }

            _logger.LogInformation("所有导出模式执行完成，成功: {SuccessCount}/{TotalCount}", 
                results.Count(r => r.Success), results.Count);

            return results;
        }

        /// <summary>
        /// 执行指定的导出模式
        /// </summary>
        public async Task<ExportModeResult> ExecuteModeAsync(string modeName)
        {
            var mode = _exportModes.FirstOrDefault(m => m.Mode.ToString() == modeName);
            if (mode == null)
            {
                throw new InvalidOperationException($"未找到导出模式: {modeName}");
            }

            _logger.LogInformation("开始执行导出模式: {ModeName}", modeName);
            return await _exportModeService.ExecuteExportModeAsync(mode);
        }

        /// <summary>
        /// 执行默认导出模式
        /// </summary>
        public async Task<ExportModeResult> ExecuteDefaultModeAsync()
        {
            var defaultMode = _exportModes.FirstOrDefault(m => m.Mode.ToString() == _globalSettings.DefaultExportMode)
                ?? _exportModes.FirstOrDefault(m => m.Mode == ExportMode.AllProjects)
                ?? _exportModes.First();

            _logger.LogInformation("开始执行默认导出模式: {Mode}", defaultMode.Mode);
            return await _exportModeService.ExecuteExportModeAsync(defaultMode);
        }

        /// <summary>
        /// 获取导出模式列表
        /// </summary>
        public List<ExportModeInfo> GetExportModes()
        {
            return _exportModes.Select(m => new ExportModeInfo
            {
                Mode = m.Mode,
                Description = GetModeDescription(m),
                Priority = m.Priority,
                IsEnabled = true,
                LastExecutionTime = null,
                ExecutionCount = 0
            }).ToList();
        }

        /// <summary>
        /// 验证导出模式配置
        /// </summary>
        public List<ValidationResult> ValidateModes()
        {
            var results = new List<ValidationResult>();

            foreach (var mode in _exportModes)
            {
                var result = ValidateMode(mode);
                if (!result.IsValid)
                {
                    results.Add(result);
                }
            }

            return results;
        }

        /// <summary>
        /// 获取排序后的导出模式
        /// </summary>
        private List<ExportModeConfig> GetOrderedModes()
        {
            return _globalSettings.ModeExecutionOrder.ToLower() switch
            {
                "priority" => _exportModes.OrderBy(m => m.Priority).ToList(),
                "name" => _exportModes.OrderBy(m => m.Mode.ToString()).ToList(),
                "custom" => _exportModes.ToList(), // 保持原有顺序
                _ => _exportModes.OrderBy(m => m.Priority).ToList()
            };
        }

        /// <summary>
        /// 执行模式切换（支持并发执行）
        /// </summary>
        private async Task<List<ExportModeResult>> ExecuteModesWithSwitchingAsync(List<ExportModeConfig> orderedModes)
        {
            var results = new List<ExportModeResult>();
            var semaphore = new SemaphoreSlim(_globalSettings.MaxConcurrentModes, _globalSettings.MaxConcurrentModes);

            var tasks = orderedModes.Select(async mode =>
            {
                await semaphore.WaitAsync();
                try
                {
                    var result = await _exportModeService.ExecuteExportModeAsync(mode);
                    lock (results)
                    {
                        results.Add(result);
                    }
                    return result;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await Task.WhenAll(tasks);
            return results.OrderBy(r => orderedModes.First(m => m.Mode == r.Mode).Priority).ToList();
        }

        /// <summary>
        /// 顺序执行导出模式
        /// </summary>
        private async Task<List<ExportModeResult>> ExecuteModesSequentiallyAsync(List<ExportModeConfig> orderedModes)
        {
            var results = new List<ExportModeResult>();

            foreach (var mode in orderedModes)
            {
                try
                {
                    var result = await _exportModeService.ExecuteExportModeAsync(mode);
                    results.Add(result);

                    if (!result.Success && _globalSettings.StopOnFirstFailure)
                    {
                        _logger.LogWarning("导出模式 {Mode} 执行失败，停止后续模式执行", mode.Mode);
                        break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "导出模式 {Mode} 执行异常", mode.Mode);
                    results.Add(new ExportModeResult
                    {
                        Mode = mode.Mode,
                        Success = false,
                        ErrorMessage = ex.Message,
                        StartTime = DateTime.Now,
                        EndTime = DateTime.Now
                    });

                    if (_globalSettings.StopOnFirstFailure)
                    {
                        break;
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// 获取模式描述
        /// </summary>
        private string GetModeDescription(ExportModeConfig mode)
        {
            return mode.Mode switch
            {
                ExportMode.AllProjects => "导出所有项目",
                ExportMode.SingleProject => $"导出单个项目: {mode.SingleProject?.ProjectName ?? "未配置"}",
                ExportMode.CustomTimeRange => $"自定义时间范围: {mode.CustomTimeRange?.Description ?? "未配置"}",
                ExportMode.BatchExport => "批量导出模式",
                ExportMode.IncrementalExport => "增量导出模式",
                _ => mode.Mode.ToString()
            };
        }

        /// <summary>
        /// 验证单个导出模式
        /// </summary>
        private ValidationResult ValidateMode(ExportModeConfig mode)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                switch (mode.Mode)
                {
                    case ExportMode.SingleProject:
                        if (mode.SingleProject == null)
                        {
                            result.IsValid = false;
                            result.Message = "单个项目导出配置不能为空";
                        }
                        break;

                    case ExportMode.CustomTimeRange:
                        if (mode.CustomTimeRange == null)
                        {
                            result.IsValid = false;
                            result.Message = "自定义时间范围导出配置不能为空";
                        }
                        break;

                    case ExportMode.BatchExport:
                        if (mode.BatchExport == null)
                        {
                            result.IsValid = false;
                            result.Message = "批量导出配置不能为空";
                        }
                        break;

                    case ExportMode.IncrementalExport:
                        if (mode.IncrementalExport == null)
                        {
                            result.IsValid = false;
                            result.Message = "增量导出配置不能为空";
                        }
                        break;
                }

                if (mode.MaxParallelCount <= 0)
                {
                    result.IsValid = false;
                    result.Message = "最大并行数量必须大于0";
                }

                if (mode.ExportInterval < 0)
                {
                    result.IsValid = false;
                    result.Message = "导出间隔不能为负数";
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.Message = $"验证异常: {ex.Message}";
            }

            return result;
        }
    }

    /// <summary>
    /// 导出模式信息
    /// </summary>
    public class ExportModeInfo
    {
        public ExportMode Mode { get; set; }
        public string Description { get; set; } = string.Empty;
        public int Priority { get; set; }
        public bool IsEnabled { get; set; }
        public DateTime? LastExecutionTime { get; set; }
        public int ExecutionCount { get; set; }
    }

    /// <summary>
    /// 全局导出设置
    /// </summary>
    public class GlobalExportSettings
    {
        public string DefaultExportMode { get; set; } = "AllProjects";
        public bool EnableModeSwitching { get; set; } = true;
        public string ModeExecutionOrder { get; set; } = "Priority";
        public int MaxConcurrentModes { get; set; } = 2;
        public int GlobalRetryCount { get; set; } = 3;
        public int GlobalRetryInterval { get; set; } = 10000;
        public bool EnableProgressTracking { get; set; } = true;
        public bool EnableResultPersistence { get; set; } = true;
        public string ResultStoragePath { get; set; } = "./export-results";
        public bool StopOnFirstFailure { get; set; } = false;
    }

    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
