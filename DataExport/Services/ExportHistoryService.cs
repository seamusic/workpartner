using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// 导出历史记录服务
    /// </summary>
    public class ExportHistoryService
    {
        private readonly ILogger<ExportHistoryService> _logger;
        private readonly string _historyFilePath;
        private readonly List<ExportHistory> _histories;
        private readonly object _lockObject = new object();

        public ExportHistoryService(ILogger<ExportHistoryService> logger, string historyFilePath = "./export-history.json")
        {
            _logger = logger;
            _historyFilePath = historyFilePath;
            _histories = new List<ExportHistory>();
            LoadHistories();
        }

        /// <summary>
        /// 添加导出历史记录
        /// </summary>
        public async Task<ExportHistory> AddHistoryAsync(ExportHistory history)
        {
            try
            {
                lock (_lockObject)
                {
                    _histories.Add(history);
                }

                await SaveHistoriesAsync();
                _logger.LogInformation($"添加导出历史记录: {history.Id} - {history.Description}");
                return history;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"添加导出历史记录失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 更新导出历史记录
        /// </summary>
        public async Task<bool> UpdateHistoryAsync(string id, Action<ExportHistory> updateAction)
        {
            try
            {
                lock (_lockObject)
                {
                    var history = _histories.FirstOrDefault(h => h.Id == id);
                    if (history == null)
                    {
                        return false;
                    }

                    updateAction(history);
                }

                await SaveHistoriesAsync();
                _logger.LogInformation($"更新导出历史记录: {id}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"更新导出历史记录失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 获取导出历史记录
        /// </summary>
        public ExportHistory? GetHistory(string id)
        {
            lock (_lockObject)
            {
                return _histories.FirstOrDefault(h => h.Id == id);
            }
        }

        /// <summary>
        /// 查询导出历史记录
        /// </summary>
        public async Task<(List<ExportHistory> Histories, int TotalCount)> QueryHistoriesAsync(ExportHistoryQuery query)
        {
            try
            {
                lock (_lockObject)
                {
                    var queryable = _histories.AsQueryable();

                    // 应用过滤条件
                    if (query.StartTime.HasValue)
                    {
                        queryable = queryable.Where(h => h.StartTime >= query.StartTime.Value);
                    }

                    if (query.EndTime.HasValue)
                    {
                        queryable = queryable.Where(h => h.StartTime <= query.EndTime.Value);
                    }

                    if (query.ExportMode.HasValue)
                    {
                        queryable = queryable.Where(h => h.ExportMode == query.ExportMode.Value);
                    }

                    if (query.Status.HasValue)
                    {
                        queryable = queryable.Where(h => h.Status == query.Status.Value);
                    }

                    if (!string.IsNullOrEmpty(query.ProjectName))
                    {
                        queryable = queryable.Where(h => h.Projects.Any(p => p.Contains(query.ProjectName)));
                    }

                    if (!string.IsNullOrEmpty(query.DataType))
                    {
                        queryable = queryable.Where(h => h.DataTypes.Any(d => d.Contains(query.DataType)));
                    }

                    // 获取总数
                    var totalCount = queryable.Count();

                    // 应用排序
                    queryable = query.SortBy.ToLower() switch
                    {
                        "starttime" => query.IsDescending ? queryable.OrderByDescending(h => h.StartTime) : queryable.OrderBy(h => h.StartTime),
                        "endtime" => query.IsDescending ? queryable.OrderByDescending(h => h.EndTime) : queryable.OrderBy(h => h.EndTime),
                        "duration" => query.IsDescending ? queryable.OrderByDescending(h => h.Duration) : queryable.OrderBy(h => h.Duration),
                        "successrate" => query.IsDescending ? queryable.OrderByDescending(h => h.SuccessRate) : queryable.OrderBy(h => h.SuccessRate),
                        "totalcount" => query.IsDescending ? queryable.OrderByDescending(h => h.TotalCount) : queryable.OrderBy(h => h.TotalCount),
                        _ => query.IsDescending ? queryable.OrderByDescending(h => h.StartTime) : queryable.OrderBy(h => h.StartTime)
                    };

                    // 应用分页
                    var histories = queryable
                        .Skip((query.PageNumber - 1) * query.PageSize)
                        .Take(query.PageSize)
                        .ToList();

                    return (histories, totalCount);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"查询导出历史记录失败: {ex.Message}");
                throw;
            }

            return (new List<ExportHistory>(), 0);
        }

        /// <summary>
        /// 获取导出历史统计信息
        /// </summary>
        public async Task<ExportHistoryStatistics> GetStatisticsAsync()
        {
            try
            {
                lock (_lockObject)
                {
                    var statistics = new ExportHistoryStatistics
                    {
                        TotalRecords = _histories.Count,
                        SuccessRecords = _histories.Count(h => h.Status == ExportStatus.Completed),
                        FailedRecords = _histories.Count(h => h.Status == ExportStatus.Failed),
                        TotalExportedFiles = _histories.Sum(h => h.SuccessCount),
                        TotalMergedFiles = _histories.Sum(h => h.MergeSuccessCount),
                        AverageDuration = _histories.Where(h => h.IsCompleted).Any() 
                            ? _histories.Where(h => h.IsCompleted).Average(h => h.Duration) 
                            : 0
                    };

                    // 按模式统计
                    statistics.ByMode = _histories
                        .GroupBy(h => h.ExportMode)
                        .ToDictionary(g => g.Key, g => g.Count());

                    // 按状态统计
                    statistics.ByStatus = _histories
                        .GroupBy(h => h.Status)
                        .ToDictionary(g => g.Key, g => g.Count());

                    // 按项目统计
                    statistics.ByProject = _histories
                        .SelectMany(h => h.Projects)
                        .GroupBy(p => p)
                        .ToDictionary(g => g.Key, g => g.Count());

                    // 按数据类型统计
                    statistics.ByDataType = _histories
                        .SelectMany(h => h.DataTypes)
                        .GroupBy(d => d)
                        .ToDictionary(g => g.Key, g => g.Count());

                    return statistics;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"获取导出历史统计信息失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 删除导出历史记录
        /// </summary>
        public async Task<bool> DeleteHistoryAsync(string id)
        {
            try
            {
                lock (_lockObject)
                {
                    var history = _histories.FirstOrDefault(h => h.Id == id);
                    if (history == null)
                    {
                        return false;
                    }

                    _histories.Remove(history);
                }

                await SaveHistoriesAsync();
                _logger.LogInformation($"删除导出历史记录: {id}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"删除导出历史记录失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 清理过期的导出历史记录
        /// </summary>
        public async Task<int> CleanupExpiredHistoriesAsync(TimeSpan retentionPeriod)
        {
            try
            {
                var cutoffDate = DateTime.Now.Subtract(retentionPeriod);
                var expiredHistories = new List<ExportHistory>();

                lock (_lockObject)
                {
                    expiredHistories = _histories
                        .Where(h => h.StartTime < cutoffDate)
                        .ToList();

                    foreach (var history in expiredHistories)
                    {
                        _histories.Remove(history);
                    }
                }

                if (expiredHistories.Any())
                {
                    await SaveHistoriesAsync();
                    _logger.LogInformation($"清理过期导出历史记录: {expiredHistories.Count} 条，保留期限: {retentionPeriod.TotalDays} 天");
                }

                return expiredHistories.Count;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"清理过期导出历史记录失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 导出历史记录到文件
        /// </summary>
        public async Task<string> ExportHistoriesToFileAsync(ExportHistoryQuery query, string outputPath)
        {
            try
            {
                var (histories, _) = await QueryHistoriesAsync(query);
                var json = JsonSerializer.Serialize(histories, new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                var fileName = $"export-histories-{DateTime.Now:yyyyMMdd-HHmmss}.json";
                var fullPath = Path.Combine(outputPath, fileName);

                await File.WriteAllTextAsync(fullPath, json);
                _logger.LogInformation($"导出历史记录到文件: {fullPath}");

                return fullPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"导出历史记录到文件失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 加载历史记录
        /// </summary>
        private void LoadHistories()
        {
            try
            {
                if (File.Exists(_historyFilePath))
                {
                    var json = File.ReadAllText(_historyFilePath);
                    var histories = JsonSerializer.Deserialize<List<ExportHistory>>(json, new JsonSerializerOptions
                    {
                        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                    });

                    if (histories != null)
                    {
                        _histories.Clear();
                        _histories.AddRange(histories);
                        _logger.LogInformation($"加载导出历史记录: {_histories.Count} 条");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"加载导出历史记录失败: {ex.Message}，将创建新的历史记录文件");
            }
        }

        /// <summary>
        /// 保存历史记录
        /// </summary>
        private async Task SaveHistoriesAsync()
        {
            try
            {
                var directory = Path.GetDirectoryName(_historyFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = JsonSerializer.Serialize(_histories, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                await File.WriteAllTextAsync(_historyFilePath, json);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"保存导出历史记录失败: {ex.Message}");
                throw;
            }
        }
    }
}
