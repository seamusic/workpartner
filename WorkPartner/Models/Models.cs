using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkPartner.Services;
using WorkPartner.Utils;

namespace WorkPartner.Models
{
    /// <summary>
    /// 数据处理器配置
    /// </summary>
    public class DataProcessorConfig
    {
        public string CumulativeColumnPrefix { get; set; } = "G";
        public string ChangeColumnPrefix { get; set; } = "D";
        public double AdjustmentRange { get; set; } = 0.05; // 5%
        public int RandomSeed { get; set; } = 42;
        public double TimeFactorWeight { get; set; } = 1.0;
        public double MinimumAdjustment { get; set; } = 0.001;
        
        /// <summary>
        /// 第4、5、6列验证时的误差容忍度
        /// </summary>
        public double ColumnValidationTolerance { get; set; } = 0.01; // 1%

        /// <summary>
        /// 累计值调整阈值，当变化量超过此值时触发累计值调整
        /// </summary>
        public double CumulativeAdjustmentThreshold { get; set; } = 1.0; // 1.0

        /// <summary>
        /// 累计值调整时的误差容忍度
        /// </summary>
        public double CumulativeAdjustmentTolerance { get; set; } = 0.01; // 1%

        // 性能优化配置
        public bool EnableCaching { get; set; } = true;
        public int MaxCacheSize { get; set; } = 10000;
        public int CacheExpirationMinutes { get; set; } = 30;
        public bool EnableBatchProcessing { get; set; } = false;
        public int BatchSize { get; set; } = 100;
        public bool EnablePerformanceMonitoring { get; set; } = true;
        public bool EnableDetailedLogging { get; set; } = true;

        /// <summary>
        /// 获取默认配置
        /// </summary>
        public static DataProcessorConfig Default => new DataProcessorConfig
        {
            CumulativeColumnPrefix = "G",
            ChangeColumnPrefix = "D",
            AdjustmentRange = 0.05,
            RandomSeed = 42,
            TimeFactorWeight = 1.0,
            MinimumAdjustment = 0.001,
            ColumnValidationTolerance = 0.01,
            CumulativeAdjustmentThreshold = 1.0,
            CumulativeAdjustmentTolerance = 0.01,
            BatchSize = 50, // 减小批次大小以提高响应性
            EnableCaching = false, // 默认禁用缓存，避免性能下降
            CacheExpirationMinutes = 30,
            MaxCacheSize = 1000,
            EnableBatchProcessing = true,
            EnableDetailedLogging = false,
            EnablePerformanceMonitoring = true
        };
    }

    /// <summary>
    /// 连续缺失时间段信息
    /// </summary>
    public class MissingPeriod
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 缺失的小时数
        /// </summary>
        public int MissingHours { get; set; }

        /// <summary>
        /// 缺失的时间点列表
        /// </summary>
        public List<DateTime> MissingTimes { get; set; } = new List<DateTime>();

        /// <summary>
        /// 缺失的数据行名称
        /// </summary>
        public List<string> MissingDataRows { get; set; } = new List<string>();

        /// <summary>
        /// 前一个有效时间点
        /// </summary>
        public DateTime? PreviousValidTime { get; set; }

        /// <summary>
        /// 后一个有效时间点
        /// </summary>
        public DateTime? NextValidTime { get; set; }
    }

    /// <summary>
    /// 缺失数据点信息
    /// </summary>
    public class MissingDataPoint
    {
        /// <summary>
        /// 数据行名称
        /// </summary>
        public string DataRowName { get; set; } = string.Empty;

        /// <summary>
        /// 值索引
        /// </summary>
        public int ValueIndex { get; set; }

        /// <summary>
        /// 时间点
        /// </summary>
        public DateTime TimePoint { get; set; }

        /// <summary>
        /// 前一个有效值
        /// </summary>
        public double? PreviousValue { get; set; }

        /// <summary>
        /// 后一个有效值
        /// </summary>
        public double? NextValue { get; set; }

        /// <summary>
        /// 计算出的基础值（通常是前后值的平均值）
        /// </summary>
        public double? BaseValue { get; set; }
    }

    /// <summary>
    /// 性能监控信息
    /// </summary>
    public class PerformanceMetrics
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 处理文件数量
        /// </summary>
        public int ProcessedFiles { get; set; }

        /// <summary>
        /// 处理数据行数量
        /// </summary>
        public int ProcessedDataRows { get; set; }

        /// <summary>
        /// 累计变化量计算次数
        /// </summary>
        public int CumulativeCalculations { get; set; }

        /// <summary>
        /// 连续缺失数据处理次数
        /// </summary>
        public int MissingDataProcessings { get; set; }

        /// <summary>
        /// 补充文件微调次数
        /// </summary>
        public int SupplementFileAdjustments { get; set; }

        /// <summary>
        /// 缓存命中次数
        /// </summary>
        public int CacheHits { get; set; }

        /// <summary>
        /// 缓存未命中次数
        /// </summary>
        public int CacheMisses { get; set; }

        /// <summary>
        /// 总处理时间
        /// </summary>
        public TimeSpan TotalProcessingTime => EndTime - StartTime;

        /// <summary>
        /// 处理文件总数
        /// </summary>
        public int TotalFilesProcessed { get; set; }

        /// <summary>
        /// 处理数据行总数
        /// </summary>
        public int TotalDataRowsProcessed { get; set; }

        /// <summary>
        /// 处理值总数
        /// </summary>
        public int TotalValuesProcessed { get; set; }

        /// <summary>
        /// 总处理时间（毫秒）
        /// </summary>
        public double TotalProcessingTimeMs => TotalProcessingTime.TotalMilliseconds;

        /// <summary>
        /// 平均每文件处理时间
        /// </summary>
        public TimeSpan AverageTimePerFile => TotalFilesProcessed > 0 ? TimeSpan.FromMilliseconds(TotalProcessingTimeMs / TotalFilesProcessed) : TimeSpan.Zero;

        /// <summary>
        /// 平均每文件处理时间（毫秒）
        /// </summary>
        public double AverageTimePerFileMs => AverageTimePerFile.TotalMilliseconds;

        /// <summary>
        /// 缓存命中率
        /// </summary>
        public double CacheHitRate => (CacheHits + CacheMisses) > 0 ? (double)CacheHits / (CacheHits + CacheMisses) : 0;
    }

    /// <summary>
    /// 缓存项
    /// </summary>
    public class CacheItem<T>
    {
        /// <summary>
        /// 缓存的值
        /// </summary>
        public T Value { get; set; }

        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreatedTime { get; set; }

        /// <summary>
        /// 最后访问时间
        /// </summary>
        public DateTime LastAccessedTime { get; set; }

        /// <summary>
        /// 访问次数
        /// </summary>
        public int AccessCount { get; set; }

        /// <summary>
        /// 是否已过期
        /// </summary>
        public bool IsExpired(int expirationMinutes)
        {
            return (DateTime.Now - CreatedTime).TotalMinutes > expirationMinutes;
        }
    }

    /// <summary>
    /// 数据缓存管理器
    /// </summary>
    public class DataCache
    {
        private readonly Dictionary<string, CacheItem<object>> _cache = new();
        private readonly object _lockObject = new();
        private readonly int _maxSize;
        private readonly int _expirationMinutes;

        public DataCache(int maxSize = 10000, int expirationMinutes = 30)
        {
            _maxSize = maxSize;
            _expirationMinutes = expirationMinutes;
        }

        /// <summary>
        /// 获取缓存值
        /// </summary>
        public T? Get<T>(string key)
        {
            lock (_lockObject)
            {
                if (_cache.TryGetValue(key, out var item) && !item.IsExpired(_expirationMinutes))
                {
                    item.LastAccessedTime = DateTime.Now;
                    item.AccessCount++;
                    return (T)item.Value;
                }
                return default;
            }
        }

        /// <summary>
        /// 设置缓存值
        /// </summary>
        public void Set<T>(string key, T value)
        {
            lock (_lockObject)
            {
                // 清理过期项
                CleanupExpiredItems();

                // 如果缓存已满，移除最少访问的项
                if (_cache.Count >= _maxSize)
                {
                    var leastAccessed = _cache.OrderBy(kvp => kvp.Value.AccessCount).First();
                    _cache.Remove(leastAccessed.Key);
                }

                _cache[key] = new CacheItem<object>
                {
                    Value = value!,
                    CreatedTime = DateTime.Now,
                    LastAccessedTime = DateTime.Now,
                    AccessCount = 1
                };
            }
        }

        /// <summary>
        /// 清理过期项
        /// </summary>
        private void CleanupExpiredItems()
        {
            var expiredKeys = _cache.Where(kvp => kvp.Value.IsExpired(_expirationMinutes))
                                   .Select(kvp => kvp.Key)
                                   .ToList();

            foreach (var key in expiredKeys)
            {
                _cache.Remove(key);
            }
        }

        /// <summary>
        /// 清空缓存
        /// </summary>
        public void Clear()
        {
            lock (_lockObject)
            {
                _cache.Clear();
            }
        }

        /// <summary>
        /// 获取缓存统计信息
        /// </summary>
        public (int TotalItems, int ExpiredItems, int ValidItems) GetStats()
        {
            lock (_lockObject)
            {
                var expiredCount = _cache.Count(kvp => kvp.Value.IsExpired(_expirationMinutes));
                var validCount = _cache.Count - expiredCount;
                return (_cache.Count, expiredCount, validCount);
            }
        }
    }

    /// <summary>
    /// 完整性检查结果
    /// </summary>
    public class CompletenessCheckResult
    {
        /// <summary>
        /// 是否所有日期都完整
        /// </summary>
        public bool IsAllComplete { get; set; }

        /// <summary>
        /// 不完整的日期列表
        /// </summary>
        public List<DateTime> IncompleteDates { get; set; } = new List<DateTime>();

        /// <summary>
        /// 每个日期的完整性信息
        /// </summary>
        public List<DateCompleteness> DateCompleteness { get; set; } = new List<DateCompleteness>();
    }

    /// <summary>
    /// 日期完整性信息
    /// </summary>
    public class DateCompleteness
    {
        /// <summary>
        /// 日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 现有的时间点
        /// </summary>
        public List<int> ExistingHours { get; set; } = new List<int>();

        /// <summary>
        /// 缺失的时间点
        /// </summary>
        public List<int> MissingHours { get; set; } = new List<int>();

        /// <summary>
        /// 是否完整
        /// </summary>
        public bool IsComplete { get; set; }
    }



    /// <summary>
    /// 数据质量报告
    /// </summary>
    public class DataQualityReport
    {
        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 总行数
        /// </summary>
        public int TotalRows { get; set; }

        /// <summary>
        /// 有效行数
        /// </summary>
        public int ValidRows { get; set; }

        /// <summary>
        /// 缺失行数
        /// </summary>
        public int MissingRows { get; set; }

        /// <summary>
        /// 整体完整性百分比
        /// </summary>
        public double OverallCompleteness { get; set; }

        /// <summary>
        /// 每个文件的质量信息
        /// </summary>
        public List<FileQualityInfo> FileQuality { get; set; } = new List<FileQualityInfo>();
    }

    /// <summary>
    /// 文件质量信息
    /// </summary>
    public class FileQualityInfo
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 总行数
        /// </summary>
        public int TotalRows { get; set; }

        /// <summary>
        /// 有效行数
        /// </summary>
        public int ValidRows { get; set; }

        /// <summary>
        /// 缺失行数
        /// </summary>
        public int MissingRows { get; set; }

        /// <summary>
        /// 全部缺失的行数
        /// </summary>
        public int AllMissingRows { get; set; }

        /// <summary>
        /// 平均完整性百分比
        /// </summary>
        public double AverageCompleteness { get; set; }
    }
    
    /// <summary>
    /// 列差异详情
    /// </summary>
    public class ColumnDifference
    {
        public int ColumnIndex { get; set; }
        public double OriginalValue { get; set; }
        public double ProcessedValue { get; set; }
        public double Difference { get; set; }
        public bool IsSignificant { get; set; }
    }
    
    /// <summary>
    /// 行比较结果
    /// </summary>
    public class RowComparisonResult
    {
        public string RowName { get; set; } = string.Empty;
        public int OriginalValuesCount { get; set; }
        public int ProcessedValuesCount { get; set; }
        public int DifferencesCount { get; set; }
        public int SignificantDifferencesCount { get; set; }
        public int MissingProcessedValues { get; set; }
        public List<ColumnDifference> ColumnDifferences { get; set; } = new List<ColumnDifference>();
    }
    
    /// <summary>
    /// 文件比较结果
    /// </summary>
    public class FileComparisonResult
    {
        public string FileName { get; set; } = string.Empty;
        public DateTime OriginalDate { get; set; }
        public DateTime ProcessedDate { get; set; }
        public int OriginalValuesCount { get; set; }
        public int ProcessedValuesCount { get; set; }
        public int DifferencesCount { get; set; }
        public int SignificantDifferencesCount { get; set; }
        public List<string> MissingProcessedRows { get; set; } = new List<string>();
        public List<RowComparisonResult> RowComparisons { get; set; } = new List<RowComparisonResult>();
    }
    
    /// <summary>
    /// 整体比较结果
    /// </summary>
    public class ComparisonResult
    {
        public bool HasError { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public int TotalOriginalValues { get; set; }
        public int TotalProcessedValues { get; set; }
        public int TotalDifferences { get; set; }
        public int TotalSignificantDifferences { get; set; }
        public List<string> MissingProcessedFiles { get; set; } = new List<string>();
        public List<string> FailedComparisons { get; set; } = new List<string>();
        public List<FileComparisonResult> FileComparisons { get; set; } = new List<FileComparisonResult>();
    }
}
