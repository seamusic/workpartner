using WorkPartner.Models;

namespace WorkPartner.Utils.Configuration
{
    /// <summary>
    /// 数据处理配置类
    /// </summary>
    public class DataProcessorConfiguration
    {
        /// <summary>
        /// 列映射配置
        /// </summary>
        public ColumnMappingConfiguration ColumnMapping { get; set; } = new();

        /// <summary>
        /// 验证设置
        /// </summary>
        public ValidationSettings ValidationSettings { get; set; } = new();

        /// <summary>
        /// 处理设置
        /// </summary>
        public ProcessingSettings ProcessingSettings { get; set; } = new();

        /// <summary>
        /// 日志设置
        /// </summary>
        public LoggingSettings LoggingSettings { get; set; } = new();

        /// <summary>
        /// 默认配置实例
        /// </summary>
        public static DataProcessorConfiguration Default => new()
        {
            ColumnMapping = ColumnMappingConfiguration.Default,
            ValidationSettings = ValidationSettings.Default,
            ProcessingSettings = ProcessingSettings.Default,
            LoggingSettings = LoggingSettings.Default
        };
    }

    /// <summary>
    /// 列映射配置
    /// </summary>
    public class ColumnMappingConfiguration
    {
        /// <summary>
        /// 累计变化量列前缀
        /// </summary>
        public string CumulativeColumnPrefix { get; set; } = "G";

        /// <summary>
        /// 变化量列前缀
        /// </summary>
        public string ChangeColumnPrefix { get; set; } = "D";

        /// <summary>
        /// 列索引映射
        /// </summary>
        public Dictionary<string, int> ColumnIndexMapping { get; set; } = new();

        /// <summary>
        /// 默认配置实例
        /// </summary>
        public static ColumnMappingConfiguration Default => new();
    }

    /// <summary>
    /// 验证设置
    /// </summary>
    public class ValidationSettings
    {
        /// <summary>
        /// 列验证容差
        /// </summary>
        public double ColumnValidationTolerance { get; set; } = 0.01;

        /// <summary>
        /// 累计调整阈值
        /// </summary>
        public double CumulativeAdjustmentThreshold { get; set; } = 0.1;

        /// <summary>
        /// 大值阈值
        /// </summary>
        public double LargeValueThreshold { get; set; } = 4.0;

        /// <summary>
        /// 默认配置实例
        /// </summary>
        public static ValidationSettings Default => new();
    }

    /// <summary>
    /// 处理设置
    /// </summary>
    public class ProcessingSettings
    {
        /// <summary>
        /// 是否启用缓存
        /// </summary>
        public bool EnableCaching { get; set; } = true;

        /// <summary>
        /// 是否启用批量处理
        /// </summary>
        public bool EnableBatchProcessing { get; set; } = true;

        /// <summary>
        /// 批次大小
        /// </summary>
        public int BatchSize { get; set; } = 100;

        /// <summary>
        /// 是否启用详细日志
        /// </summary>
        public bool EnableDetailedLogging { get; set; } = false;

        /// <summary>
        /// 是否启用性能监控
        /// </summary>
        public bool EnablePerformanceMonitoring { get; set; } = true;

        /// <summary>
        /// 最大缓存大小
        /// </summary>
        public int MaxCacheSize { get; set; } = 1000;

        /// <summary>
        /// 缓存过期时间（分钟）
        /// </summary>
        public int CacheExpirationMinutes { get; set; } = 30;

        /// <summary>
        /// 随机种子
        /// </summary>
        public int RandomSeed { get; set; } = 42;

        /// <summary>
        /// 默认配置实例
        /// </summary>
        public static ProcessingSettings Default => new();
    }

    /// <summary>
    /// 日志设置
    /// </summary>
    public class LoggingSettings
    {
        /// <summary>
        /// 是否启用控制台日志
        /// </summary>
        public bool EnableConsoleLogging { get; set; } = true;

        /// <summary>
        /// 是否启用文件日志
        /// </summary>
        public bool EnableFileLogging { get; set; } = true;

        /// <summary>
        /// 日志文件路径
        /// </summary>
        public string? LogFilePath { get; set; }

        /// <summary>
        /// 最小日志级别
        /// </summary>
        public string MinLogLevel { get; set; } = "Info";

        /// <summary>
        /// 默认配置实例
        /// </summary>
        public static LoggingSettings Default => new();
    }
}
