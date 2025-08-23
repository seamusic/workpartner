using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace DataExport.Models
{
    /// <summary>
    /// 导出模式配置
    /// </summary>
    public class ExportModeConfig
    {
        /// <summary>
        /// 导出模式类型
        /// </summary>
        [JsonPropertyName("Mode")]
        public ExportMode Mode { get; set; } = ExportMode.AllProjects;

        /// <summary>
        /// 单个项目导出配置（当Mode为SingleProject时使用）
        /// </summary>
        [JsonPropertyName("SingleProject")]
        public SingleProjectExportConfig? SingleProject { get; set; }

        /// <summary>
        /// 指定时间范围导出配置（当Mode为CustomTimeRange时使用）
        /// </summary>
        [JsonPropertyName("CustomTimeRange")]
        public CustomTimeRangeExportConfig? CustomTimeRange { get; set; }

        /// <summary>
        /// 批量导出配置（当Mode为BatchExport时使用）
        /// </summary>
        [JsonPropertyName("BatchExport")]
        public BatchExportConfig? BatchExport { get; set; }

        /// <summary>
        /// 增量导出配置（当Mode为IncrementalExport时使用）
        /// </summary>
        [JsonPropertyName("IncrementalExport")]
        public IncrementalExportConfig? IncrementalExport { get; set; }

        /// <summary>
        /// 是否自动合并Excel文件
        /// </summary>
        [JsonPropertyName("AutoMerge")]
        public bool AutoMerge { get; set; } = true;

        /// <summary>
        /// 导出优先级（数字越小优先级越高）
        /// </summary>
        [JsonPropertyName("Priority")]
        public int Priority { get; set; } = 1;

        /// <summary>
        /// 是否启用并行导出
        /// </summary>
        [JsonPropertyName("EnableParallel")]
        public bool EnableParallel { get; set; } = false;

        /// <summary>
        /// 最大并行数量
        /// </summary>
        [JsonPropertyName("MaxParallelCount")]
        public int MaxParallelCount { get; set; } = 3;

        /// <summary>
        /// 导出间隔（毫秒）
        /// </summary>
        [JsonPropertyName("ExportInterval")]
        public int ExportInterval { get; set; } = 1000;

        /// <summary>
        /// 重试次数
        /// </summary>
        [JsonPropertyName("RetryCount")]
        public int RetryCount { get; set; } = 3;

        /// <summary>
        /// 重试间隔（毫秒）
        /// </summary>
        [JsonPropertyName("RetryInterval")]
        public int RetryInterval { get; set; } = 5000;
    }

    /// <summary>
    /// 批量导出配置
    /// </summary>
    public class BatchExportConfig
    {
        /// <summary>
        /// 要导出的项目ID列表
        /// </summary>
        [JsonPropertyName("ProjectIds")]
        public List<string> ProjectIds { get; set; } = new List<string>();

        /// <summary>
        /// 要导出的数据类型代码列表
        /// </summary>
        [JsonPropertyName("DataTypeCodes")]
        public List<string> DataTypeCodes { get; set; } = new List<string>();

        /// <summary>
        /// 时间范围列表
        /// </summary>
        [JsonPropertyName("TimeRanges")]
        public List<TimeRange> TimeRanges { get; set; } = new List<TimeRange>();

        /// <summary>
        /// 批量大小
        /// </summary>
        [JsonPropertyName("BatchSize")]
        public int BatchSize { get; set; } = 10;

        /// <summary>
        /// 是否按项目分组
        /// </summary>
        [JsonPropertyName("GroupByProject")]
        public bool GroupByProject { get; set; } = true;

        /// <summary>
        /// 是否按数据类型分组
        /// </summary>
        [JsonPropertyName("GroupByDataType")]
        public bool GroupByDataType { get; set; } = false;
    }

    /// <summary>
    /// 增量导出配置
    /// </summary>
    public class IncrementalExportConfig
    {
        /// <summary>
        /// 增量类型
        /// </summary>
        [JsonPropertyName("IncrementalType")]
        public IncrementalType IncrementalType { get; set; } = IncrementalType.Daily;

        /// <summary>
        /// 增量间隔（天）
        /// </summary>
        [JsonPropertyName("IncrementalInterval")]
        public int IncrementalInterval { get; set; } = 1;

        /// <summary>
        /// 开始日期
        /// </summary>
        [JsonPropertyName("StartDate")]
        public DateTime StartDate { get; set; } = DateTime.Today.AddDays(-30);

        /// <summary>
        /// 结束日期
        /// </summary>
        [JsonPropertyName("EndDate")]
        public DateTime EndDate { get; set; } = DateTime.Today;

        /// <summary>
        /// 是否包含今天
        /// </summary>
        [JsonPropertyName("IncludeToday")]
        public bool IncludeToday { get; set; } = false;

        /// <summary>
        /// 是否自动跳过已导出的日期
        /// </summary>
        [JsonPropertyName("SkipExportedDates")]
        public bool SkipExportedDates { get; set; } = true;

        /// <summary>
        /// 导出文件命名模式
        /// </summary>
        [JsonPropertyName("FileNamePattern")]
        public string FileNamePattern { get; set; } = "{ProjectName}_{DataType}_{Date:yyyyMMdd}";

        /// <summary>
        /// 是否启用增量合并
        /// </summary>
        [JsonPropertyName("EnableIncrementalMerge")]
        public bool EnableIncrementalMerge { get; set; } = true;

        /// <summary>
        /// 增量合并策略
        /// </summary>
        [JsonPropertyName("MergeStrategy")]
        public IncrementalMergeStrategy MergeStrategy { get; set; } = IncrementalMergeStrategy.Append;

        /// <summary>
        /// 增量合并文件大小限制（MB）
        /// </summary>
        [JsonPropertyName("MergeFileSizeLimit")]
        public int MergeFileSizeLimit { get; set; } = 100;

        /// <summary>
        /// 增量合并行数限制
        /// </summary>
        [JsonPropertyName("MergeRowLimit")]
        public int MergeRowLimit { get; set; } = 1000000;
    }

    /// <summary>
    /// 导出模式配置列表包装器
    /// </summary>
    public class ExportModeConfigList
    {
        /// <summary>
        /// 导出模式配置列表
        /// </summary>
        [JsonPropertyName("ExportModes")]
        public List<ExportModeConfig> ExportModes { get; set; } = new List<ExportModeConfig>();
    }

    /// <summary>
    /// 增量类型
    /// </summary>
    public enum IncrementalType
    {
        /// <summary>
        /// 按天增量
        /// </summary>
        Daily,
        /// <summary>
        /// 按周增量
        /// </summary>
        Weekly,
        /// <summary>
        /// 按月增量
        /// </summary>
        Monthly
    }

    /// <summary>
    /// 增量合并策略
    /// </summary>
    public enum IncrementalMergeStrategy
    {
        /// <summary>
        /// 追加模式
        /// </summary>
        Append,
        /// <summary>
        /// 覆盖模式
        /// </summary>
        Overwrite,
        /// <summary>
        /// 智能模式（根据时间戳决定）
        /// </summary>
        Smart
    }
}
