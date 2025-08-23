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
        /// 上次导出的时间
        /// </summary>
        [JsonPropertyName("LastExportTime")]
        public DateTime LastExportTime { get; set; } = DateTime.Now.AddDays(-1);

        /// <summary>
        /// 增量时间间隔（小时）
        /// </summary>
        [JsonPropertyName("IncrementHours")]
        public int IncrementHours { get; set; } = 24;

        /// <summary>
        /// 是否自动更新上次导出时间
        /// </summary>
        [JsonPropertyName("AutoUpdateLastExportTime")]
        public bool AutoUpdateLastExportTime { get; set; } = true;

        /// <summary>
        /// 最大增量时间范围（小时）
        /// </summary>
        [JsonPropertyName("MaxIncrementHours")]
        public int MaxIncrementHours { get; set; } = 168; // 7天

        /// <summary>
        /// 是否检查数据完整性
        /// </summary>
        [JsonPropertyName("CheckDataIntegrity")]
        public bool CheckDataIntegrity { get; set; } = true;
    }
}
