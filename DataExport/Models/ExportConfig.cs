using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace DataExport.Models
{
    /// <summary>
    /// 数据导出配置类
    /// </summary>
    public class ExportConfig
    {
        [JsonPropertyName("ApiSettings")]
        public ApiSettings ApiSettings { get; set; } = new();

        [JsonPropertyName("ExportSettings")]
        public ExportSettings ExportSettings { get; set; } = new();

        [JsonPropertyName("Projects")]
        public List<ProjectConfig> Projects { get; set; } = new();

        [JsonPropertyName("Logging")]
        public LoggingConfig Logging { get; set; } = new();
    }

    public class ApiSettings
    {
        [JsonPropertyName("BaseUrl")]
        public string BaseUrl { get; set; } = string.Empty;

        [JsonPropertyName("Endpoint")]
        public string Endpoint { get; set; } = string.Empty;

        [JsonPropertyName("Cookie")]
        public string Cookie { get; set; } = string.Empty;

        [JsonPropertyName("UserAgent")]
        public string UserAgent { get; set; } = string.Empty;

        [JsonPropertyName("Referer")]
        public string Referer { get; set; } = string.Empty;
    }

    public class ExportSettings
    {
        [JsonPropertyName("OutputDirectory")]
        public string OutputDirectory { get; set; } = "./exports";

        [JsonPropertyName("WithDetail")]
        public int WithDetail { get; set; } = 1;

        [JsonPropertyName("PointCodes")]
        public string PointCodes { get; set; } = string.Empty;

        [JsonPropertyName("DefaultTimeRange")]
        public TimeRange DefaultTimeRange { get; set; } = new();

        [JsonPropertyName("MonthlyExport")]
        public MonthlyExportSettings MonthlyExport { get; set; } = new();

        [JsonPropertyName("AutoMerge")]
        public bool AutoMerge { get; set; } = true;
    }

    public class TimeRange
    {
        [JsonPropertyName("StartTime")]
        public string StartTime { get; set; } = string.Empty;

        [JsonPropertyName("EndTime")]
        public string EndTime { get; set; } = string.Empty;
    }

    public class MonthlyExportSettings
    {
        [JsonPropertyName("Enabled")]
        public bool Enabled { get; set; } = true;

        [JsonPropertyName("Months")]
        public List<MonthConfig> Months { get; set; } = new();
    }

    public class MonthConfig
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("StartTime")]
        public string StartTime { get; set; } = string.Empty;

        [JsonPropertyName("EndTime")]
        public string EndTime { get; set; } = string.Empty;
    }

    public class ProjectConfig
    {
        [JsonPropertyName("ProjectId")]
        public string ProjectId { get; set; } = string.Empty;

        [JsonPropertyName("ProjectName")]
        public string ProjectName { get; set; } = string.Empty;

        [JsonPropertyName("DataTypes")]
        public List<DataTypeConfig> DataTypes { get; set; } = new();
    }

    public class DataTypeConfig
    {
        [JsonPropertyName("DataCode")]
        public string DataCode { get; set; } = string.Empty;

        [JsonPropertyName("DataName")]
        public string DataName { get; set; } = string.Empty;
    }

    public class LoggingConfig
    {
        [JsonPropertyName("LogLevel")]
        public LogLevelConfig LogLevel { get; set; } = new();
    }

    public class LogLevelConfig
    {
        [JsonPropertyName("Default")]
        public string Default { get; set; } = "Information";

        [JsonPropertyName("Microsoft")]
        public string Microsoft { get; set; } = "Warning";

        [JsonPropertyName("Microsoft.Hosting.Lifetime")]
        public string MicrosoftHostingLifetime { get; set; } = "Information";
    }



    /// <summary>
    /// 导出模式枚举
    /// </summary>
    public enum ExportMode
    {
        /// <summary>
        /// 导出所有项目（默认模式）
        /// </summary>
        AllProjects,

        /// <summary>
        /// 导出单个项目
        /// </summary>
        SingleProject,

        /// <summary>
        /// 导出指定时间范围
        /// </summary>
        CustomTimeRange,

        /// <summary>
        /// 批量导出
        /// </summary>
        BatchExport,

        /// <summary>
        /// 增量导出
        /// </summary>
        IncrementalExport,

        /// <summary>
        /// 定时导出
        /// </summary>
        ScheduledExport,

        /// <summary>
        /// 条件导出
        /// </summary>
        ConditionalExport
    }

    /// <summary>
    /// 单个项目导出配置
    /// </summary>
    public class SingleProjectExportConfig
    {
        /// <summary>
        /// 项目ID
        /// </summary>
        [JsonPropertyName("ProjectId")]
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>
        /// 项目名称
        /// </summary>
        [JsonPropertyName("ProjectName")]
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 要导出的数据类型列表（为空则导出所有类型）
        /// </summary>
        [JsonPropertyName("DataTypes")]
        public List<string> DataTypeCodes { get; set; } = new();

        /// <summary>
        /// 时间范围（为空则使用配置的月度时间）
        /// </summary>
        [JsonPropertyName("TimeRange")]
        public TimeRange? TimeRange { get; set; }
    }

    /// <summary>
    /// 指定时间范围导出配置
    /// </summary>
    public class CustomTimeRangeExportConfig
    {
        /// <summary>
        /// 要导出的项目ID列表（为空则导出所有项目）
        /// </summary>
        [JsonPropertyName("ProjectIds")]
        public List<string> ProjectIds { get; set; } = new();

        /// <summary>
        /// 要导出的数据类型代码列表（为空则导出所有类型）
        /// </summary>
        [JsonPropertyName("DataTypeCodes")]
        public List<string> DataTypeCodes { get; set; } = new();

        /// <summary>
        /// 自定义时间范围
        /// </summary>
        [JsonPropertyName("TimeRange")]
        public TimeRange TimeRange { get; set; } = new();

        /// <summary>
        /// 时间范围描述
        /// </summary>
        [JsonPropertyName("Description")]
        public string Description { get; set; } = string.Empty;
    }
}
