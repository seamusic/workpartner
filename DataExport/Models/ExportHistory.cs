using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace DataExport.Models
{
    /// <summary>
    /// 导出历史记录
    /// </summary>
    public class ExportHistory
    {
        /// <summary>
        /// 记录ID
        /// </summary>
        [JsonPropertyName("id")]
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// 导出模式
        /// </summary>
        [JsonPropertyName("exportMode")]
        public ExportMode ExportMode { get; set; }

        /// <summary>
        /// 导出描述
        /// </summary>
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// 开始时间
        /// </summary>
        [JsonPropertyName("startTime")]
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        [JsonPropertyName("endTime")]
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// 执行状态
        /// </summary>
        [JsonPropertyName("status")]
        public ExportStatus Status { get; set; }

        /// <summary>
        /// 总导出数量
        /// </summary>
        [JsonPropertyName("totalCount")]
        public int TotalCount { get; set; }

        /// <summary>
        /// 成功数量
        /// </summary>
        [JsonPropertyName("successCount")]
        public int SuccessCount { get; set; }

        /// <summary>
        /// 失败数量
        /// </summary>
        [JsonPropertyName("failedCount")]
        public int FailedCount { get; set; }

        /// <summary>
        /// 合并成功数量
        /// </summary>
        [JsonPropertyName("mergeSuccessCount")]
        public int MergeSuccessCount { get; set; }

        /// <summary>
        /// 合并失败数量
        /// </summary>
        [JsonPropertyName("mergeFailedCount")]
        public int MergeFailedCount { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        [JsonPropertyName("errorMessage")]
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 导出的项目列表
        /// </summary>
        [JsonPropertyName("projects")]
        public List<string> Projects { get; set; } = new();

        /// <summary>
        /// 导出的数据类型列表
        /// </summary>
        [JsonPropertyName("dataTypes")]
        public List<string> DataTypes { get; set; } = new();

        /// <summary>
        /// 时间范围
        /// </summary>
        [JsonPropertyName("timeRange")]
        public TimeRange? TimeRange { get; set; }

        /// <summary>
        /// 输出文件路径
        /// </summary>
        [JsonPropertyName("outputFiles")]
        public List<string> OutputFiles { get; set; } = new();

        /// <summary>
        /// 合并文件路径
        /// </summary>
        [JsonPropertyName("mergedFiles")]
        public List<string> MergedFiles { get; set; } = new();

        /// <summary>
        /// 配置信息
        /// </summary>
        [JsonPropertyName("configuration")]
        public string Configuration { get; set; } = string.Empty;

        /// <summary>
        /// 执行耗时（毫秒）
        /// </summary>
        [JsonPropertyName("duration")]
        public long Duration => EndTime.HasValue ? (long)(EndTime.Value - StartTime).TotalMilliseconds : 0;

        /// <summary>
        /// 成功率
        /// </summary>
        [JsonPropertyName("successRate")]
        public double SuccessRate => TotalCount > 0 ? (double)SuccessCount / TotalCount * 100 : 0;

        /// <summary>
        /// 是否完成
        /// </summary>
        [JsonPropertyName("isCompleted")]
        public bool IsCompleted => Status == ExportStatus.Completed || Status == ExportStatus.Failed;

        /// <summary>
        /// 获取摘要信息
        /// </summary>
        public string GetSummary()
        {
            return $"{ExportMode} - {Status} - 成功: {SuccessCount}/{TotalCount} ({SuccessRate:F1}%) - 耗时: {Duration}ms";
        }
    }

    /// <summary>
    /// 导出状态枚举
    /// </summary>
    public enum ExportStatus
    {
        /// <summary>
        /// 等待中
        /// </summary>
        Pending,

        /// <summary>
        /// 执行中
        /// </summary>
        Running,

        /// <summary>
        /// 已完成
        /// </summary>
        Completed,

        /// <summary>
        /// 失败
        /// </summary>
        Failed,

        /// <summary>
        /// 已取消
        /// </summary>
        Cancelled
    }

    /// <summary>
    /// 导出历史查询条件
    /// </summary>
    public class ExportHistoryQuery
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        [JsonPropertyName("startTime")]
        public DateTime? StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        [JsonPropertyName("endTime")]
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// 导出模式
        /// </summary>
        [JsonPropertyName("exportMode")]
        public ExportMode? ExportMode { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        [JsonPropertyName("status")]
        public ExportStatus? Status { get; set; }

        /// <summary>
        /// 项目名称（模糊匹配）
        /// </summary>
        [JsonPropertyName("projectName")]
        public string? ProjectName { get; set; }

        /// <summary>
        /// 数据类型（模糊匹配）
        /// </summary>
        [JsonPropertyName("dataType")]
        public string? DataType { get; set; }

        /// <summary>
        /// 页码
        /// </summary>
        [JsonPropertyName("pageNumber")]
        public int PageNumber { get; set; } = 1;

        /// <summary>
        /// 每页大小
        /// </summary>
        [JsonPropertyName("pageSize")]
        public int PageSize { get; set; } = 20;

        /// <summary>
        /// 排序字段
        /// </summary>
        [JsonPropertyName("sortBy")]
        public string SortBy { get; set; } = "StartTime";

        /// <summary>
        /// 是否降序
        /// </summary>
        [JsonPropertyName("isDescending")]
        public bool IsDescending { get; set; } = true;
    }

    /// <summary>
    /// 导出历史统计信息
    /// </summary>
    public class ExportHistoryStatistics
    {
        /// <summary>
        /// 总记录数
        /// </summary>
        [JsonPropertyName("totalRecords")]
        public int TotalRecords { get; set; }

        /// <summary>
        /// 成功记录数
        /// </summary>
        [JsonPropertyName("successRecords")]
        public int SuccessRecords { get; set; }

        /// <summary>
        /// 失败记录数
        /// </summary>
        [JsonPropertyName("failedRecords")]
        public int FailedRecords { get; set; }

        /// <summary>
        /// 总导出文件数
        /// </summary>
        [JsonPropertyName("totalExportedFiles")]
        public int TotalExportedFiles { get; set; }

        /// <summary>
        /// 总合并文件数
        /// </summary>
        [JsonPropertyName("totalMergedFiles")]
        public int TotalMergedFiles { get; set; }

        /// <summary>
        /// 平均执行时间（毫秒）
        /// </summary>
        [JsonPropertyName("averageDuration")]
        public double AverageDuration { get; set; }

        /// <summary>
        /// 成功率
        /// </summary>
        [JsonPropertyName("successRate")]
        public double SuccessRate => TotalRecords > 0 ? (double)SuccessRecords / TotalRecords * 100 : 0;

        /// <summary>
        /// 按模式统计
        /// </summary>
        [JsonPropertyName("byMode")]
        public Dictionary<ExportMode, int> ByMode { get; set; } = new();

        /// <summary>
        /// 按状态统计
        /// </summary>
        [JsonPropertyName("byStatus")]
        public Dictionary<ExportStatus, int> ByStatus { get; set; } = new();

        /// <summary>
        /// 按项目统计
        /// </summary>
        [JsonPropertyName("byProject")]
        public Dictionary<string, int> ByProject { get; set; } = new();

        /// <summary>
        /// 按数据类型统计
        /// </summary>
        [JsonPropertyName("byDataType")]
        public Dictionary<string, int> ByDataType { get; set; } = new();
    }
}
