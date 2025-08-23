using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace DataExport.Models
{
    /// <summary>
    /// 全局导出设置
    /// </summary>
    public class GlobalExportSettings
    {
        /// <summary>
        /// 默认导出模式
        /// </summary>
        [JsonPropertyName("DefaultExportMode")]
        public string DefaultExportMode { get; set; } = "AllProjects";

        /// <summary>
        /// 是否启用模式切换
        /// </summary>
        [JsonPropertyName("EnableModeSwitching")]
        public bool EnableModeSwitching { get; set; } = true;

        /// <summary>
        /// 模式执行顺序策略
        /// </summary>
        [JsonPropertyName("ModeExecutionOrder")]
        public string ModeExecutionOrder { get; set; } = "Priority";

        /// <summary>
        /// 最大并发模式数量
        /// </summary>
        [JsonPropertyName("MaxConcurrentModes")]
        public int MaxConcurrentModes { get; set; } = 2;

        /// <summary>
        /// 全局重试次数
        /// </summary>
        [JsonPropertyName("GlobalRetryCount")]
        public int GlobalRetryCount { get; set; } = 3;

        /// <summary>
        /// 全局重试间隔（毫秒）
        /// </summary>
        [JsonPropertyName("GlobalRetryInterval")]
        public int GlobalRetryInterval { get; set; } = 10000;

        /// <summary>
        /// 是否启用进度跟踪
        /// </summary>
        [JsonPropertyName("EnableProgressTracking")]
        public bool EnableProgressTracking { get; set; } = true;

        /// <summary>
        /// 是否启用结果持久化
        /// </summary>
        [JsonPropertyName("EnableResultPersistence")]
        public bool EnableResultPersistence { get; set; } = true;

        /// <summary>
        /// 结果存储路径
        /// </summary>
        [JsonPropertyName("ResultStoragePath")]
        public string ResultStoragePath { get; set; } = "./export-results";

        /// <summary>
        /// 是否在第一个失败时停止
        /// </summary>
        [JsonPropertyName("StopOnFirstFailure")]
        public bool StopOnFirstFailure { get; set; } = false;
    }
}
