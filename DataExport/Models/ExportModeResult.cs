using System;
using System.Collections.Generic;
using System.Linq;
using DataExport.Services; // Added for .Any()

namespace DataExport.Models
{
    /// <summary>
    /// 导出模式执行结果
    /// </summary>
    public class ExportModeResult
    {
        /// <summary>
        /// 导出模式类型
        /// </summary>
        public ExportMode Mode { get; set; }

        /// <summary>
        /// 项目名称（单个项目导出时使用）
        /// </summary>
        public string? ProjectName { get; set; }

        /// <summary>
        /// 描述（自定义时间范围导出时使用）
        /// </summary>
        public string? Description { get; set; }

        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 总导出数量
        /// </summary>
        public int TotalCount { get; set; }

        /// <summary>
        /// 成功导出数量
        /// </summary>
        public int SuccessCount { get; set; }

        /// <summary>
        /// 失败导出数量
        /// </summary>
        public int FailedCount { get; set; }

        /// <summary>
        /// 导出结果列表
        /// </summary>
        public List<ExportResult> ExportResults { get; set; } = new List<ExportResult>();

        /// <summary>
        /// 合并结果列表
        /// </summary>
        public List<MergeResult> MergeResults { get; set; } = new List<MergeResult>();

        /// <summary>
        /// 合并成功数量
        /// </summary>
        public int MergeSuccessCount { get; set; }

        /// <summary>
        /// 合并失败数量
        /// </summary>
        public int MergeFailedCount { get; set; }

        /// <summary>
        /// 总耗时
        /// </summary>
        public TimeSpan Duration => EndTime - StartTime;

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要</returns>
        public string GetSummary()
        {
            var summary = $"导出模式: {Mode}";
            
            if (!string.IsNullOrEmpty(ProjectName))
                summary += $", 项目: {ProjectName}";
            
            if (!string.IsNullOrEmpty(Description))
                summary += $", 描述: {Description}";
            
            summary += $", 状态: {(Success ? "成功" : "失败")}";
            
            if (Success)
            {
                summary += $", 导出: {SuccessCount}/{TotalCount}";
                if (MergeResults.Any())
                {
                    summary += $", 合并: {MergeSuccessCount}/{MergeResults.Count}";
                }
                summary += $", 耗时: {Duration:hh\\:mm\\:ss}";
            }
            else
            {
                summary += $", 错误: {ErrorMessage}";
            }
            
            return summary;
        }
    }
}
