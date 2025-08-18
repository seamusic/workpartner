using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 批量处理结果
    /// </summary>
    public class ProcessingResult
    {
        /// <summary>
        /// 处理状态
        /// </summary>
        public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;

        /// <summary>
        /// 处理消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 处理的文件数量
        /// </summary>
        public int ProcessedFiles { get; set; }

        /// <summary>
        /// 对比文件数量
        /// </summary>
        public int ComparisonFiles { get; set; }

        /// <summary>
        /// 监测点数量
        /// </summary>
        public int MonitoringPoints { get; set; }

        /// <summary>
        /// 验证结果列表
        /// </summary>
        public List<ValidationResult> ValidationResults { get; set; } = new List<ValidationResult>();

        /// <summary>
        /// 修正结果
        /// </summary>
        public CorrectionResult? CorrectionResult { get; set; }

        /// <summary>
        /// 输出结果
        /// </summary>
        public OutputResult? OutputResult { get; set; }

        /// <summary>
        /// 报告结果
        /// </summary>
        public ReportResult? ReportResult { get; set; }

        /// <summary>
        /// 处理开始时间
        /// </summary>
        public DateTime StartTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 处理结束时间
        /// </summary>
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// 获取处理摘要
        /// </summary>
        /// <returns>处理摘要</returns>
        public string GetSummary()
        {
            var duration = EndTime.HasValue ? EndTime.Value - StartTime : DateTime.Now - StartTime;
            var statusText = Status switch
            {
                ProcessingStatus.Success => "成功",
                ProcessingStatus.Error => "失败",
                ProcessingStatus.Pending => "进行中",
                _ => "未知"
            };

            var summary = $"处理状态: {statusText}\n";
            summary += $"处理时间: {duration.TotalSeconds:F1} 秒\n";
            summary += $"待处理文件: {ProcessedFiles} 个\n";
            summary += $"对比文件: {ComparisonFiles} 个\n";
            summary += $"监测点: {MonitoringPoints} 个\n";

            if (ValidationResults.Any())
            {
                var validCount = ValidationResults.Count(v => v.Status == ValidationStatus.Valid);
                var invalidCount = ValidationResults.Count(v => v.Status == ValidationStatus.Invalid);
                var needsAdjustmentCount = ValidationResults.Count(v => v.Status == ValidationStatus.NeedsAdjustment);
                var canAdjustment = ValidationResults.Count(v => v.Status == ValidationStatus.CanAdjustment);

                summary += $"验证结果: 通过 {validCount} 条, 失败 {invalidCount} 条, 需要修正 {needsAdjustmentCount} 条，可以修正 {canAdjustment}\n";
            }

            if (CorrectionResult != null)
            {
                summary += $"修正记录: {CorrectionResult.AdjustmentRecords.Count} 条\n";
            }

            if (OutputResult != null)
            {
                var successFiles = OutputResult.FileResults.Count(r => r.Status == OutputStatus.Success);
                var errorFiles = OutputResult.FileResults.Count(r => r.Status == OutputStatus.Error);
                summary += $"输出文件: 成功 {successFiles} 个, 失败 {errorFiles} 个\n";
            }

            if (!string.IsNullOrEmpty(Message))
            {
                summary += $"消息: {Message}";
            }

            return summary;
        }

        /// <summary>
        /// 获取详细统计信息
        /// </summary>
        /// <returns>详细统计信息</returns>
        public string GetDetailedStatistics()
        {
            var stats = new System.Text.StringBuilder();

            stats.AppendLine("=== 详细统计信息 ===");
            stats.AppendLine($"处理开始时间: {StartTime:yyyy-MM-dd HH:mm:ss}");
            if (EndTime.HasValue)
            {
                stats.AppendLine($"处理结束时间: {EndTime.Value:yyyy-MM-dd HH:mm:ss}");
                stats.AppendLine($"总耗时: {(EndTime.Value - StartTime).TotalSeconds:F1} 秒");
            }
            stats.AppendLine();

            // 文件统计
            stats.AppendLine("【文件统计】");
            stats.AppendLine($"待处理文件: {ProcessedFiles} 个");
            stats.AppendLine($"对比文件: {ComparisonFiles} 个");
            stats.AppendLine();

            // 数据统计
            stats.AppendLine("【数据统计】");
            stats.AppendLine($"监测点数量: {MonitoringPoints} 个");
            if (ValidationResults.Any())
            {
                stats.AppendLine($"验证记录总数: {ValidationResults.Count} 条");
            }
            if (CorrectionResult != null)
            {
                stats.AppendLine($"修正记录总数: {CorrectionResult.AdjustmentRecords.Count} 条");
            }
            stats.AppendLine();

            // 验证结果统计
            if (ValidationResults.Any())
            {
                stats.AppendLine("【验证结果统计】");
                var statusGroups = ValidationResults.GroupBy(v => v.Status);
                foreach (var group in statusGroups)
                {
                    stats.AppendLine($"{group.Key}: {group.Count()} 条");
                }
                stats.AppendLine();
            }

            // 修正结果统计
            if (CorrectionResult != null)
            {
                stats.AppendLine("【修正结果统计】");
                var adjustmentTypeGroups = CorrectionResult.AdjustmentRecords.GroupBy(r => r.AdjustmentType);
                foreach (var group in adjustmentTypeGroups)
                {
                    stats.AppendLine($"{group.Key}: {group.Count()} 条");
                }
                stats.AppendLine();
            }

            return stats.ToString();
        }
    }

    /// <summary>
    /// 处理状态
    /// </summary>
    public enum ProcessingStatus
    {
        /// <summary>
        /// 待处理
        /// </summary>
        Pending,

        /// <summary>
        /// 处理中
        /// </summary>
        Processing,

        /// <summary>
        /// 成功
        /// </summary>
        Success,

        /// <summary>
        /// 失败
        /// </summary>
        Error
    }
}
