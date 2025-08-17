using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 输出结果
    /// </summary>
    public class OutputResult
    {
        /// <summary>
        /// 状态
        /// </summary>
        public OutputStatus Status { get; set; } = OutputStatus.Success;

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 文件输出结果列表
        /// </summary>
        public List<FileOutputResult> FileResults { get; set; } = new List<FileOutputResult>();

        /// <summary>
        /// 添加文件输出结果
        /// </summary>
        /// <param name="result">文件输出结果</param>
        public void AddFileResult(FileOutputResult result)
        {
            FileResults.Add(result);
        }

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要</returns>
        public string GetSummary()
        {
            var totalFiles = FileResults.Count;
            var successFiles = FileResults.Count(r => r.Status == OutputStatus.Success);
            var errorFiles = FileResults.Count(r => r.Status == OutputStatus.Error);
            var totalRows = FileResults.Where(r => r.Status == OutputStatus.Success).Sum(r => r.RowCount);

            return $"输出统计: 总计{totalFiles}个文件, 成功{successFiles}个, 失败{errorFiles}个, 输出{totalRows}行数据";
        }
    }

    /// <summary>
    /// 文件输出结果
    /// </summary>
    public class FileOutputResult
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 状态
        /// </summary>
        public OutputStatus Status { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 输出路径
        /// </summary>
        public string OutputPath { get; set; } = string.Empty;

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; set; }
    }

    /// <summary>
    /// 输出状态
    /// </summary>
    public enum OutputStatus
    {
        /// <summary>
        /// 成功
        /// </summary>
        Success,

        /// <summary>
        /// 失败
        /// </summary>
        Error,

        /// <summary>
        /// 跳过
        /// </summary>
        Skipped
    }

    /// <summary>
    /// 报告结果
    /// </summary>
    public class ReportResult
    {
        /// <summary>
        /// 状态
        /// </summary>
        public ReportStatus Status { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 详细报告路径
        /// </summary>
        public string DetailedReportPath { get; set; } = string.Empty;

        /// <summary>
        /// 统计报告路径
        /// </summary>
        public string StatisticsReportPath { get; set; } = string.Empty;

        /// <summary>
        /// Excel报告路径
        /// </summary>
        public string ExcelReportPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// 报告状态
    /// </summary>
    public enum ReportStatus
    {
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
