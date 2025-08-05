namespace WorkPartner.Models
{
    /// <summary>
    /// 处理结果模型
    /// </summary>
    public class ProcessingResult
    {
        /// <summary>
        /// 处理是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 处理开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 处理结束时间
        /// </summary>
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// 处理耗时
        /// </summary>
        public TimeSpan? Duration => EndTime?.Subtract(StartTime);

        /// <summary>
        /// 输入文件夹路径
        /// </summary>
        public string InputFolder { get; set; } = string.Empty;

        /// <summary>
        /// 输出文件夹路径
        /// </summary>
        public string OutputFolder { get; set; } = string.Empty;

        /// <summary>
        /// 处理的文件总数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 成功处理的文件数
        /// </summary>
        public int SuccessFiles { get; set; }

        /// <summary>
        /// 失败的文件数
        /// </summary>
        public int FailedFiles { get; set; }

        /// <summary>
        /// 跳过的文件数
        /// </summary>
        public int SkippedFiles { get; set; }

        /// <summary>
        /// 补充的文件数
        /// </summary>
        public int SupplementedFiles { get; set; }

        /// <summary>
        /// 补充的数据点总数
        /// </summary>
        public int TotalSupplementedDataPoints { get; set; }

        /// <summary>
        /// 错误信息列表
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();

        /// <summary>
        /// 警告信息列表
        /// </summary>
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>
        /// 处理详情
        /// </summary>
        public List<FileProcessingDetail> FileDetails { get; set; } = new List<FileProcessingDetail>();

        /// <summary>
        /// 内存使用情况（MB）
        /// </summary>
        public double MemoryUsageMB { get; set; }

        /// <summary>
        /// 处理成功率
        /// </summary>
        public double SuccessRate => TotalFiles > 0 ? (double)SuccessFiles / TotalFiles * 100 : 0;

        /// <summary>
        /// 失败率
        /// </summary>
        public double FailureRate => TotalFiles > 0 ? (double)FailedFiles / TotalFiles * 100 : 0;

        /// <summary>
        /// 跳过率
        /// </summary>
        public double SkipRate => TotalFiles > 0 ? (double)SkippedFiles / TotalFiles * 100 : 0;

        /// <summary>
        /// 添加错误信息
        /// </summary>
        /// <param name="error">错误信息</param>
        public void AddError(string error)
        {
            Errors.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {error}");
        }

        /// <summary>
        /// 添加警告信息
        /// </summary>
        /// <param name="warning">警告信息</param>
        public void AddWarning(string warning)
        {
            Warnings.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {warning}");
        }

        /// <summary>
        /// 添加文件处理详情
        /// </summary>
        /// <param name="detail">处理详情</param>
        public void AddFileDetail(FileProcessingDetail detail)
        {
            FileDetails.Add(detail);
        }

        /// <summary>
        /// 完成处理
        /// </summary>
        public void Complete()
        {
            EndTime = DateTime.Now;
        }

        /// <summary>
        /// 获取处理摘要
        /// </summary>
        /// <returns>处理摘要字符串</returns>
        public string GetSummary()
        {
            return $"处理完成: {SuccessFiles}/{TotalFiles} 成功, {FailedFiles} 失败, {SkippedFiles} 跳过, 耗时: {Duration?.TotalSeconds:F2}秒";
        }

        public override string ToString()
        {
            return GetSummary();
        }
    }

    /// <summary>
    /// 文件处理详情
    /// </summary>
    public class FileProcessingDetail
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 处理状态
        /// </summary>
        public ProcessingStatus Status { get; set; }

        /// <summary>
        /// 处理时间
        /// </summary>
        public DateTime ProcessedTime { get; set; }

        /// <summary>
        /// 处理耗时（毫秒）
        /// </summary>
        public long ProcessingTimeMs { get; set; }

        /// <summary>
        /// 数据行数
        /// </summary>
        public int DataRowCount { get; set; }

        /// <summary>
        /// 补充的数据点数量
        /// </summary>
        public int SupplementedDataPoints { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// 输出文件路径
        /// </summary>
        public string? OutputFilePath { get; set; }

        public override string ToString()
        {
            return $"{FileName} - {Status} ({ProcessingTimeMs}ms)";
        }
    }

    /// <summary>
    /// 处理状态枚举
    /// </summary>
    public enum ProcessingStatus
    {
        /// <summary>
        /// 成功
        /// </summary>
        Success,

        /// <summary>
        /// 失败
        /// </summary>
        Failed,

        /// <summary>
        /// 跳过
        /// </summary>
        Skipped,

        /// <summary>
        /// 处理中
        /// </summary>
        Processing
    }
} 