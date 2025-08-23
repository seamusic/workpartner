namespace DataExport.Models
{
    /// <summary>
    /// Excel合并结果
    /// </summary>
    public class MergeResult
    {
        /// <summary>
        /// 合并是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 数据名称
        /// </summary>
        public string DataName { get; set; } = string.Empty;

        /// <summary>
        /// 合并后的文件路径
        /// </summary>
        public string MergedFilePath { get; set; } = string.Empty;

        /// <summary>
        /// 合并后的文件名
        /// </summary>
        public string MergedFileName { get; set; } = string.Empty;

        /// <summary>
        /// 源文件路径列表
        /// </summary>
        public List<string> SourceFiles { get; set; } = new List<string>();

        /// <summary>
        /// 源文件数量
        /// </summary>
        public int SourceFileCount { get; set; }

        /// <summary>
        /// 数据行数（不包括标题行）
        /// </summary>
        public int DataRowCount { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 验证结果
        /// </summary>
        public MergeValidationResult? ValidationResult { get; set; }

        /// <summary>
        /// 合并开始时间
        /// </summary>
        public DateTime StartTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 合并结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 合并耗时
        /// </summary>
        public TimeSpan Duration => EndTime - StartTime;
    }

    /// <summary>
    /// 合并验证结果
    /// </summary>
    public class MergeValidationResult
    {
        /// <summary>
        /// 验证是否通过
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 验证消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; set; }

        /// <summary>
        /// 列数
        /// </summary>
        public int ColumnCount { get; set; }

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSizeBytes { get; set; }

        /// <summary>
        /// 文件大小（可读格式）
        /// </summary>
        public string FileSize => FormatFileSize(FileSizeBytes);

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }
}
