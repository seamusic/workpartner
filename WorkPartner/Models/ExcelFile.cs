namespace WorkPartner.Models
{
    /// <summary>
    /// Excel文件信息模型
    /// </summary>
    public class ExcelFile
    {
        /// <summary>
        /// 文件完整路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 文件名（不含路径）
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 文件日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 文件时间点（0、8、16）
        /// </summary>
        public int Hour { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 数据行列表
        /// </summary>
        public List<DataRow> DataRows { get; set; } = new List<DataRow>();

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// 最后修改时间
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// 文件是否被锁定
        /// </summary>
        public bool IsLocked { get; set; }

        /// <summary>
        /// 文件是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 是否已处理
        /// </summary>
        public bool IsProcessed { get; set; }

        /// <summary>
        /// 处理时间
        /// </summary>
        public DateTime? ProcessedTime { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 获取文件大小（KB）
        /// </summary>
        public double FileSizeKB => FileSize / 1024.0;

        /// <summary>
        /// 获取文件大小（MB）
        /// </summary>
        public double FileSizeMB => FileSize / (1024.0 * 1024.0);

        /// <summary>
        /// 获取格式化的日期字符串
        /// </summary>
        public string FormattedDate => Date.ToString("yyyy.M.d");

        /// <summary>
        /// 获取格式化的时间字符串
        /// </summary>
        public string FormattedHour => Hour.ToString("D2");

        /// <summary>
        /// 获取完整的文件标识
        /// </summary>
        public string FileIdentifier => $"{FormattedDate}-{FormattedHour}{ProjectName}";

        public override string ToString()
        {
            return $"{FileName} ({FormattedDate}-{FormattedHour})";
        }
    }
} 