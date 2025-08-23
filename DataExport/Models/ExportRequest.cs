namespace DataExport.Models
{
    /// <summary>
    /// 导出请求参数
    /// </summary>
    public class ExportRequest
    {
        /// <summary>
        /// 项目ID
        /// </summary>
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>
        /// 数据类型代码
        /// </summary>
        public string DataCode { get; set; } = string.Empty;

        /// <summary>
        /// 开始时间
        /// </summary>
        public string StartTime { get; set; } = string.Empty;

        /// <summary>
        /// 结束时间
        /// </summary>
        public string EndTime { get; set; } = string.Empty;

        /// <summary>
        /// 是否包含详细信息
        /// </summary>
        public int WithDetail { get; set; } = 1;

        /// <summary>
        /// 监测点代码
        /// </summary>
        public string PointCodes { get; set; } = string.Empty;

        /// <summary>
        /// 项目名称（可选）
        /// </summary>
        public string? ProjectName { get; set; }

        /// <summary>
        /// 数据类型名称（可选）
        /// </summary>
        public string? DataName { get; set; }
    }
}
