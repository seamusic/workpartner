namespace WorkPartner.Models
{
    /// <summary>
    /// 补充文件信息
    /// </summary>
    public class SupplementFileInfo
    {
        /// <summary>
        /// 目标日期
        /// </summary>
        public DateTime TargetDate { get; set; }

        /// <summary>
        /// 目标时间点
        /// </summary>
        public int TargetHour { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 源文件
        /// </summary>
        public ExcelFile SourceFile { get; set; } = null!;

        /// <summary>
        /// 目标文件名
        /// </summary>
        public string TargetFileName { get; set; } = string.Empty;

        /// <summary>
        /// 是否需要微调
        /// </summary>
        public bool NeedsAdjustment { get; set; } = true;

        /// <summary>
        /// 微调参数
        /// </summary>
        public AdjustmentParameters AdjustmentParams { get; set; } = new AdjustmentParameters();
    }

    /// <summary>
    /// 微调参数
    /// </summary>
    public class AdjustmentParameters
    {
        /// <summary>
        /// 调整范围百分比
        /// </summary>
        public double AdjustmentRange { get; set; } = 0.05; // 5%

        /// <summary>
        /// 随机种子
        /// </summary>
        public int RandomSeed { get; set; } = 42;

        /// <summary>
        /// 最小调整值
        /// </summary>
        public double MinimumAdjustment { get; set; } = 0.001;

        /// <summary>
        /// 是否保持数据相关性
        /// </summary>
        public bool MaintainDataCorrelation { get; set; } = true;

        /// <summary>
        /// 相关性权重
        /// </summary>
        public double CorrelationWeight { get; set; } = 0.7;
    }
}
