namespace DataFixter.Models;

/// <summary>
/// 修正统计信息
/// </summary>
public class CorrectionStatistics
{
    /// <summary>
    /// 总修正次数
    /// </summary>
    public int TotalAdjustments { get; set; }

    /// <summary>
    /// 涉及监测点数
    /// </summary>
    public int TotalPoints { get; set; }

    /// <summary>
    /// 涉及文件数
    /// </summary>
    public int TotalFiles { get; set; }

    /// <summary>
    /// 按修正类型统计
    /// </summary>
    public Dictionary<string, int> CorrectionTypeStats { get; set; } = new Dictionary<string, int>();

    /// <summary>
    /// 按数据方向统计
    /// </summary>
    public Dictionary<string, int> DirectionStats { get; set; } = new Dictionary<string, int>();

    /// <summary>
    /// 按点名统计
    /// </summary>
    public Dictionary<string, int> PointNameStats { get; set; } = new Dictionary<string, int>();
}