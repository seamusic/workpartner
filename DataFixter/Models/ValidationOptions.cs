namespace DataFixter.Models;

/// <summary>
/// 验证选项
/// </summary>
public class ValidationOptions
{
    /// <summary>
    /// 累计变化量容差
    /// </summary>
    public double CumulativeTolerance { get; set; } = 2.0;

    /// <summary>
    /// 严重错误阈值
    /// </summary>
    public double CriticalThreshold { get; set; } = 5.0;

    /// <summary>
    /// 错误阈值
    /// </summary>
    public double ErrorThreshold { get; set; } = 3.0;

    /// <summary>
    /// 最小数值阈值
    /// </summary>
    public double MinValueThreshold { get; set; } = 0.01;

    /// <summary>
    /// 最大本期变化量
    /// </summary>
    public double MaxCurrentPeriodValue { get; set; } = 5.0;

    /// <summary>
    /// 里程容差
    /// </summary>
    public double MileageTolerance { get; set; } = 0.01;

    /// <summary>
    /// 最大时间间隔（天）
    /// </summary>
    public double MaxTimeInterval { get; set; } = 30.0;

    /// <summary>
    /// 数据验证最大处理时间（分钟）
    /// </summary>
    public int MaxProcessingTimeMinutes { get; set; } = 30;

    /// <summary>
    /// 批处理大小
    /// </summary>
    public int BatchSize { get; set; } = 50;

    /// <summary>
    /// 是否启用内存清理
    /// </summary>
    public bool EnableMemoryCleanup { get; set; } = true;

    /// <summary>
    /// 内存清理频率（每处理多少个监测点清理一次）
    /// </summary>
    public int MemoryCleanupFrequency { get; set; } = 200;

    /// <summary>
    /// 并行度
    /// </summary>
    public int MaxDegreeOfParallelism { get; set; } = 0; // 0 表示不限制，使用 CPU 核心数
}