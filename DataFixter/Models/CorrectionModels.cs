using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 修正选项
    /// </summary>
    public class CorrectionOptions
    {
        /// <summary>
        /// 累计值容差
        /// </summary>
        public double CumulativeTolerance { get; set; } = 1e-6;

        /// <summary>
        /// 最大本期变化量
        /// </summary>
        public double MaxCurrentPeriodValue { get; set; } = 1.0;

        /// <summary>
        /// 最大累计变化量
        /// </summary>
        public double MaxCumulativeValue { get; set; } = 4.0;

        /// <summary>
        /// 是否启用最小化修改策略
        /// </summary>
        public bool EnableMinimalModification { get; set; } = true;
    }

    /// <summary>
    /// 数据修正信息
    /// </summary>
    public class DataCorrection
    {
        /// <summary>
        /// 期数据
        /// </summary>
        public PeriodData PeriodData { get; set; } = null!;

        /// <summary>
        /// 数据方向
        /// </summary>
        public DataDirection Direction { get; set; }

        /// <summary>
        /// 修正类型
        /// </summary>
        public CorrectionType CorrectionType { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public double OriginalValue { get; set; }

        /// <summary>
        /// 修正后的值
        /// </summary>
        public double CorrectedValue { get; set; }

        /// <summary>
        /// 修正原因
        /// </summary>
        public string Reason { get; set; } = string.Empty;

        /// <summary>
        /// 额外数据，用于存储双重修正时的累计值等信息
        /// </summary>
        public Dictionary<string, object>? AdditionalData { get; set; }
    }

    /// <summary>
    /// 修正类型
    /// </summary>
    public enum CorrectionType
    {
        /// <summary>
        /// 不修正
        /// </summary>
        None,

        /// <summary>
        /// 本期变化量
        /// </summary>
        CurrentPeriodValue,

        /// <summary>
        /// 累计变化量
        /// </summary>
        CumulativeValue,

        /// <summary>
        /// 两者都修正
        /// </summary>
        Both
    }

    /// <summary>
    /// 修正状态
    /// </summary>
    public enum CorrectionStatus
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
    /// 监测点修正结果
    /// </summary>
    public class PointCorrectionResult
    {
        /// <summary>
        /// 点名
        /// </summary>
        public string PointName { get; set; } = string.Empty;

        /// <summary>
        /// 修正状态
        /// </summary>
        public CorrectionStatus Status { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 修正列表
        /// </summary>
        public List<DataCorrection> Corrections { get; set; } = new List<DataCorrection>();

        /// <summary>
        /// 修正的期次数
        /// </summary>
        public int CorrectedPeriods { get; set; }

        /// <summary>
        /// 修正的数据值数量
        /// </summary>
        public int CorrectedValues { get; set; }
    }

    /// <summary>
    /// 修正结果
    /// </summary>
    public class CorrectionResult
    {
        /// <summary>
        /// 状态
        /// </summary>
        public CorrectionStatus Status { get; set; } = CorrectionStatus.Success;

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 监测点修正结果列表
        /// </summary>
        public List<PointCorrectionResult> PointResults { get; set; } = new List<PointCorrectionResult>();

        /// <summary>
        /// 修正记录列表
        /// </summary>
        public List<AdjustmentRecord> AdjustmentRecords { get; set; } = new List<AdjustmentRecord>();

        /// <summary>
        /// 添加监测点修正结果
        /// </summary>
        /// <param name="result">修正结果</param>
        public void AddPointResult(PointCorrectionResult result)
        {
            PointResults.Add(result);
        }

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要</returns>
        public string GetSummary()
        {
            var totalPoints = PointResults.Count;
            var successPoints = PointResults.Count(r => r.Status == CorrectionStatus.Success);
            var errorPoints = PointResults.Count(r => r.Status == CorrectionStatus.Error);
            var skippedPoints = PointResults.Count(r => r.Status == CorrectionStatus.Skipped);
            var totalCorrections = PointResults.Sum(r => r.CorrectedValues);

            return $"修正统计: 总计{totalPoints}个监测点, 成功{successPoints}个, 失败{errorPoints}个, 跳过{skippedPoints}个, 修正{totalCorrections}个数据值";
        }
    }

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
}
