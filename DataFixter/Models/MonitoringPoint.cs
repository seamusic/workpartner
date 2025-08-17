using System;
using System.Collections.Generic;
using System.Linq;

namespace DataFixter.Models
{
    /// <summary>
    /// 监测点数据模型类
    /// 包含点名、里程等基础信息，包含该点名的所有期数据列表
    /// </summary>
    public class MonitoringPoint
    {
        /// <summary>
        /// 点名（唯一标识）
        /// </summary>
        public string? PointName { get; set; }

        /// <summary>
        /// 里程
        /// </summary>
        public double Mileage { get; set; }

        /// <summary>
        /// 该点名的所有期数据列表（按时间排序）
        /// </summary>
        public List<PeriodData> PeriodDataList { get; set; } = new List<PeriodData>();

        /// <summary>
        /// 对比数据（来自对比目录）
        /// </summary>
        public PeriodData? ComparisonData { get; set; }

        /// <summary>
        /// 验证状态
        /// </summary>
        public ValidationStatus ValidationStatus { get; set; } = ValidationStatus.NotValidated;

        /// <summary>
        /// 验证失败的原因
        /// </summary>
        public List<string> ValidationErrors { get; set; } = new List<string>();

        /// <summary>
        /// 调整记录
        /// </summary>
        public List<AdjustmentRecord> Adjustments { get; set; } = new List<AdjustmentRecord>();

        /// <summary>
        /// 构造函数
        /// </summary>
        public MonitoringPoint()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="pointName">点名</param>
        /// <param name="mileage">里程</param>
        public MonitoringPoint(string pointName, double mileage)
        {
            PointName = pointName;
            Mileage = mileage;
        }

        /// <summary>
        /// 添加期数据
        /// </summary>
        /// <param name="periodData">期数据</param>
        public void AddPeriodData(PeriodData periodData)
        {
            if (periodData == null)
                throw new ArgumentNullException(nameof(periodData));

            if (string.IsNullOrEmpty(periodData.PointName))
                throw new ArgumentException("期数据的点名不能为空");

            if (!string.Equals(periodData.PointName, PointName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"期数据的点名 {periodData.PointName} 与监测点名 {PointName} 不匹配");

            PeriodDataList.Add(periodData);
            
            // 按时间排序
            SortPeriodDataByTime();
        }

        /// <summary>
        /// 按时间排序期数据
        /// </summary>
        public void SortPeriodDataByTime()
        {
            PeriodDataList = PeriodDataList
                .OrderBy(pd => pd.FileInfo)
                .ToList();
        }

        /// <summary>
        /// 获取指定时间的期数据
        /// </summary>
        /// <param name="dateTime">时间</param>
        /// <returns>期数据，如果不存在则返回null</returns>
        public PeriodData? GetPeriodDataByTime(DateTime dateTime)
        {
            return PeriodDataList.FirstOrDefault(pd => 
                pd.FileInfo?.Date == dateTime.Date && 
                pd.FileInfo?.Hour == dateTime.Hour);
        }

        /// <summary>
        /// 获取指定文件的期数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>期数据，如果不存在则返回null</returns>
        public PeriodData? GetPeriodDataByFile(string filePath)
        {
            return PeriodDataList.FirstOrDefault(pd => 
                string.Equals(pd.FileInfo?.FullPath, filePath, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 获取期数据数量
        /// </summary>
        public int PeriodDataCount => PeriodDataList.Count;

        /// <summary>
        /// 检查是否有期数据
        /// </summary>
        public bool HasPeriodData => PeriodDataList.Count > 0;

        /// <summary>
        /// 获取最早的时间
        /// </summary>
        public DateTime? EarliestTime => PeriodDataList.Count > 0 ? PeriodDataList.First().FileInfo?.FullDateTime : null;

        /// <summary>
        /// 获取最晚的时间
        /// </summary>
        public DateTime? LatestTime => PeriodDataList.Count > 0 ? PeriodDataList.Last().FileInfo?.FullDateTime : null;

        /// <summary>
        /// 获取时间跨度（天数）
        /// </summary>
        public int? TimeSpanDays
        {
            get
            {
                if (EarliestTime.HasValue && LatestTime.HasValue)
                {
                    return (LatestTime.Value - EarliestTime.Value).Days + 1;
                }
                return null;
            }
        }

        /// <summary>
        /// 获取验证失败的期数据
        /// </summary>
        /// <returns>验证失败的期数据列表</returns>
        public List<PeriodData> GetInvalidPeriodData()
        {
            return PeriodDataList.Where(pd => pd.ValidationStatus == ValidationStatus.Invalid).ToList();
        }

        /// <summary>
        /// 获取需要调整的期数据
        /// </summary>
        /// <returns>需要调整的期数据列表</returns>
        public List<PeriodData> GetPeriodDataNeedingAdjustment()
        {
            return PeriodDataList.Where(pd => pd.ValidationStatus == ValidationStatus.NeedsAdjustment).ToList();
        }

        /// <summary>
        /// 获取已调整的期数据
        /// </summary>
        /// <returns>已调整的期数据列表</returns>
        public List<PeriodData> GetAdjustedPeriodData()
        {
            return PeriodDataList.Where(pd => pd.HasBeenAdjusted).ToList();
        }

        /// <summary>
        /// 获取调整次数
        /// </summary>
        public int TotalAdjustmentCount => PeriodDataList.Sum(pd => pd.AdjustmentCount);

        /// <summary>
        /// 检查是否所有期数据都验证通过
        /// </summary>
        public bool AllPeriodDataValid => PeriodDataList.All(pd => pd.ValidationStatus == ValidationStatus.Valid);

        /// <summary>
        /// 检查是否有验证失败的期数据
        /// </summary>
        public bool HasInvalidPeriodData => PeriodDataList.Any(pd => pd.ValidationStatus == ValidationStatus.Invalid);

        /// <summary>
        /// 检查是否有需要调整的期数据
        /// </summary>
        public bool HasPeriodDataNeedingAdjustment => PeriodDataList.Any(pd => pd.ValidationStatus == ValidationStatus.NeedsAdjustment);

        /// <summary>
        /// 检查是否有已调整的期数据
        /// </summary>
        public bool HasAdjustedPeriodData => PeriodDataList.Any(pd => pd.HasBeenAdjusted);

        /// <summary>
        /// 设置对比数据
        /// </summary>
        /// <param name="comparisonData">对比数据</param>
        public void SetComparisonData(PeriodData? comparisonData)
        {
            if (comparisonData != null && 
                !string.Equals(comparisonData.PointName, PointName, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"对比数据的点名 {comparisonData.PointName} 与监测点名 {PointName} 不匹配");
            }

            ComparisonData = comparisonData;
        }

        /// <summary>
        /// 检查是否有对比数据
        /// </summary>
        public bool HasComparisonData => ComparisonData != null;

        /// <summary>
        /// 添加验证错误
        /// </summary>
        /// <param name="error">错误信息</param>
        public void AddValidationError(string error)
        {
            ValidationErrors.Add(error);
            ValidationStatus = ValidationStatus.Invalid;
        }

        /// <summary>
        /// 清除验证错误
        /// </summary>
        public void ClearValidationErrors()
        {
            ValidationErrors.Clear();
            ValidationStatus = ValidationStatus.Valid;
        }

        /// <summary>
        /// 添加调整记录
        /// </summary>
        /// <param name="adjustment">调整记录</param>
        public void AddAdjustment(AdjustmentRecord adjustment)
        {
            Adjustments.Add(adjustment);
        }

        /// <summary>
        /// 获取指定方向的累计变化量变化趋势
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>累计变化量变化趋势数据</returns>
        public List<(DateTime Time, double Value)> GetCumulativeTrend(DataDirection direction)
        {
            return PeriodDataList
                .Where(pd => pd.FileInfo != null)
                .Select(pd => (pd.FileInfo!.FullDateTime, pd.GetCumulativeValue(direction)))
                .OrderBy(x => x.FullDateTime)
                .ToList();
        }

        /// <summary>
        /// 获取指定方向的本期变化量变化趋势
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>本期变化量变化趋势数据</returns>
        public List<(DateTime Time, double Value)> GetCurrentPeriodTrend(DataDirection direction)
        {
            return PeriodDataList
                .Where(pd => pd.FileInfo != null)
                .Select(pd => (pd.FileInfo!.FullDateTime, pd.GetCurrentPeriodValue(direction)))
                .OrderBy(x => x.FullDateTime)
                .ToList();
        }

        /// <summary>
        /// 重写ToString方法
        /// </summary>
        public override string ToString()
        {
            return $"{PointName} (里程:{Mileage:F2}) - {PeriodDataCount}期数据";
        }

        /// <summary>
        /// 创建深拷贝
        /// </summary>
        /// <returns>深拷贝对象</returns>
        public MonitoringPoint Clone()
        {
            return new MonitoringPoint
            {
                PointName = PointName,
                Mileage = Mileage,
                PeriodDataList = PeriodDataList.Select(pd => pd.Clone()).ToList(),
                ComparisonData = ComparisonData?.Clone(),
                ValidationStatus = ValidationStatus,
                ValidationErrors = new List<string>(ValidationErrors),
                Adjustments = new List<AdjustmentRecord>(Adjustments)
            };
        }
    }
}
