using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据分组服务，负责将数据按监测点名称分组并排序
    /// </summary>
    public class DataGroupingService
    {
        private readonly ILogger _logger;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public DataGroupingService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// 按点名分组数据
        /// </summary>
        /// <param name="periodDataList">期数据列表</param>
        /// <returns>按点名分组的监测点列表</returns>
        public List<MonitoringPoint> GroupByPointName(List<PeriodData> periodDataList)
        {
            var monitoringPoints = new List<MonitoringPoint>();
            var groupedData = new Dictionary<string, List<PeriodData>>();

            try
            {
                _logger.Information("开始按点名分组数据，总计 {TotalRows} 行", periodDataList.Count);

                // 按点名分组
                foreach (var data in periodDataList)
                {
                    if (string.IsNullOrWhiteSpace(data.PointName))
                    {
                        _logger.Warning("跳过空点名的数据行: 行号 {RowNumber}", data.RowNumber);
                        continue;
                    }

                    var pointName = data.PointName.Trim();
                    if (!groupedData.ContainsKey(pointName))
                    {
                        groupedData[pointName] = new List<PeriodData>();
                    }
                    groupedData[pointName].Add(data);
                }

                // 创建监测点对象
                foreach (var kvp in groupedData)
                {
                    var pointName = kvp.Key;
                    var dataList = kvp.Value;

                    if (dataList.Count == 0) continue;

                    // 获取里程（使用第一个非零里程，如果没有则使用0）
                    var mileage = dataList.FirstOrDefault(d => d.Mileage != 0)?.Mileage ?? 0.0;

                    var monitoringPoint = new MonitoringPoint(pointName, mileage);

                    // 添加所有期数据
                    foreach (var data in dataList)
                    {
                        monitoringPoint.AddPeriodData(data);
                    }

                    monitoringPoints.Add(monitoringPoint);
                }

                _logger.Information("数据分组完成: 总计 {PointCount} 个监测点", monitoringPoints.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "按点名分组数据时发生异常");
            }

            return monitoringPoints;
        }

        /// <summary>
        /// 按时间排序所有监测点的数据
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        public void SortAllPointsByTime(List<MonitoringPoint> monitoringPoints)
        {
            try
            {
                _logger.Information("开始按时间排序 {PointCount} 个监测点的数据", monitoringPoints.Count);

                foreach (var point in monitoringPoints)
                {
                    point.SortPeriodDataByTime();
                }

                _logger.Information("时间排序完成");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "按时间排序数据时发生异常");
            }
        }

        /// <summary>
        /// 检查期数据完整性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>问题列表</returns>
        private List<DataIntegrityIssue> CheckPeriodDataIntegrity(MonitoringPoint point)
        {
            var issues = new List<DataIntegrityIssue>();

            try
            {
                foreach (var periodData in point.PeriodDataList)
                {
                    // 检查文件信息
                    if (periodData.FileInfo == null)
                    {
                        issues.Add(new DataIntegrityIssue
                        {
                            PointName = point.PointName ?? "未知",
                            IssueType = DataIntegrityIssueType.MissingFileInfo,
                            Description = $"行号 {periodData.RowNumber} 缺少文件信息",
                            Severity = DataIntegrityIssueSeverity.Warning
                        });
                    }

                    // 检查数值是否异常
                    var valueIssues = CheckValueIntegrity(point.PointName ?? "未知", periodData);
                    issues.AddRange(valueIssues);

                    // 检查时间连续性
                    var timeIssues = CheckTimeContinuity(point, periodData);
                    issues.AddRange(timeIssues);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "检查期数据完整性失败: 点名 {PointName}", point.PointName);
            }

            return issues;
        }

        /// <summary>
        /// 检查数值完整性
        /// </summary>
        /// <param name="pointName">点名</param>
        /// <param name="periodData">期数据</param>
        /// <returns>问题列表</returns>
        private List<DataIntegrityIssue> CheckValueIntegrity(string pointName, PeriodData periodData)
        {
            var issues = new List<DataIntegrityIssue>();

            try
            {
                // 检查本期变化量
                if (Math.Abs(periodData.CurrentPeriodX) > 1000 ||
                    Math.Abs(periodData.CurrentPeriodY) > 1000 ||
                    Math.Abs(periodData.CurrentPeriodZ) > 1000)
                {
                    issues.Add(new DataIntegrityIssue
                    {
                        PointName = pointName,
                        IssueType = DataIntegrityIssueType.ExtremeValue,
                        Description = $"行号 {periodData.RowNumber} 本期变化量超出正常范围",
                        Severity = DataIntegrityIssueSeverity.Warning
                    });
                }

                // 检查累计变化量
                if (Math.Abs(periodData.CumulativeX) > 10000 ||
                    Math.Abs(periodData.CumulativeY) > 10000 ||
                    Math.Abs(periodData.CumulativeZ) > 10000)
                {
                    issues.Add(new DataIntegrityIssue
                    {
                        PointName = pointName,
                        IssueType = DataIntegrityIssueType.ExtremeValue,
                        Description = $"行号 {periodData.RowNumber} 累计变化量超出正常范围",
                        Severity = DataIntegrityIssueSeverity.Warning
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "检查数值完整性失败: 点名 {PointName}, 行号 {RowNumber}", pointName, periodData.RowNumber);
            }

            return issues;
        }

        /// <summary>
        /// 检查时间连续性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="periodData">期数据</param>
        /// <returns>问题列表</returns>
        private List<DataIntegrityIssue> CheckTimeContinuity(MonitoringPoint point, PeriodData periodData)
        {
            var issues = new List<DataIntegrityIssue>();

            try
            {
                if (periodData.FileInfo == null) return issues;

                // 检查时间是否在合理范围内
                var currentTime = periodData.FileInfo.FullDateTime;
                var minTime = new DateTime(2020, 1, 1);
                var maxTime = DateTime.Now.AddDays(30);

                if (currentTime < minTime || currentTime > maxTime)
                {
                    issues.Add(new DataIntegrityIssue
                    {
                        PointName = point.PointName ?? "未知",
                        IssueType = DataIntegrityIssueType.InvalidTime,
                        Description = $"行号 {periodData.RowNumber} 时间异常: {currentTime:yyyy-MM-dd HH:mm}",
                        Severity = DataIntegrityIssueSeverity.Warning
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "检查时间连续性失败: 点名 {PointName}, 行号 {RowNumber}", point.PointName, periodData.RowNumber);
            }

            return issues;
        }
    }

    /// <summary>
    /// 数据完整性问题类型
    /// </summary>
    public enum DataIntegrityIssueType
    {
        /// <summary>
        /// 空点名
        /// </summary>
        EmptyPointName,

        /// <summary>
        /// 无效里程
        /// </summary>
        InvalidMileage,

        /// <summary>
        /// 缺少文件信息
        /// </summary>
        MissingFileInfo,

        /// <summary>
        /// 没有期数据
        /// </summary>
        NoPeriodData,

        /// <summary>
        /// 重复数据
        /// </summary>
        DuplicateData,

        /// <summary>
        /// 极端值
        /// </summary>
        ExtremeValue,

        /// <summary>
        /// 无效时间
        /// </summary>
        InvalidTime
    }

    /// <summary>
    /// 数据完整性问题严重程度
    /// </summary>
    public enum DataIntegrityIssueSeverity
    {
        /// <summary>
        /// 信息
        /// </summary>
        Info,

        /// <summary>
        /// 警告
        /// </summary>
        Warning,

        /// <summary>
        /// 严重
        /// </summary>
        Critical
    }

    /// <summary>
    /// 数据分组统计信息
    /// </summary>
    public class DataGroupingStatistics
    {
        /// <summary>
        /// 总监测点数
        /// </summary>
        public int TotalPoints { get; set; }

        /// <summary>
        /// 总期数据数
        /// </summary>
        public int TotalPeriods { get; set; }

        /// <summary>
        /// 每个监测点的平均期数据数
        /// </summary>
        public double AveragePeriodsPerPoint { get; set; }

        /// <summary>
        /// 单个监测点的最大期数据数
        /// </summary>
        public int MaxPeriodsPerPoint { get; set; }

        /// <summary>
        /// 单个监测点的最小期数据数
        /// </summary>
        public int MinPeriodsPerPoint { get; set; }

        /// <summary>
        /// 只有一期数据的监测点数量
        /// </summary>
        public int PointsWithSinglePeriod { get; set; }

        /// <summary>
        /// 有多期数据的监测点数量
        /// </summary>
        public int PointsWithMultiplePeriods { get; set; }

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>摘要字符串</returns>
        public string GetSummary()
        {
            return $"数据分组统计: {TotalPoints}个点, {TotalPeriods}期数据, 平均每点{AveragePeriodsPerPoint:F1}期";
        }
    }
}
