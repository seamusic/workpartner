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
        /// 检查数据完整性
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <returns>数据完整性检查结果</returns>
        public DataIntegrityReport CheckDataIntegrity(List<MonitoringPoint> monitoringPoints)
        {
            var report = new DataIntegrityReport();
            var totalPoints = monitoringPoints.Count;
            var totalPeriods = 0;
            var missingDataCount = 0;
            var duplicateDataCount = 0;
            var invalidDataCount = 0;

            try
            {
                _logger.Information("开始检查 {PointCount} 个监测点的数据完整性", totalPoints);

                foreach (var point in monitoringPoints)
                {
                    totalPeriods += point.PeriodDataCount;
                    
                    // 检查点名是否为空
                    if (string.IsNullOrWhiteSpace(point.PointName))
                    {
                        report.AddIssue(new DataIntegrityIssue
                        {
                            PointName = "未知",
                            IssueType = DataIntegrityIssueType.EmptyPointName,
                            Description = "点名为空",
                            Severity = DataIntegrityIssueSeverity.Critical
                        });
                        invalidDataCount++;
                    }

                    // 检查里程是否有效
                    if (point.Mileage < 0 || point.Mileage > 100000)
                    {
                        report.AddIssue(new DataIntegrityIssue
                        {
                            PointName = point.PointName ?? "未知",
                            IssueType = DataIntegrityIssueType.InvalidMileage,
                            Description = $"里程值异常: {point.Mileage}",
                            Severity = DataIntegrityIssueSeverity.Warning
                        });
                        invalidDataCount++;
                    }

                    // 检查期数据完整性
                    var periodIssues = CheckPeriodDataIntegrity(point);
                    report.AddIssues(periodIssues);

                    // 统计缺失数据
                    if (point.PeriodDataCount == 0)
                    {
                        missingDataCount++;
                        report.AddIssue(new DataIntegrityIssue
                        {
                            PointName = point.PointName ?? "未知",
                            IssueType = DataIntegrityIssueType.NoPeriodData,
                            Description = "没有期数据",
                            Severity = DataIntegrityIssueSeverity.Critical
                        });
                    }

                    // 检查重复数据
                    var duplicateCount = CheckDuplicateData(point);
                    if (duplicateCount > 0)
                    {
                        duplicateDataCount += duplicateCount;
                        report.AddIssue(new DataIntegrityIssue
                        {
                            PointName = point.PointName ?? "未知",
                            IssueType = DataIntegrityIssueType.DuplicateData,
                            Description = $"发现 {duplicateCount} 条重复数据",
                            Severity = DataIntegrityIssueSeverity.Warning
                        });
                    }
                }

                // 设置报告统计信息
                report.TotalPoints = totalPoints;
                report.TotalPeriods = totalPeriods;
                report.MissingDataCount = missingDataCount;
                report.DuplicateDataCount = duplicateDataCount;
                report.InvalidDataCount = invalidDataCount;
                report.IssueCount = report.Issues.Count;

                _logger.Information("数据完整性检查完成: 总计 {PointCount} 个点, {PeriodCount} 期数据, 发现 {IssueCount} 个问题", 
                    totalPoints, totalPeriods, report.IssueCount);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "检查数据完整性时发生异常");
            }

            return report;
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

        /// <summary>
        /// 检查重复数据
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>重复数据数量</returns>
        private int CheckDuplicateData(MonitoringPoint point)
        {
            try
            {
                var duplicates = point.PeriodDataList
                    .GroupBy(pd => new { pd.FileInfo?.FullDateTime, pd.RowNumber })
                    .Where(g => g.Count() > 1)
                    .Sum(g => g.Count() - 1);

                return duplicates;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "检查重复数据失败: 点名 {PointName}", point.PointName);
                return 0;
            }
        }

        /// <summary>
        /// 获取分组统计信息
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <returns>统计信息</returns>
        public DataGroupingStatistics GetGroupingStatistics(List<MonitoringPoint> monitoringPoints)
        {
            var statistics = new DataGroupingStatistics
            {
                TotalPoints = monitoringPoints.Count,
                TotalPeriods = monitoringPoints.Sum(p => p.PeriodDataCount),
                AveragePeriodsPerPoint = monitoringPoints.Count > 0 ? 
                    (double)monitoringPoints.Sum(p => p.PeriodDataCount) / monitoringPoints.Count : 0
            };

            if (monitoringPoints.Count > 0)
            {
                statistics.MaxPeriodsPerPoint = monitoringPoints.Max(p => p.PeriodDataCount);
                statistics.MinPeriodsPerPoint = monitoringPoints.Min(p => p.PeriodDataCount);
                statistics.PointsWithSinglePeriod = monitoringPoints.Count(p => p.PeriodDataCount == 1);
                statistics.PointsWithMultiplePeriods = monitoringPoints.Count(p => p.PeriodDataCount > 1);
            }

            return statistics;
        }
    }

    /// <summary>
    /// 数据完整性报告
    /// </summary>
    public class DataIntegrityReport
    {
        /// <summary>
        /// 问题列表
        /// </summary>
        public List<DataIntegrityIssue> Issues { get; set; } = new List<DataIntegrityIssue>();

        /// <summary>
        /// 总监测点数
        /// </summary>
        public int TotalPoints { get; set; }

        /// <summary>
        /// 总期数据数
        /// </summary>
        public int TotalPeriods { get; set; }

        /// <summary>
        /// 缺失数据数量
        /// </summary>
        public int MissingDataCount { get; set; }

        /// <summary>
        /// 重复数据数量
        /// </summary>
        public int DuplicateDataCount { get; set; }

        /// <summary>
        /// 无效数据数量
        /// </summary>
        public int InvalidDataCount { get; set; }

        /// <summary>
        /// 问题总数
        /// </summary>
        public int IssueCount { get; set; }

        /// <summary>
        /// 添加问题
        /// </summary>
        /// <param name="issue">问题</param>
        public void AddIssue(DataIntegrityIssue issue)
        {
            Issues.Add(issue);
        }

        /// <summary>
        /// 添加多个问题
        /// </summary>
        /// <param name="issues">问题列表</param>
        public void AddIssues(List<DataIntegrityIssue> issues)
        {
            Issues.AddRange(issues);
        }

        /// <summary>
        /// 获取严重问题数量
        /// </summary>
        public int CriticalIssueCount => Issues.Count(i => i.Severity == DataIntegrityIssueSeverity.Critical);

        /// <summary>
        /// 获取警告问题数量
        /// </summary>
        public int WarningIssueCount => Issues.Count(i => i.Severity == DataIntegrityIssueSeverity.Warning);

        /// <summary>
        /// 获取信息问题数量
        /// </summary>
        public int InfoIssueCount => Issues.Count(i => i.Severity == DataIntegrityIssueSeverity.Info);

        /// <summary>
        /// 获取报告摘要
        /// </summary>
        /// <returns>摘要字符串</returns>
        public string GetSummary()
        {
            return $"数据完整性报告: {TotalPoints}个点, {TotalPeriods}期数据, {IssueCount}个问题 (严重{CriticalIssueCount}, 警告{WarningIssueCount})";
        }
    }

    /// <summary>
    /// 数据完整性问题
    /// </summary>
    public class DataIntegrityIssue
    {
        /// <summary>
        /// 点名
        /// </summary>
        public string PointName { get; set; } = string.Empty;

        /// <summary>
        /// 问题类型
        /// </summary>
        public DataIntegrityIssueType IssueType { get; set; }

        /// <summary>
        /// 问题描述
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// 严重程度
        /// </summary>
        public DataIntegrityIssueSeverity Severity { get; set; }
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
