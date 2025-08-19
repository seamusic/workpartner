using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using DataFixter.Utils;
using Serilog;
using System.Threading;
using System.Threading.Tasks;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据验证服务，负责验证监测数据的完整性和一致性
    /// </summary>
    public class DataValidationService
    {
        private readonly ILogger _logger;
        private readonly ValidationOptions _options;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">验证选项</param>
        public DataValidationService(ILogger logger, ValidationOptions? options = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _options = options ?? new ValidationOptions();
        }

        /// <summary>
        /// 验证所有监测点的数据逻辑
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <param name="comparisonData">对比数据</param>
        /// <returns>验证结果列表</returns>
        public List<ValidationResult> ValidateAllPoints(List<MonitoringPoint> monitoringPoints, List<PeriodData> comparisonData)
        {
            var validationResults = new List<ValidationResult>();
            var totalPoints = monitoringPoints.Count;
            var processedPoints = 0;
            var lockObject = new object();

            try
            {
                _logger.Information("开始验证 {TotalPoints} 个监测点的数据逻辑", totalPoints);

                // 使用并行计算来提升性能，因为每个监测点之间没有依赖关系
                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = _options.MaxDegreeOfParallelism > 0 
                        ? _options.MaxDegreeOfParallelism 
                        : Environment.ProcessorCount, // 如果未配置，使用CPU核心数
                    CancellationToken = CancellationToken.None
                };

                _logger.Information("使用并行计算，并行度: {MaxDegreeOfParallelism}", parallelOptions.MaxDegreeOfParallelism);

                // 并行处理所有监测点
                Parallel.ForEach(monitoringPoints, parallelOptions, point =>
                {
                    try
                    {
                        var pointResults = ValidateSinglePoint(point, comparisonData);
                        
                        // 线程安全地添加结果
                        lock (lockObject)
                        {
                            validationResults.AddRange(pointResults);
                            processedPoints++;

                            if (processedPoints % 100 == 0)
                            {
                                _logger.Information("已处理 {ProcessedPoints}/{TotalPoints} 个监测点", processedPoints, totalPoints);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "验证监测点 {PointName} 时发生异常", point.PointName);

                        // 线程安全地添加验证失败的结果
                        lock (lockObject)
                        {
                            validationResults.Add(new ValidationResult(ValidationStatus.Invalid, "系统异常", $"验证过程中发生异常: {ex.Message}")
                            {
                                PointName = point.PointName,
                                Severity = ValidationSeverity.Critical
                            });
                        }
                    }
                });

                _logger.Information("数据验证完成: 总计 {TotalPoints} 个监测点, 生成 {ResultCount} 个验证结果",
                    totalPoints, validationResults.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "验证所有监测点数据时发生异常");
            }

            return validationResults;
        }

        /// <summary>
        /// 验证单个监测点的数据逻辑
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonData">对比数据</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateSinglePoint(MonitoringPoint point, List<PeriodData> comparisonData)
        {
            var results = new List<ValidationResult>();

            try
            {
                if (point.PeriodDataCount == 0)
                {
                    results.Add(new ValidationResult(ValidationStatus.Invalid, "数据缺失", "监测点没有期数据")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Critical
                    });
                    return results;
                }

                // 验证累计变化量计算逻辑
                var cumulativeValidationResults = ValidateCumulativeCalculation(point);
                // 如果直接校验正确，都不需要校验其它了，直接当这行记录是不需要修改的
                if (cumulativeValidationResults.Count == 0)
                {
                    _logger.Information("验证监测点 {PointName} 正确，不需要再校验", point.PointName);
                    results.Add(new ValidationResult(ValidationStatus.Valid, "数据验证", "累计变化量计算逻辑规则通过，不需要考虑其它处理了")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Info
                    });
                    return results;
                }

                // 与对比数据进行交叉验证，得出哪一些数据可以修改
                foreach (ValidationResult result in cumulativeValidationResults)
                {
                    // 查找对比数据中对应的点名
                    var comparisonPointData = comparisonData
                        .Where(pd => string.Equals(pd.PointName, point.PointName, StringComparison.OrdinalIgnoreCase) && pd.FormattedTime == result.FormattedTime)
                        .ToList();
                    // 取上一期数据，如果找不到，也可以修改
                    var perviousData = GetPreviousPeriodData(point, result.FormattedTime);
                    // 如果找不到，表示可以修改
                    if (comparisonPointData.Count == 0 || perviousData == null)
                    {
                        result.CanAdjustment = true;
                    }
                }
                results.AddRange(cumulativeValidationResults);

                // 如果没有验证失败，添加验证通过的结果
                if (results.All(r => r.Status != ValidationStatus.Invalid))
                {
                    results.Add(new ValidationResult(ValidationStatus.Valid, "数据验证", "所有验证规则通过")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Info
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "验证监测点 {PointName} 时发生异常", point.PointName);
                results.Add(new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证过程中发生异常: {ex.Message}")
                {
                    PointName = point.PointName,
                    Severity = ValidationSeverity.Error
                });
            }

            return results;
        }

        private PeriodData? GetPreviousPeriodData(MonitoringPoint point, string formattedTime)
        {
            // 将 formattedTime 转换为 DateTime 进行比较
            if (!DateTime.TryParse(formattedTime, out DateTime targetTime))
            {
                _logger.Warning("无法解析时间格式: {FormattedTime}", formattedTime);
                return null;
            }

            // 按时间排序，找到早于目标时间的最近一期数据
            var previousPeriod = point.PeriodDataList
                .Where(pd => pd.FileInfo?.FullDateTime != null && pd.FileInfo.FullDateTime < targetTime)
                .OrderByDescending(pd => pd.FileInfo.FullDateTime)
                .FirstOrDefault();

            return previousPeriod;
        }

        /// <summary>
        /// 验证累计变化量计算逻辑
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateCumulativeCalculation(MonitoringPoint point)
        {
            var results = new List<ValidationResult>();

            try
            {
                if (point.PeriodDataCount < 2) return results;

                // 按时间排序
                var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();

                // 逐期验证累计变化量计算逻辑
                for (int i = 1; i < sortedData.Count; i++)
                {
                    var previousPeriod = sortedData[i - 1];
                    var currentPeriod = sortedData[i];

                    // 验证X方向
                    var xValidation = ValidateCumulativeDirection(
                        point.PointName,
                        currentPeriod,
                        previousPeriod.CumulativeX,
                        currentPeriod.CurrentPeriodX,
                        currentPeriod.CumulativeX,
                        DataDirection.X);
                    if (xValidation != null)
                    {
                        xValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(xValidation);
                    }

                    // 验证Y方向
                    var yValidation = ValidateCumulativeDirection(
                        point.PointName,
                        currentPeriod,
                        previousPeriod.CumulativeY,
                        currentPeriod.CurrentPeriodY,
                        currentPeriod.CumulativeY,
                        DataDirection.Y);
                    if (yValidation != null)
                    {
                        yValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(yValidation);
                    }

                    // 验证Z方向
                    var zValidation = ValidateCumulativeDirection(
                        point.PointName,
                        currentPeriod,
                        previousPeriod.CumulativeZ,
                        currentPeriod.CurrentPeriodZ,
                        currentPeriod.CumulativeZ,
                        DataDirection.Z);
                    if (zValidation != null)
                    {
                        zValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(zValidation);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "验证监测点 {PointName} 累计变化量计算逻辑时发生异常", point.PointName);
            }

            return results;
        }

        /// <summary>
        /// 验证单个方向的累计变化量计算逻辑
        /// </summary>
        /// <param name="pointName">点名</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="previousCumulative">上一期累计值</param>
        /// <param name="currentPeriodValue">当前期变化量</param>
        /// <param name="currentCumulative">当前期累计值</param>
        /// <param name="direction">数据方向</param>
        /// <returns>验证结果</returns>
        private ValidationResult? ValidateCumulativeDirection(
            string pointName,
            PeriodData currentPeriod,
            double previousCumulative,
            double currentPeriodValue,
            double currentCumulative,
            DataDirection direction)
        {
            try
            {
                var expectedCumulative = FloatingPointUtils.ToDouble(FloatingPointUtils.SafeAdd(previousCumulative, currentPeriodValue));
                var difference = FloatingPointUtils.SafeAbsoluteDifference(currentCumulative, expectedCumulative);

                // 检查是否在容差范围内
                if (difference <= _options.CumulativeTolerance)
                {
                    return null; // 验证通过
                }

                // 计算验证失败的程度
                var severity = difference > _options.CriticalThreshold ? ValidationSeverity.Critical :
                              difference > _options.ErrorThreshold ? ValidationSeverity.Error :
                              ValidationSeverity.Warning;

                var result = new ValidationResult(ValidationStatus.Invalid, "累计变化量计算错误",
                    $"{direction}方向累计变化量计算错误")
                {
                    PointName = pointName,
                    FileName = currentPeriod.FileInfo?.OriginalFileName,
                    RowNumber = currentPeriod.RowNumber,
                    DataDirection = direction,
                    Severity = severity
                };

                // 添加详细错误信息
                result.AddErrorDetail($"期望累计值: {expectedCumulative:F6}");
                result.AddErrorDetail($"实际累计值: {currentCumulative:F6}");
                result.AddErrorDetail($"差异: {difference:F6}");
                result.AddErrorDetail($"上一期累计值: {previousCumulative:F6}");
                result.AddErrorDetail($"本期变化量: {currentPeriodValue:F6}");

                // 添加失败的数据值
                result.AddFailedValue("上一期累计值", previousCumulative);
                result.AddFailedValue("本期变化量", currentPeriodValue);
                result.AddFailedValue("实际累计值", currentCumulative);

                // 添加期望的数据值
                result.AddExpectedValue("期望累计值", expectedCumulative);

                // 设置验证规则
                result.SetValidationRule("累计变化量 = 上一期累计变化量 + 本期变化量");

                return result;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "验证 {Direction} 方向累计变化量时发生异常", direction);
                return null;
            }
        }

        /// <summary>
        /// 验证数据跳跃
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="pointName">点名</param>
        /// <returns>验证结果</returns>
        private ValidationResult? ValidateDataJump(PeriodData previousPeriod, PeriodData currentPeriod, string pointName)
        {
            try
            {
                // 检查本期变化量是否在合理范围内
                if (FloatingPointUtils.SafeAbs(currentPeriod.CurrentPeriodX) > _options.MaxCurrentPeriodValue ||
                    FloatingPointUtils.SafeAbs(currentPeriod.CurrentPeriodY) > _options.MaxCurrentPeriodValue ||
                    FloatingPointUtils.SafeAbs(currentPeriod.CurrentPeriodZ) > _options.MaxCurrentPeriodValue)
                {
                    return new ValidationResult(ValidationStatus.NeedsAdjustment, "数据跳跃异常",
                        "本期变化量超出正常范围")
                    {
                        PointName = pointName,
                        FileName = currentPeriod.FileInfo?.OriginalFileName,
                        RowNumber = currentPeriod.RowNumber,
                        Severity = ValidationSeverity.Warning
                    };
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "验证数据跳跃时发生异常");
                return null;
            }
        }

        /// <summary>
        /// 获取验证统计信息
        /// </summary>
        /// <param name="validationResults">验证结果列表</param>
        /// <returns>统计信息</returns>
        public ValidationStatistics GetValidationStatistics(List<ValidationResult> validationResults)
        {
            var statistics = new ValidationStatistics
            {
                TotalResults = validationResults.Count,
                ValidCount = validationResults.Count(r => r.Status == ValidationStatus.Valid),
                InvalidCount = validationResults.Count(r => r.Status == ValidationStatus.Invalid),
                NeedsAdjustmentCount = validationResults.Count(r => r.Status == ValidationStatus.NeedsAdjustment),
                CriticalCount = validationResults.Count(r => r.Severity == ValidationSeverity.Critical),
                ErrorCount = validationResults.Count(r => r.Severity == ValidationSeverity.Error),
                WarningCount = validationResults.Count(r => r.Severity == ValidationSeverity.Warning),
                InfoCount = validationResults.Count(r => r.Severity == ValidationSeverity.Info)
            };

            // 按验证类型统计
            statistics.ValidationTypeStats = validationResults
                .GroupBy(r => r.ValidationType)
                .ToDictionary(g => g.Key ?? "未知", g => g.Count());

            // 按点名统计
            statistics.PointNameStats = validationResults
                .Where(r => !string.IsNullOrEmpty(r.PointName))
                .GroupBy(r => r.PointName)
                .ToDictionary(g => g.Key!, g => g.Count());

            // 按文件统计
            statistics.FileNameStats = validationResults
                .Where(r => !string.IsNullOrEmpty(r.FileName))
                .GroupBy(r => r.FileName)
                .ToDictionary(g => g.Key!, g => g.Count());

            return statistics;
        }
    }

    /// <summary>
    /// 验证统计信息
    /// </summary>
    public class ValidationStatistics
    {
        /// <summary>
        /// 总验证结果数
        /// </summary>
        public int TotalResults { get; set; }

        /// <summary>
        /// 验证通过数
        /// </summary>
        public int ValidCount { get; set; }

        /// <summary>
        /// 验证失败数
        /// </summary>
        public int InvalidCount { get; set; }

        /// <summary>
        /// 需要调整数
        /// </summary>
        public int NeedsAdjustmentCount { get; set; }

        /// <summary>
        /// 严重问题数
        /// </summary>
        public int CriticalCount { get; set; }

        /// <summary>
        /// 错误数
        /// </summary>
        public int ErrorCount { get; set; }

        /// <summary>
        /// 警告数
        /// </summary>
        public int WarningCount { get; set; }

        /// <summary>
        /// 信息数
        /// </summary>
        public int InfoCount { get; set; }

        /// <summary>
        /// 按验证类型统计
        /// </summary>
        public Dictionary<string, int> ValidationTypeStats { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// 按点名统计
        /// </summary>
        public Dictionary<string, int> PointNameStats { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// 按文件统计
        /// </summary>
        public Dictionary<string, int> FileNameStats { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要字符串</returns>
        public string GetSummary()
        {
            return $"验证统计: 总计{TotalResults}个结果, 通过{ValidCount}个, 失败{InvalidCount}个, 需调整{NeedsAdjustmentCount}个";
        }

        /// <summary>
        /// 获取详细统计信息
        /// </summary>
        /// <returns>详细统计信息字符串</returns>
        public string GetDetailedInfo()
        {
            var info = $"验证详细统计:\n";
            info += $"总验证结果数: {TotalResults}\n";
            info += $"验证通过数: {ValidCount}\n";
            info += $"验证失败数: {InvalidCount}\n";
            info += $"需要调整数: {NeedsAdjustmentCount}\n";
            info += $"严重问题数: {CriticalCount}\n";
            info += $"错误数: {ErrorCount}\n";
            info += $"警告数: {WarningCount}\n";
            info += $"信息数: {InfoCount}\n";

            if (ValidationTypeStats.Count > 0)
            {
                info += $"\n按验证类型统计:\n";
                foreach (var kvp in ValidationTypeStats.OrderByDescending(x => x.Value))
                {
                    info += $"  {kvp.Key}: {kvp.Value}个\n";
                }
            }

            if (PointNameStats.Count > 0)
            {
                info += $"\n按点名统计 (前10名):\n";
                foreach (var kvp in PointNameStats.OrderByDescending(x => x.Value).Take(10))
                {
                    info += $"  {kvp.Key}: {kvp.Value}个\n";
                }
            }

            if (FileNameStats.Count > 0)
            {
                info += $"\n按文件统计 (前10名):\n";
                foreach (var kvp in FileNameStats.OrderByDescending(x => x.Value).Take(10))
                {
                    info += $"  {kvp.Key}: {kvp.Value}个\n";
                }
            }

            return info;
        }
    }
}
