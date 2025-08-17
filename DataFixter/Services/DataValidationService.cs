using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using Microsoft.Extensions.Logging;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据验证服务
    /// 实现累计变化量计算逻辑验证、与对比数据的交叉验证和异常数据标记功能
    /// </summary>
    public class DataValidationService
    {
        private readonly ILogger<DataValidationService> _logger;
        private readonly ValidationOptions _options;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">验证选项</param>
        public DataValidationService(ILogger<DataValidationService> logger, ValidationOptions? options = null)
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

            try
            {
                _logger.LogInformation("开始验证 {TotalPoints} 个监测点的数据逻辑", totalPoints);

                foreach (var point in monitoringPoints)
                {
                    try
                    {
                        var pointResults = ValidateSinglePoint(point, comparisonData);
                        validationResults.AddRange(pointResults);
                        processedPoints++;

                        if (processedPoints % 10 == 0)
                        {
                            _logger.LogInformation("已处理 {ProcessedPoints}/{TotalPoints} 个监测点", processedPoints, totalPoints);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "验证监测点 {PointName} 时发生异常", point.PointName);
                        
                        // 添加验证失败的结果
                        validationResults.Add(new ValidationResult(ValidationStatus.Invalid, "系统异常", $"验证过程中发生异常: {ex.Message}")
                        {
                            PointName = point.PointName,
                            Severity = ValidationSeverity.Critical
                        });
                    }
                }

                _logger.LogInformation("数据验证完成: 总计 {TotalPoints} 个监测点, 生成 {ResultCount} 个验证结果", 
                    totalPoints, validationResults.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "验证所有监测点数据时发生异常");
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
                results.AddRange(cumulativeValidationResults);

                // 与对比数据进行交叉验证
                var comparisonValidationResults = ValidateAgainstComparisonData(point, comparisonData);
                results.AddRange(comparisonValidationResults);

                // 验证数据连续性
                var continuityValidationResults = ValidateDataContinuity(point);
                results.AddRange(continuityValidationResults);

                // 如果没有验证失败，添加验证通过的结果
                if (!results.Any(r => r.Status == ValidationStatus.Invalid))
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
                _logger.LogWarning(ex, "验证监测点 {PointName} 时发生异常", point.PointName);
                results.Add(new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证过程中发生异常: {ex.Message}")
                {
                    PointName = point.PointName,
                    Severity = ValidationSeverity.Error
                });
            }

            return results;
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
                    if (xValidation != null) results.Add(xValidation);

                    // 验证Y方向
                    var yValidation = ValidateCumulativeDirection(
                        point.PointName, 
                        currentPeriod, 
                        previousPeriod.CumulativeY, 
                        currentPeriod.CurrentPeriodY, 
                        currentPeriod.CumulativeY, 
                        DataDirection.Y);
                    if (yValidation != null) results.Add(yValidation);

                    // 验证Z方向
                    var zValidation = ValidateCumulativeDirection(
                        point.PointName, 
                        currentPeriod, 
                        previousPeriod.CumulativeZ, 
                        currentPeriod.CurrentPeriodZ, 
                        currentPeriod.CumulativeZ, 
                        DataDirection.Z);
                    if (zValidation != null) results.Add(zValidation);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "验证监测点 {PointName} 累计变化量计算逻辑时发生异常", point.PointName);
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
                var expectedCumulative = previousCumulative + currentPeriodValue;
                var difference = Math.Abs(currentCumulative - expectedCumulative);

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
                _logger.LogWarning(ex, "验证 {Direction} 方向累计变化量时发生异常", direction);
                return null;
            }
        }

        /// <summary>
        /// 与对比数据进行交叉验证
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonData">对比数据</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateAgainstComparisonData(MonitoringPoint point, List<PeriodData> comparisonData)
        {
            var results = new List<ValidationResult>();

            try
            {
                if (comparisonData.Count == 0) return results;

                // 查找对比数据中对应的点名
                var comparisonPointData = comparisonData
                    .Where(pd => string.Equals(pd.PointName, point.PointName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (comparisonPointData.Count == 0)
                {
                    // 对比数据中没有找到对应点名
                    results.Add(new ValidationResult(ValidationStatus.NeedsAdjustment, "对比数据缺失", 
                        "对比数据中未找到对应点名")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Warning
                    });
                    return results;
                }

                // 验证对比数据的完整性
                foreach (var comparisonDataItem in comparisonPointData)
                {
                    var comparisonValidation = ValidateComparisonDataIntegrity(point, comparisonDataItem);
                    if (comparisonValidation != null) results.Add(comparisonValidation);
                }

                // 交叉验证数据一致性
                var crossValidationResults = ValidateDataConsistency(point, comparisonPointData);
                results.AddRange(crossValidationResults);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "与对比数据交叉验证监测点 {PointName} 时发生异常", point.PointName);
            }

            return results;
        }

        /// <summary>
        /// 验证对比数据完整性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonData">对比数据</param>
        /// <returns>验证结果</returns>
        private ValidationResult? ValidateComparisonDataIntegrity(MonitoringPoint point, PeriodData comparisonData)
        {
            try
            {
                // 检查对比数据是否为空
                if (Math.Abs(comparisonData.CurrentPeriodX) < _options.MinValueThreshold &&
                    Math.Abs(comparisonData.CurrentPeriodY) < _options.MinValueThreshold &&
                    Math.Abs(comparisonData.CurrentPeriodZ) < _options.MinValueThreshold)
                {
                    return new ValidationResult(ValidationStatus.NeedsAdjustment, "对比数据为空", 
                        "对比数据中所有方向的变化量都为空")
                    {
                        PointName = point.PointName,
                        FileName = comparisonData.FileInfo?.OriginalFileName,
                        RowNumber = comparisonData.RowNumber,
                        Severity = ValidationSeverity.Warning
                    };
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "验证对比数据完整性时发生异常");
                return null;
            }
        }

        /// <summary>
        /// 验证数据一致性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonDataList">对比数据列表</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateDataConsistency(MonitoringPoint point, List<PeriodData> comparisonDataList)
        {
            var results = new List<ValidationResult>();

            try
            {
                foreach (var comparisonData in comparisonDataList)
                {
                                    // 检查里程是否一致
                if (Math.Abs(point.Mileage - comparisonData.Mileage) > _options.MileageTolerance)
                {
                    results.Add(new ValidationResult(ValidationStatus.NeedsAdjustment, "里程不一致", 
                        $"监测点里程与对比数据不一致")
                    {
                        PointName = point.PointName,
                        FileName = comparisonData.FileInfo?.OriginalFileName,
                        RowNumber = comparisonData.RowNumber,
                        Severity = ValidationSeverity.Warning
                    });
                }

                    // 检查数据量级是否合理
                    var dataMagnitudeValidation = ValidateDataMagnitude(point, comparisonData);
                    if (dataMagnitudeValidation != null) results.Add(dataMagnitudeValidation);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "验证数据一致性时发生异常");
            }

            return results;
        }

        /// <summary>
        /// 验证数据量级
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonData">对比数据</param>
        /// <returns>验证结果</returns>
        private ValidationResult? ValidateDataMagnitude(MonitoringPoint point, PeriodData comparisonData)
        {
            try
            {
                // 检查本期变化量是否在合理范围内
                if (Math.Abs(comparisonData.CurrentPeriodX) > _options.MaxCurrentPeriodValue ||
                    Math.Abs(comparisonData.CurrentPeriodY) > _options.MaxCurrentPeriodValue ||
                    Math.Abs(comparisonData.CurrentPeriodZ) > _options.MaxCurrentPeriodValue)
                {
                    return new ValidationResult(ValidationStatus.NeedsAdjustment, "数据量级异常", 
                        "对比数据中本期变化量超出正常范围")
                    {
                        PointName = point.PointName,
                        FileName = comparisonData.FileInfo?.OriginalFileName,
                        RowNumber = comparisonData.RowNumber,
                        Severity = ValidationSeverity.Warning
                    };
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "验证数据量级时发生异常");
                return null;
            }
        }

        /// <summary>
        /// 验证数据连续性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateDataContinuity(MonitoringPoint point)
        {
            var results = new List<ValidationResult>();

            try
            {
                if (point.PeriodDataCount < 2) return results;

                var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();

                for (int i = 1; i < sortedData.Count; i++)
                {
                    var previousPeriod = sortedData[i - 1];
                    var currentPeriod = sortedData[i];

                    // 检查时间间隔是否合理
                    if (previousPeriod.FileInfo?.FullDateTime != null && 
                        currentPeriod.FileInfo?.FullDateTime != null)
                    {
                        var timeSpan = currentPeriod.FileInfo.FullDateTime - previousPeriod.FileInfo.FullDateTime;
                        
                        if (timeSpan.TotalDays > _options.MaxTimeInterval)
                        {
                            results.Add(new ValidationResult(ValidationStatus.NeedsAdjustment, "时间间隔异常", 
                                $"相邻两期数据时间间隔过长: {timeSpan.TotalDays:F1}天")
                            {
                                PointName = point.PointName,
                                FileName = currentPeriod.FileInfo?.OriginalFileName,
                                RowNumber = currentPeriod.RowNumber,
                                Severity = ValidationSeverity.Warning
                            });
                        }
                    }

                    // 检查数据跳跃是否合理
                    var dataJumpValidation = ValidateDataJump(previousPeriod, currentPeriod, point.PointName);
                    if (dataJumpValidation != null) results.Add(dataJumpValidation);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "验证数据连续性时发生异常");
            }

            return results;
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
                if (Math.Abs(currentPeriod.CurrentPeriodX) > _options.MaxCurrentPeriodValue ||
                    Math.Abs(currentPeriod.CurrentPeriodY) > _options.MaxCurrentPeriodValue ||
                    Math.Abs(currentPeriod.CurrentPeriodZ) > _options.MaxCurrentPeriodValue)
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
                _logger.LogWarning(ex, "验证数据跳跃时发生异常");
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
    /// 验证选项
    /// </summary>
    public class ValidationOptions
    {
        /// <summary>
        /// 累计变化量容差
        /// </summary>
        public double CumulativeTolerance { get; set; } = 0.001;

        /// <summary>
        /// 严重错误阈值
        /// </summary>
        public double CriticalThreshold { get; set; } = 1.0;

        /// <summary>
        /// 错误阈值
        /// </summary>
        public double ErrorThreshold { get; set; } = 0.5;

        /// <summary>
        /// 最小数值阈值
        /// </summary>
        public double MinValueThreshold { get; set; } = 0.001;

        /// <summary>
        /// 最大本期变化量
        /// </summary>
        public double MaxCurrentPeriodValue { get; set; } = 1.0;

        /// <summary>
        /// 里程容差
        /// </summary>
        public double MileageTolerance { get; set; } = 0.01;

        /// <summary>
        /// 最大时间间隔（天）
        /// </summary>
        public double MaxTimeInterval { get; set; } = 30.0;
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
