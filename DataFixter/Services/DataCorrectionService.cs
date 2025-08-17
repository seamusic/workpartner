using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using Microsoft.Extensions.Logging;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据修正服务
    /// 实现数据修正算法，包括本期变化量调整、累计变化量调整、前后期衔接验证和最小化修改策略
    /// </summary>
    public class DataCorrectionService
    {
        private readonly ILogger<DataCorrectionService> _logger;
        private readonly CorrectionOptions _options;
        private readonly List<AdjustmentRecord> _adjustmentRecords;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">修正选项</param>
        public DataCorrectionService(ILogger<DataCorrectionService> logger, CorrectionOptions? options = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _options = options ?? new CorrectionOptions();
            _adjustmentRecords = new List<AdjustmentRecord>();
        }

        /// <summary>
        /// 修正所有监测点的数据
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <param name="validationResults">验证结果列表</param>
        /// <returns>修正结果</returns>
        public CorrectionResult CorrectAllPoints(List<MonitoringPoint> monitoringPoints, List<ValidationResult> validationResults)
        {
            var result = new CorrectionResult();
            var totalPoints = monitoringPoints.Count;
            var processedPoints = 0;

            try
            {
                _logger.LogInformation("开始修正 {TotalPoints} 个监测点的数据", totalPoints);

                // 按点名分组验证结果
                var validationByPoint = validationResults
                    .Where(v => v.Status == ValidationStatus.Invalid)
                    .GroupBy(v => v.PointName)
                    .ToDictionary(g => g.Key!, g => g.ToList());

                foreach (var point in monitoringPoints)
                {
                    try
                    {
                        if (validationByPoint.TryGetValue(point.PointName, out var pointValidations))
                        {
                            var pointResult = CorrectSinglePoint(point, pointValidations);
                            result.AddPointResult(pointResult);
                        }
                        else
                        {
                            // 没有验证错误，添加成功结果
                            result.AddPointResult(new PointCorrectionResult
                            {
                                PointName = point.PointName,
                                Status = CorrectionStatus.Success,
                                Message = "数据验证通过，无需修正"
                            });
                        }

                        processedPoints++;
                        if (processedPoints % 10 == 0)
                        {
                            _logger.LogInformation("已处理 {ProcessedPoints}/{TotalPoints} 个监测点", processedPoints, totalPoints);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "修正监测点 {PointName} 时发生异常", point.PointName);
                        
                        result.AddPointResult(new PointCorrectionResult
                        {
                            PointName = point.PointName,
                            Status = CorrectionStatus.Error,
                            Message = $"修正过程中发生异常: {ex.Message}"
                        });
                    }
                }

                result.AdjustmentRecords = _adjustmentRecords;
                _logger.LogInformation("数据修正完成: 总计 {TotalPoints} 个监测点, 生成 {RecordCount} 个修正记录", 
                    totalPoints, _adjustmentRecords.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "修正所有监测点数据时发生异常");
                result.Status = CorrectionStatus.Error;
                result.Message = $"修正过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 修正单个监测点的数据
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="validations">验证结果列表</param>
        /// <returns>修正结果</returns>
        private PointCorrectionResult CorrectSinglePoint(MonitoringPoint point, List<ValidationResult> validations)
        {
            var result = new PointCorrectionResult
            {
                PointName = point.PointName,
                Status = CorrectionStatus.Success,
                Message = "修正完成"
            };

            try
            {
                if (point.PeriodDataCount < 2)
                {
                    result.Status = CorrectionStatus.Skipped;
                    result.Message = "监测点数据不足，无法进行修正";
                    return result;
                }

                // 按时间排序
                var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();
                var corrections = new List<DataCorrection>();

                // 从前往后逐期修正，确保前后期衔接
                for (int i = 1; i < sortedData.Count; i++)
                {
                    var previousPeriod = sortedData[i - 1];
                    var currentPeriod = sortedData[i];

                    // 验证并修正累计变化量计算错误
                    var periodCorrections = CorrectPeriodData(previousPeriod, currentPeriod, point.PointName);
                    if (periodCorrections.Any())
                    {
                        corrections.AddRange(periodCorrections);
                        
                        // 应用修正
                        ApplyCorrections(currentPeriod, periodCorrections);
                        
                        // 记录修正记录
                        foreach (var correction in periodCorrections)
                        {
                            var adjustmentRecord = new AdjustmentRecord(
                                GetAdjustmentType(correction.CorrectionType),
                                correction.Direction,
                                correction.OriginalValue,
                                correction.CorrectedValue,
                                correction.Reason
                            );
                            
                            adjustmentRecord.SetFileInfo(
                                currentPeriod.FileInfo?.OriginalFileName ?? "",
                                point.PointName,
                                currentPeriod.RowNumber
                            );
                            
                            _adjustmentRecords.Add(adjustmentRecord);
                        }
                    }
                }

                result.Corrections = corrections;
                result.CorrectedPeriods = corrections.Select(c => c.PeriodData).Distinct().Count();
                result.CorrectedValues = corrections.Count;

                if (corrections.Any())
                {
                    result.Message = $"修正了 {corrections.Count} 个数据值，涉及 {result.CorrectedPeriods} 个期次";
                }
                else
                {
                    result.Message = "数据验证通过，无需修正";
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "修正监测点 {PointName} 时发生异常", point.PointName);
                result.Status = CorrectionStatus.Error;
                result.Message = $"修正过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 修正单期数据
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="pointName">点名</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> CorrectPeriodData(PeriodData previousPeriod, PeriodData currentPeriod, string pointName)
        {
            var corrections = new List<DataCorrection>();

            try
            {
                // 验证并修正X方向
                var xCorrection = CorrectDirection(previousPeriod, currentPeriod, DataDirection.X, pointName);
                if (xCorrection != null) corrections.Add(xCorrection);

                // 验证并修正Y方向
                var yCorrection = CorrectDirection(previousPeriod, currentPeriod, DataDirection.Y, pointName);
                if (yCorrection != null) corrections.Add(yCorrection);

                // 验证并修正Z方向
                var zCorrection = CorrectDirection(previousPeriod, currentPeriod, DataDirection.Z, pointName);
                if (zCorrection != null) corrections.Add(zCorrection);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "修正监测点 {PointName} 期数据时发生异常", pointName);
            }

            return corrections;
        }

        /// <summary>
        /// 修正单个方向的数据
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="direction">数据方向</param>
        /// <param name="pointName">点名</param>
        /// <returns>修正信息</returns>
        private DataCorrection? CorrectDirection(PeriodData previousPeriod, PeriodData currentPeriod, DataDirection direction, string pointName)
        {
            try
            {
                var (previousCumulative, currentPeriodValue, currentCumulative) = GetDirectionValues(previousPeriod, currentPeriod, direction);
                var expectedCumulative = previousCumulative + currentPeriodValue;
                var difference = Math.Abs(currentCumulative - expectedCumulative);

                // 检查是否在容差范围内
                if (difference <= _options.CumulativeTolerance)
                {
                    return null; // 无需修正
                }

                // 实现最小化修改策略
                return ApplyMinimalCorrectionStrategy(
                    previousPeriod, currentPeriod, direction, 
                    previousCumulative, currentPeriodValue, currentCumulative, 
                    expectedCumulative, difference, pointName);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "修正 {Direction} 方向数据时发生异常", direction);
                return null;
            }
        }

        /// <summary>
        /// 应用最小化修正策略
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="direction">数据方向</param>
        /// <param name="previousCumulative">上一期累计值</param>
        /// <param name="currentPeriodValue">当前期变化量</param>
        /// <param name="currentCumulative">当前期累计值</param>
        /// <param name="expectedCumulative">期望累计值</param>
        /// <param name="difference">差异值</param>
        /// <param name="pointName">点名</param>
        /// <returns>修正信息</returns>
        private DataCorrection ApplyMinimalCorrectionStrategy(
            PeriodData previousPeriod, PeriodData currentPeriod, DataDirection direction,
            double previousCumulative, double currentPeriodValue, double currentCumulative,
            double expectedCumulative, double difference, string pointName)
        {
            // 策略1: 优先调整本期变化量
            var correctedPeriodValue = currentPeriodValue + (expectedCumulative - currentCumulative);
            
            // 检查调整后的本期变化量是否在合理范围内
            if (Math.Abs(correctedPeriodValue) <= _options.MaxCurrentPeriodValue)
            {
                return new DataCorrection
                {
                    PeriodData = currentPeriod,
                    Direction = direction,
                    CorrectionType = CorrectionType.CurrentPeriodValue,
                    OriginalValue = currentPeriodValue,
                    CorrectedValue = correctedPeriodValue,
                    Reason = $"调整本期变化量以修正累计变化量计算错误，期望累计值: {expectedCumulative:F6}"
                };
            }

            // 策略2: 当本期变化量调整后仍不满足要求时，调整累计变化量
            var correctedCumulative = expectedCumulative;
            
            // 确保调整后的累计变化量绝对值不超过限制
            if (Math.Abs(correctedCumulative) > _options.MaxCumulativeValue)
            {
                // 如果超出限制，尝试调整本期变化量到合理范围
                var maxAllowedPeriodValue = _options.MaxCurrentPeriodValue;
                var minAllowedPeriodValue = -_options.MaxCurrentPeriodValue;
                
                var adjustedPeriodValue = Math.Max(minAllowedPeriodValue, 
                    Math.Min(maxAllowedPeriodValue, currentPeriodValue));
                
                correctedCumulative = previousCumulative + adjustedPeriodValue;
                
                return new DataCorrection
                {
                    PeriodData = currentPeriod,
                    Direction = direction,
                    CorrectionType = CorrectionType.Both,
                    OriginalValue = currentPeriodValue,
                    CorrectedValue = adjustedPeriodValue,
                    Reason = $"同时调整本期变化量和累计变化量，确保数据在合理范围内"
                };
            }

            return new DataCorrection
            {
                PeriodData = currentPeriod,
                Direction = direction,
                CorrectionType = CorrectionType.CumulativeValue,
                OriginalValue = currentCumulative,
                CorrectedValue = correctedCumulative,
                Reason = $"调整累计变化量以修正计算错误，期望值: {expectedCumulative:F6}"
            };
        }

        /// <summary>
        /// 获取指定方向的数据值
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="direction">数据方向</param>
        /// <returns>数据值元组</returns>
        private (double previousCumulative, double currentPeriodValue, double currentCumulative) 
            GetDirectionValues(PeriodData previousPeriod, PeriodData currentPeriod, DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => (previousPeriod.CumulativeX, currentPeriod.CurrentPeriodX, currentPeriod.CumulativeX),
                DataDirection.Y => (previousPeriod.CumulativeY, currentPeriod.CurrentPeriodY, currentPeriod.CumulativeY),
                DataDirection.Z => (previousPeriod.CumulativeZ, currentPeriod.CurrentPeriodZ, currentPeriod.CumulativeZ),
                _ => throw new ArgumentException($"不支持的数据方向: {direction}")
            };
        }

        /// <summary>
        /// 应用修正到数据
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="corrections">修正列表</param>
        private void ApplyCorrections(PeriodData periodData, List<DataCorrection> corrections)
        {
            foreach (var correction in corrections)
            {
                switch (correction.CorrectionType)
                {
                    case CorrectionType.CurrentPeriodValue:
                        ApplyCurrentPeriodValueCorrection(periodData, correction);
                        break;
                    case CorrectionType.CumulativeValue:
                        ApplyCumulativeValueCorrection(periodData, correction);
                        break;
                    case CorrectionType.Both:
                        ApplyBothCorrections(periodData, correction);
                        break;
                }
            }
        }

        /// <summary>
        /// 应用本期变化量修正
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="correction">修正信息</param>
        private void ApplyCurrentPeriodValueCorrection(PeriodData periodData, DataCorrection correction)
        {
            switch (correction.Direction)
            {
                case DataDirection.X:
                    periodData.CurrentPeriodX = correction.CorrectedValue;
                    break;
                case DataDirection.Y:
                    periodData.CurrentPeriodY = correction.CorrectedValue;
                    break;
                case DataDirection.Z:
                    periodData.CurrentPeriodZ = correction.CorrectedValue;
                    break;
            }
        }

        /// <summary>
        /// 应用累计变化量修正
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="correction">修正信息</param>
        private void ApplyCumulativeValueCorrection(PeriodData periodData, DataCorrection correction)
        {
            switch (correction.Direction)
            {
                case DataDirection.X:
                    periodData.CumulativeX = correction.CorrectedValue;
                    break;
                case DataDirection.Y:
                    periodData.CumulativeY = correction.CorrectedValue;
                    break;
                case DataDirection.Z:
                    periodData.CumulativeZ = correction.CorrectedValue;
                    break;
            }
        }

        /// <summary>
        /// 应用双重修正
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="correction">修正信息</param>
        private void ApplyBothCorrections(PeriodData periodData, DataCorrection correction)
        {
            // 先应用本期变化量修正
            ApplyCurrentPeriodValueCorrection(periodData, correction);
            
            // 重新计算累计变化量
            var previousPeriod = GetPreviousPeriod(periodData);
            if (previousPeriod != null)
            {
                var newCumulative = GetPreviousCumulative(previousPeriod, correction.Direction) + correction.CorrectedValue;
                ApplyCumulativeValueCorrection(periodData, new DataCorrection
                {
                    Direction = correction.Direction,
                    CorrectedValue = newCumulative
                });
            }
        }

        /// <summary>
        /// 获取上一期数据
        /// </summary>
        /// <param name="currentPeriod">当前期数据</param>
        /// <returns>上一期数据</returns>
        private PeriodData? GetPreviousPeriod(PeriodData currentPeriod)
        {
            // 这里需要根据实际的数据结构来获取上一期数据
            // 由于当前实现中PeriodData没有直接引用上一期，我们需要通过其他方式获取
            // 暂时返回null，实际使用时需要传入完整的监测点数据
            return null;
        }

        /// <summary>
        /// 获取上一期累计值
        /// </summary>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="direction">数据方向</param>
        /// <returns>累计值</returns>
        private double GetPreviousCumulative(PeriodData previousPeriod, DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => previousPeriod.CumulativeX,
                DataDirection.Y => previousPeriod.CumulativeY,
                DataDirection.Z => previousPeriod.CumulativeZ,
                _ => 0.0
            };
        }

        /// <summary>
        /// 将修正类型转换为调整类型
        /// </summary>
        /// <param name="correctionType">修正类型</param>
        /// <returns>调整类型</returns>
        private AdjustmentType GetAdjustmentType(CorrectionType correctionType)
        {
            return correctionType switch
            {
                CorrectionType.CurrentPeriodValue => AdjustmentType.CurrentPeriod,
                CorrectionType.CumulativeValue => AdjustmentType.Cumulative,
                CorrectionType.Both => AdjustmentType.Cumulative, // 双重修正时使用累计类型
                _ => AdjustmentType.None
            };
        }

        /// <summary>
        /// 获取修正统计信息
        /// </summary>
        /// <returns>统计信息</returns>
        public CorrectionStatistics GetCorrectionStatistics()
        {
            var statistics = new CorrectionStatistics
            {
                TotalAdjustments = _adjustmentRecords.Count,
                TotalPoints = _adjustmentRecords.Select(r => r.PointName).Distinct().Count(),
                TotalFiles = _adjustmentRecords.Select(r => r.FileName).Where(f => !string.IsNullOrEmpty(f)).Distinct().Count()
            };

            // 按修正类型统计
            statistics.CorrectionTypeStats = _adjustmentRecords
                .GroupBy(r => r.AdjustmentType)
                .ToDictionary(g => g.Key.ToString(), g => g.Count());

            // 按数据方向统计
            statistics.DirectionStats = _adjustmentRecords
                .GroupBy(r => r.DataDirection)
                .ToDictionary(g => g.Key.ToString(), g => g.Count());

            // 按点名统计
            statistics.PointNameStats = _adjustmentRecords
                .GroupBy(r => r.PointName)
                .ToDictionary(g => g.Key, g => g.Count());

            return statistics;
        }
    }

    /// <summary>
    /// 修正选项
    /// </summary>
    public class CorrectionOptions
    {
        /// <summary>
        /// 累计变化量容差
        /// </summary>
        public double CumulativeTolerance { get; set; } = 0.001;

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

        /// <summary>
        /// 是否启用前后期衔接验证
        /// </summary>
        public bool EnableContinuityValidation { get; set; } = true;
    }

    // DataCorrection 和 CorrectionType 已移动到 Models/CorrectionModels.cs

    // CorrectionStatus 和 PointCorrectionResult 已移动到 Models/CorrectionModels.cs

    // CorrectionResult 类已移动到 Models/CorrectionModels.cs

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
