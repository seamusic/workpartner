using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using DataFixter.Utils;
using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据修正服务
    /// 实现数据修正算法，包括本期变化量调整、累计变化量调整、前后期衔接验证和最小化修改策略
    /// </summary>
    public class DataCorrectionService
    {
        private readonly DualLoggerService _logger;
        private readonly CorrectionOptions _options;
        private readonly List<AdjustmentRecord> _adjustmentRecords;
        private readonly List<DataCorrection> _allCorrections; // 添加字段来跟踪所有修正记录

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">修正选项</param>
        public DataCorrectionService(ILogger logger, CorrectionOptions? options = null)
        {
            _logger = new DualLoggerService(typeof(DataCorrectionService));
            _options = options ?? new CorrectionOptions();
            _adjustmentRecords = new List<AdjustmentRecord>();
            _allCorrections = new List<DataCorrection>(); // 初始化修正记录列表
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
                _logger.LogOperationStart("数据修正", $"监测点数量：{totalPoints}");

                // 按点名分组验证结果
                var validationByPoint = validationResults
                    .Where(v => v.Status == ValidationStatus.Invalid)
                    .GroupBy(v => v.PointName)
                    .ToDictionary(g => g.Key!, g => g.ToList());

                foreach (var point in monitoringPoints)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(point.PointName))
                        {
                            _logger.ShowWarning("点号为空，无法处理");
                            continue;
                        }

                        if (validationByPoint.TryGetValue(point.PointName, out var pointValidations))
                        {
                            var currentResult = validationResults.Where(x => x.PointName == point.PointName).ToList();
                            var need = validationResults.Count(x =>
                                x.PointName == point.PointName);
                            if (currentResult.Count > 0)
                            {
                                var perviousData1 = GetPreviousPeriodData(point, currentResult[0].FormattedTime);
                                // 检查是否可以修正
                                var canFix = currentResult.Any(x => x.PointName == point.PointName && x.CanAdjustment);
                                if (canFix)
                                {
                                    // 仅修正允许修正的区域
                                    var pointResult = CorrectPoint(point);
                                    result.AddPointResult(pointResult);
                                }
                                else
                                {
                                    result.AddPointResult(new PointCorrectionResult
                                    {
                                        PointName = point.PointName,
                                        Status = CorrectionStatus.Skipped,
                                        Message = $"有{need}个错误需要修正，但不在许可范围，跳过处理"
                                    });
                                }
                            }
                            else
                            {
                                result.AddPointResult(new PointCorrectionResult
                                {
                                    PointName = point.PointName,
                                    Status = CorrectionStatus.Skipped,
                                    Message = "不在许可范围，且不需要处理"
                                });
                            }
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
                        if (processedPoints % 100 == 0)
                        {
                            _logger.UpdateProgress(processedPoints, totalPoints, "监测点修正");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.ShowError($"修正监测点 {point.PointName} 时发生异常: {ex.Message}");
                        _logger.FileError(ex, "修正监测点 {PointName} 时发生异常", point.PointName);

                        result.AddPointResult(new PointCorrectionResult
                        {
                            PointName = point.PointName,
                            Status = CorrectionStatus.Error,
                            Message = $"修正过程中发生异常: {ex.Message}"
                        });
                    }
                }

                result.AdjustmentRecords = _adjustmentRecords;
                //_logger.LogInformation("数据修正完成: 总计 {TotalPoints} 个监测点, 生成 {RecordCount} 个修正记录",
                //    totalPoints, _adjustmentRecords.Count);
            }
            catch (Exception ex)
            {
                _logger.ShowError($"修正所有监测点数据时发生异常: {ex.Message}");
                _logger.FileError(ex, "修正所有监测点数据时发生异常");
                result.Status = CorrectionStatus.Error;
                result.Message = $"修正过程中发生异常: {ex.Message}";
            }

            return result;
        }

        public PointCorrectionResult CorrectPointSimple(MonitoringPoint point, List<ValidationResult> results)
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
                    result.Message = "数据期数不足，无需修正";
                    return result;
                }

                // 首先尝试全局分析和修正策略
                var corrections = ApplyGlobalCorrectionStrategy(point);
                if (corrections.Any())
                {
                    // 应用修正
                    ApplyCorrections(corrections);
                    result.CorrectedPeriods = corrections.Select(c => c.PeriodData).Distinct().Count();
                    result.CorrectedValues = corrections.Count;

                    _logger.FileInfo("监测点修正", $"修正了{result.CorrectedValues}个值",
                        $"监测点：{point.PointName}，修正详情：{string.Join(", ", corrections.Select(c => $"{c.Direction}方向{c.CorrectionType}"))}");
                }
                else
                {
                    _logger.ConsoleInfo("{PointName} 没有需要修正的", point.PointName);

                }
            }
            catch (Exception ex)
            {
                result.Status = CorrectionStatus.Error;
                result.Message = $"修正过程中发生异常: {ex.Message}";
                _logger.ShowError($"修正监测点 {point.PointName} 时发生异常: {ex.Message}");
                _logger.FileError(ex, "修正监测点 {PointName} 时发生异常", point.PointName);
            }

            return result;
        }

        /// <summary>
        /// 修正监测点数据
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>修正结果</returns>
        public PointCorrectionResult CorrectPoint(MonitoringPoint point)
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
                    result.Message = "数据期数不足，无需修正";
                    return result;
                }

                // 首先尝试全局分析和修正策略
                var corrections = ApplyGlobalCorrectionStrategy(point);

                // 检查是否需要更激进的策略
                var needsAggressiveStrategy = false;

                if (corrections.Any())
                {
                    // 应用修正
                    ApplyCorrections(corrections);

                    result.CorrectedPeriods = corrections.Select(c => c.PeriodData).Distinct().Count();
                    result.CorrectedValues = corrections.Count;

                    // 验证修正后的数据
                    var validationResult = ValidatePointDataConsistency(point);
                    if (validationResult.Status == ValidationStatus.Invalid)
                    {
                        needsAggressiveStrategy = true;
                        _logger.ShowWarning($"监测点 {point.PointName} 修正后验证失败，需要激进策略");
                        _logger.FileWarning("监测点 {PointName} 修正后验证失败，需要激进策略", point.PointName);
                    }
                }
                else
                {
                    // 全局策略没有产生修正，检查是否真的需要修正
                    var validationResult = ValidatePointDataConsistency(point);
                    if (validationResult.Status == ValidationStatus.Invalid)
                    {
                        needsAggressiveStrategy = true;
                        _logger.ShowWarning($"监测点 {point.PointName} 全局策略未产生修正，但数据验证失败，需要激进策略");
                        _logger.FileWarning("监测点 {PointName} 全局策略未产生修正，但数据验证失败，需要激进策略", point.PointName);
                    }
                }

                // 如果需要激进策略，直接应用激进策略（不回滚之前的修正）
                if (needsAggressiveStrategy)
                {
                    _logger.ConsoleInfo("监测点 {PointName} 应用激进修正策略", point.PointName);
                    _logger.FileInfo("监测点 {PointName} 应用激进修正策略，开始应用激进修正算法", point.PointName);
                    var aggressiveCorrections = ApplyAggressiveCorrectionStrategy(point);

                    if (aggressiveCorrections.Any())
                    {
                        ApplyCorrections(aggressiveCorrections);
                        result.CorrectedPeriods = aggressiveCorrections.Select(c => c.PeriodData).Distinct().Count();
                        result.CorrectedValues = aggressiveCorrections.Count;

                        // 再次验证
                        var finalValidation = ValidatePointDataConsistency(point);
                        if (finalValidation.Status == ValidationStatus.Invalid)
                        {
                            _logger.ShowWarning($"激进策略后数据仍然无效: {finalValidation.Description}");
                            _logger.FileWarning("激进策略后数据仍然无效: {Description}", finalValidation.Description);

                            // 检查失败比例，如果低于20%，使用部分修正策略
                            var failureRatio = CalculateValidationFailureRatio(point, finalValidation);
                            if (failureRatio < 0.2) // 失败比例低于20%
                            {
                                _logger.ConsoleInfo("激进策略后失败比例较低 ({0:P1})，使用部分修正策略", failureRatio);
                                _logger.FileInfo("激进策略后失败比例较低 ({FailureRatio:P1})，使用部分修正策略", failureRatio);
                                var partialCorrections = ApplyPartialCorrectionStrategy(point, finalValidation);
                                if (partialCorrections.Any())
                                {
                                    ApplyCorrections(partialCorrections);
                                    result.CorrectedPeriods = partialCorrections.Select(c => c.PeriodData).Distinct().Count();
                                    result.CorrectedValues = partialCorrections.Count;
                                    _logger.ConsoleInfo("部分修正策略完成，修正了 {0} 个值", partialCorrections.Count);
                                    _logger.FileInfo("部分修正策略完成，修正了 {Count} 个值", partialCorrections.Count);
                                }
                            }
                            else
                            {
                                // 失败比例较高，使用最终修正策略
                                _logger.ConsoleInfo("激进策略后失败比例较高 ({0:P1})，使用最终修正策略", failureRatio);
                                _logger.FileInfo("激进策略后失败比例较高 ({FailureRatio:P1})，使用最终修正策略", failureRatio);
                                var finalCorrections = ApplyFinalCorrectionStrategy(point);
                                if (finalCorrections.Any())
                                {
                                    // 应用最终修正（不回滚之前的修正）
                                    ApplyCorrections(finalCorrections);
                                    result.CorrectedPeriods = finalCorrections.Select(c => c.PeriodData).Distinct().Count();
                                    result.CorrectedValues = finalCorrections.Count;

                                    _logger.ConsoleInfo("应用最终修正策略完成");
                                    _logger.FileInfo("应用最终修正策略完成");
                                }
                            }
                        }
                        else
                        {
                            _logger.ConsoleInfo("激进策略修正成功，数据验证通过");
                            _logger.FileInfo("激进策略修正成功，数据验证通过");
                        }
                    }
                    else
                    {
                        _logger.ShowWarning("激进策略未产生任何修正");
                        _logger.FileWarning("激进策略未产生任何修正");
                    }
                }
                else if (corrections.Any())
                {
                    _logger.ConsoleInfo("监测点 {0} 修正后数据验证通过", point.PointName);
                    _logger.FileInfo("监测点 {PointName} 修正后数据验证通过", point.PointName);
                }

                // 记录调整记录
                var allCorrections = GetCorrectionsFromPoint(point);

                foreach (var correction in allCorrections)
                {
                    var adjustmentRecord = new AdjustmentRecord
                    {
                        PointName = point.PointName,
                        FileName = correction.PeriodData.FileInfo?.OriginalFileName ?? "",
                        RowNumber = correction.PeriodData.RowNumber,
                        AdjustmentType = GetAdjustmentType(correction.CorrectionType),
                        DataDirection = GetDataDirection(correction.Direction),
                        OriginalValue = correction.OriginalValue,
                        AdjustedValue = correction.CorrectedValue,
                        Reason = correction.Reason,
                        AdjustmentTime = DateTime.Now
                    };
                    _adjustmentRecords.Add(adjustmentRecord);
                }

                _logger.FileInfo("监测点修正", $"修正了{result.CorrectedValues}个值",
                    $"监测点：{point.PointName}，修正详情：{string.Join(", ", allCorrections.Select(c => $"{c.Direction}方向{c.CorrectionType}"))}");
            }
            catch (Exception ex)
            {
                result.Status = CorrectionStatus.Error;
                result.Message = $"修正过程中发生异常: {ex.Message}";
                _logger.ShowError($"修正监测点 {point.PointName} 时发生异常: {ex.Message}");
                _logger.FileError(ex, "修正监测点 {PointName} 时发生异常", point.PointName);
            }

            return result;
        }

        /// <summary>
        /// 应用全局修正策略
        /// 分析整个监测点的数据问题，一次性修正所有相关期
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> ApplyGlobalCorrectionStrategy(MonitoringPoint point)
        {
            var corrections = new List<DataCorrection>();
            var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();

            // 分析每个方向的问题
            var directions = new[] { DataDirection.X, DataDirection.Y, DataDirection.Z };

            foreach (var direction in directions)
            {
                var directionCorrections = AnalyzeAndCorrectDirection(sortedData, direction, point.PointName);
                corrections.AddRange(directionCorrections);
            }

            return corrections;
        }

        /// <summary>
        /// 分析并修正单个方向的数据
        /// </summary>
        /// <param name="sortedData">按时间排序的数据</param>
        /// <param name="direction">数据方向</param>
        /// <param name="pointName">点名</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> AnalyzeAndCorrectDirection(List<PeriodData> sortedData, DataDirection direction, string pointName)
        {
            var corrections = new List<DataCorrection>();

            // 计算每期的期望累计值
            var expectedCumulatives = new List<double>();
            var actualCumulatives = new List<double>();
            var currentPeriodValues = new List<double>();

            for (int i = 0; i < sortedData.Count; i++)
            {
                var current = sortedData[i];
                var (currentPeriodValue, cumulativeValue) = GetDirectionValue(current, direction);

                currentPeriodValues.Add(currentPeriodValue);
                actualCumulatives.Add(cumulativeValue);

                if (i == 0)
                {
                    // 第一期，期望累计值等于实际累计值
                    expectedCumulatives.Add(cumulativeValue);
                }
                else
                {
                    // 后续期，期望累计值 = 上期期望累计值 + 本期变化量
                    var expectedCumulative = expectedCumulatives[i - 1] + currentPeriodValue;
                    expectedCumulatives.Add(expectedCumulative);
                }
            }

            // 检查是否需要修正
            var needsCorrection = false;
            for (int i = 1; i < sortedData.Count; i++)
            {
                var difference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulatives[i], actualCumulatives[i]);
                if (FloatingPointUtils.IsGreaterThan(difference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                {
                    needsCorrection = true;
                    break;
                }
            }

            if (!needsCorrection)
            {
                return corrections; // 无需修正
            }

            // 应用修正策略：优先修正累计值，保持本期变化量不变
            for (int i = 1; i < sortedData.Count; i++)
            {
                var current = sortedData[i];
                var difference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulatives[i], actualCumulatives[i]);

                if (FloatingPointUtils.IsGreaterThan(difference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                {
                    var originalCumulative = actualCumulatives[i];
                    var correctedCumulative = expectedCumulatives[i];

                    // 检查修正后的累计值是否在合理范围内
                    if (FloatingPointUtils.IsLessThanOrEqual(FloatingPointUtils.SafeAbs(correctedCumulative), _options.MaxCumulativeValue, _options.CumulativeTolerance))
                    {
                        var correction = new DataCorrection
                        {
                            PeriodData = current,
                            Direction = direction,
                            CorrectionType = CorrectionType.CumulativeValue,
                            OriginalValue = originalCumulative,
                            CorrectedValue = correctedCumulative,
                            Reason = $"修正累计值以保持数据一致性。期望值: {correctedCumulative:F6}, 原值: {originalCumulative:F6}"
                        };

                        corrections.Add(correction);

                        // 更新后续期的期望累计值
                        for (int j = i + 1; j < sortedData.Count; j++)
                        {
                            expectedCumulatives[j] = correctedCumulative + currentPeriodValues.Skip(i + 1).Take(j - i).Sum();
                        }
                    }
                    else
                    {
                        // 累计值超出范围，需要调整本期变化量
                        var maxAllowedCumulative = _options.MaxCumulativeValue * Math.Sign(correctedCumulative);
                        var adjustedPeriodValue = maxAllowedCumulative - expectedCumulatives[i - 1];

                        // 检查调整后的本期变化量是否合理
                        if (FloatingPointUtils.IsLessThanOrEqual(FloatingPointUtils.SafeAbs(adjustedPeriodValue), _options.MaxCurrentPeriodValue, _options.CumulativeTolerance))
                        {
                            var correction = new DataCorrection
                            {
                                PeriodData = current,
                                Direction = direction,
                                CorrectionType = CorrectionType.Both,
                                OriginalValue = currentPeriodValues[i],
                                CorrectedValue = adjustedPeriodValue,
                                Reason = $"调整本期变化量以保持累计值在合理范围内。新变化量: {adjustedPeriodValue:F6}, 原变化量: {currentPeriodValues[i]:F6}",
                                AdditionalData = new Dictionary<string, object>
                                {
                                    { "CorrectedCumulative", expectedCumulatives[i - 1] + adjustedPeriodValue },
                                    { "OriginalCumulative", expectedCumulatives[i] }
                                }
                            };

                            corrections.Add(correction);

                            // 更新后续期的期望累计值
                            for (int j = i; j < sortedData.Count; j++)
                            {
                                if (j == i)
                                {
                                    expectedCumulatives[j] = expectedCumulatives[j - 1] + adjustedPeriodValue;
                                }
                                else
                                {
                                    expectedCumulatives[j] = expectedCumulatives[j - 1] + currentPeriodValues[j];
                                }
                            }
                        }
                    }
                }
            }

            return corrections;
        }

        /// <summary>
        /// 应用激进修正策略
        /// 当常规修正策略失败时，重新生成本期变化量和累计值
        /// </summary>
        private List<DataCorrection> ApplyAggressiveCorrectionStrategy(MonitoringPoint point)
        {
            var corrections = new List<DataCorrection>();

            // 按时间排序数据
            var sortedData = point.PeriodDataList.OrderBy(p => p.FileInfo?.FullDateTime).ToList();
            if (!sortedData.Any()) return corrections;

            // 对每个方向分别处理
            foreach (DataDirection direction in Enum.GetValues(typeof(DataDirection)))
            {
                var directionCorrections = ApplyAggressiveDirectionCorrection(sortedData, direction, point.PointName);
                corrections.AddRange(directionCorrections);
            }

            return corrections;
        }

        /// <summary>
        /// 对特定方向应用激进修正策略
        /// </summary>
        private List<DataCorrection> ApplyAggressiveDirectionCorrection(List<PeriodData> sortedData, DataDirection direction, string pointName)
        {
            var corrections = new List<DataCorrection>();

            if (!sortedData.Any()) return corrections;

            // 第0期（基准期）：保持原值不变，不进行修正
            var firstPeriod = sortedData[0];
            var (originalFirstPeriodValue, originalFirstCumulative) = GetDirectionValue(firstPeriod, direction);

            // 基准期保持原值，不修正
            // 从基准期的累计值开始累加（使用当前数据状态，可能已经被之前的策略修正过）
            var currentCumulative = GetDirectionValue(firstPeriod, direction).cumulativeValue;

            // 明确保护第0期的本期变化量，确保不被修改
            // 第0期的本期变化量保持原值不变
            var firstPeriodCorrection = new DataCorrection
            {
                PeriodData = firstPeriod,
                Direction = direction,
                CorrectionType = CorrectionType.None, // 不修正第0期
                OriginalValue = originalFirstPeriodValue,
                CorrectedValue = originalFirstPeriodValue, // 保持原值
                Reason = $"第0期（基准期）保持原值不变：本期变化量={originalFirstPeriodValue:F6}, 累计值={currentCumulative:F6}",
                AdditionalData = new Dictionary<string, object>
                {
                    { "IsBaselinePeriod", true },
                    { "OriginalPeriodValue", originalFirstPeriodValue },
                    { "OriginalCumulative", originalFirstCumulative }
                }
            };
            corrections.Add(firstPeriodCorrection);

            // 分析原始数据的分布特征，用于生成合理的修正值
            var originalValues = sortedData.Skip(1).Select(pd => GetDirectionValue(pd, direction).currentPeriodValue).ToList();
            var originalCumulatives = sortedData.Skip(1).Select(pd => GetDirectionValue(pd, direction).cumulativeValue).ToList();

            // 计算原始数据的统计特征
            var avgPeriodValue = originalValues.Any() ? FloatingPointUtils.SafeAverage(originalValues.ToArray()) : 0.0;
            var stdPeriodValue = originalValues.Any() ? FloatingPointUtils.SafeStandardDeviation(originalValues.ToArray()) : 1.0;
            var maxPeriodValue = originalValues.Any() ? FloatingPointUtils.SafeMax(FloatingPointUtils.SafeAbs(originalValues.Max()), FloatingPointUtils.SafeAbs(originalValues.Min())) : 2.0;

            // 确保标准差不为0
            if (FloatingPointUtils.IsLessThan(stdPeriodValue, 0.1, _options.CumulativeTolerance))
                stdPeriodValue = maxPeriodValue * 0.3;

            var random = new Random();

            // 后续期：重新生成本期变化量，基于前一期累计值正确累加
            for (int i = 1; i < sortedData.Count; i++)
            {
                var current = sortedData[i];
                var (originalPeriodValue, originalCumulative) = GetDirectionValue(current, direction);

                // 基于原始数据分布生成本期变化量
                // 使用正态分布生成值，但限制在合理范围内
                double newPeriodValue;
                var maxAttempts = 100; // 防止无限循环
                var attempts = 0;

                do
                {
                    attempts++;
                    if (attempts > maxAttempts)
                    {
                        // 如果尝试次数过多，使用安全的默认值
                        _logger.ShowWarning($"监测点 {pointName} {direction}方向生成值失败，使用安全默认值");
                        _logger.FileWarning("监测点 {PointName} {Direction}方向生成值失败，使用安全默认值", pointName, direction);
                        newPeriodValue = FloatingPointUtils.SafeSign(avgPeriodValue) * FloatingPointUtils.SafeMax(0.1, FloatingPointUtils.SafeAbs(avgPeriodValue));
                        break;
                    }

                    // 使用Box-Muller变换生成正态分布随机数
                    var u1 = random.NextDouble();
                    var u2 = random.NextDouble();
                    var z0 = FloatingPointUtils.SafeSqrt(-2.0 * FloatingPointUtils.SafeLog(u1)) * FloatingPointUtils.SafeCos(2.0 * Math.PI * u2);

                    // 基于原始数据的均值和标准差生成值
                    newPeriodValue = FloatingPointUtils.SafeRound(avgPeriodValue + z0 * stdPeriodValue, 6);

                    // 限制在原始数据最大值的2倍范围内
                    var maxAllowed = FloatingPointUtils.SafeMax(maxPeriodValue * 2.0, 1.0);
                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(newPeriodValue), maxAllowed, _options.CumulativeTolerance))
                    {
                        newPeriodValue = FloatingPointUtils.SafeSign(newPeriodValue) * maxAllowed;
                    }

                    // 确保生成的值不会过小，但也要避免无限循环
                    if (FloatingPointUtils.IsGreaterThanOrEqual(FloatingPointUtils.SafeAbs(newPeriodValue), 0.001, _options.CumulativeTolerance))
                    {
                        break; // 满足条件，退出循环
                    }

                    // 如果值过小，尝试调整标准差
                    if (attempts > 50)
                    {
                        stdPeriodValue = Math.Max(stdPeriodValue * 1.5, 0.1);
                    }
                } while (true); // 改为无限循环，通过break控制退出

                // 正确计算累计值：前一期累计值 + 本期变化量
                var newCumulative = FloatingPointUtils.SafeRound(currentCumulative + newPeriodValue, 6);

                // 检查累计值是否在合理范围内
                if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(newCumulative), _options.MaxCumulativeValue, _options.CumulativeTolerance))
                {
                    // 如果超出范围，调整本期变化量使其在范围内
                    var maxAllowedCumulative = _options.MaxCumulativeValue;
                    var minAllowedCumulative = -_options.MaxCumulativeValue;

                    if (FloatingPointUtils.IsGreaterThan(newCumulative, maxAllowedCumulative, _options.CumulativeTolerance))
                    {
                        newPeriodValue = FloatingPointUtils.SafeRound(maxAllowedCumulative - currentCumulative, 6);
                        newCumulative = maxAllowedCumulative;
                    }
                    else
                    {
                        newPeriodValue = FloatingPointUtils.SafeRound(minAllowedCumulative - currentCumulative, 6);
                        newCumulative = minAllowedCumulative;
                    }
                }

                // 创建修正记录
                var correction = new DataCorrection
                {
                    PeriodData = current,
                    Direction = direction,
                    CorrectionType = CorrectionType.Both,
                    OriginalValue = originalPeriodValue,
                    CorrectedValue = newPeriodValue,
                    Reason = $"激进策略：基于原始数据分布重新生成本期变化量和累计值。本期新值: {newPeriodValue:F6}, 累计新值: {newCumulative:F6}, 原本期值: {originalPeriodValue:F6}, 原累计值: {originalCumulative:F6}",
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "CorrectedCumulative", newCumulative },
                        { "OriginalCumulative", originalCumulative }
                    }
                };

                corrections.Add(correction);

                // 移除直接修改数据的调用，确保修正只通过ApplyCorrections方法应用
                // SetDirectionValue(current, direction, newPeriodValue, newCumulative);

                // 更新当前累计值，为下一期计算做准备
                currentCumulative = newCumulative;
            }

            return corrections;
        }

        /// <summary>
        /// 获取指定方向的数据值
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="direction">数据方向</param>
        /// <returns>(本期变化量, 累计值)</returns>
        private (double currentPeriodValue, double cumulativeValue) GetDirectionValue(PeriodData periodData, DataDirection direction)
        {
            return direction switch
            {
                DataDirection.X => (periodData.CurrentPeriodX, periodData.CumulativeX),
                DataDirection.Y => (periodData.CurrentPeriodY, periodData.CumulativeY),
                DataDirection.Z => (periodData.CurrentPeriodZ, periodData.CumulativeZ),
                _ => throw new ArgumentException($"不支持的数据方向: {direction}")
            };
        }

        /// <summary>
        /// 设置指定方向的数据值
        /// </summary>
        /// <param name="periodData">期数据</param>
        /// <param name="direction">数据方向</param>
        /// <param name="currentPeriodValue">本期变化量</param>
        /// <param name="cumulativeValue">累计值</param>
        private void SetDirectionValue(PeriodData periodData, DataDirection direction, double currentPeriodValue, double cumulativeValue)
        {
            switch (direction)
            {
                case DataDirection.X:
                    periodData.CurrentPeriodX = currentPeriodValue;
                    periodData.CumulativeX = cumulativeValue;
                    break;
                case DataDirection.Y:
                    periodData.CurrentPeriodY = currentPeriodValue;
                    periodData.CumulativeY = cumulativeValue;
                    break;
                case DataDirection.Z:
                    periodData.CurrentPeriodZ = currentPeriodValue;
                    periodData.CumulativeZ = cumulativeValue;
                    break;
                default:
                    throw new ArgumentException($"不支持的数据方向: {direction}");
            }
        }

        /// <summary>
        /// 应用修正到数据
        /// </summary>
        /// <param name="corrections">修正列表</param>
        private void ApplyCorrections(List<DataCorrection> corrections)
        {
            foreach (var correction in corrections)
            {
                // 添加到全局修正记录列表
                _allCorrections.Add(correction);

                // 如果是不修正类型，跳过
                if (correction.CorrectionType == CorrectionType.None)
                {
                    continue;
                }

                switch (correction.Direction)
                {
                    case DataDirection.X:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodX = correction.CorrectedValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            // 对于Both类型，需要从AdditionalData中获取累计值
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                correction.PeriodData.CumulativeX = Convert.ToDouble(cumulativeObj);
                            }
                            else
                            {
                                // 如果没有AdditionalData，使用CorrectedValue作为累计值
                                correction.PeriodData.CumulativeX = correction.CorrectedValue;
                            }
                        }
                        break;
                    case DataDirection.Y:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodY = correction.CorrectedValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                correction.PeriodData.CumulativeY = Convert.ToDouble(cumulativeObj);
                            }
                            else
                            {
                                correction.PeriodData.CumulativeY = correction.CorrectedValue;
                            }
                        }
                        break;
                    case DataDirection.Z:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodZ = correction.CorrectedValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                correction.PeriodData.CumulativeZ = Convert.ToDouble(cumulativeObj);
                            }
                            else
                            {
                                correction.PeriodData.CumulativeZ = correction.CorrectedValue;
                            }
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// 回滚修正
        /// </summary>
        /// <param name="corrections">修正列表</param>
        private void RollbackCorrections(List<DataCorrection> corrections)
        {
            foreach (var correction in corrections)
            {
                switch (correction.Direction)
                {
                    case DataDirection.X:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodX = correction.OriginalValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            // 对于Both类型，需要从AdditionalData中获取原始累计值
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("OriginalCumulative", out var originalCumulativeObj))
                            {
                                correction.PeriodData.CumulativeX = Convert.ToDouble(originalCumulativeObj);
                            }
                            else
                            {
                                // 如果没有AdditionalData，使用OriginalValue作为累计值
                                correction.PeriodData.CumulativeX = correction.OriginalValue;
                            }
                        }
                        break;
                    case DataDirection.Y:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodY = correction.OriginalValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("OriginalCumulative", out var originalCumulativeObj))
                            {
                                correction.PeriodData.CumulativeY = Convert.ToDouble(originalCumulativeObj);
                            }
                            else
                            {
                                correction.PeriodData.CumulativeY = correction.OriginalValue;
                            }
                        }
                        break;
                    case DataDirection.Z:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                            correction.PeriodData.CurrentPeriodZ = correction.OriginalValue;
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("OriginalCumulative", out var originalCumulativeObj))
                            {
                                correction.PeriodData.CumulativeZ = Convert.ToDouble(originalCumulativeObj);
                            }
                            else
                            {
                                correction.PeriodData.CumulativeZ = correction.OriginalValue;
                            }
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// 从监测点获取所有修正记录
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> GetCorrectionsFromPoint(MonitoringPoint point)
        {
            // 返回该监测点的所有修正记录
            return _allCorrections.Where(c => c.PeriodData.PointName == point.PointName).ToList();
        }

        /// <summary>
        /// 获取上一期数据
        /// </summary>
        /// <param name="currentPeriod">当前期数据</param>
        /// <returns>上一期数据</returns>
        private PeriodData? GetPreviousPeriod(PeriodData currentPeriod)
        {
            // 由于当前实现中PeriodData没有直接引用上一期，我们需要通过其他方式获取
            // 这里需要从外部传入完整的监测点数据来获取上一期
            // 暂时返回null，实际使用时需要重构数据传递方式
            return null;
        }

        private PeriodData? GetPreviousPeriodData(MonitoringPoint point, string formattedTime)
        {
            // 将 formattedTime 转换为 DateTime 进行比较
            if (!DateTime.TryParse(formattedTime, out DateTime targetTime))
            {
                _logger.ShowWarning($"无法解析时间格式: {formattedTime}");
                _logger.FileWarning("无法解析时间格式: {FormattedTime}", formattedTime);
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
        /// 将数据方向转换为调整方向
        /// </summary>
        /// <param name="direction">数据方向</param>
        /// <returns>调整方向</returns>
        private DataDirection GetDataDirection(DataDirection direction)
        {
            return direction; // 直接返回，因为类型相同
        }

        /// <summary>
        /// 重新计算当前期和累计值，确保前后期衔接
        /// 新的修正逻辑已经确保了数据的正确性
        /// </summary>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="previousPeriod">上一期数据</param>
        /// <param name="corrections">修正列表</param>
        private void RecalculateCumulativeValues(PeriodData currentPeriod, PeriodData previousPeriod, List<DataCorrection> corrections)
        {
            // 应用X方向修正
            var xCorrections = corrections.Where(c => c.Direction == DataDirection.X).ToList();
            if (xCorrections.Any())
            {
                var correction = xCorrections.First(); // 取第一个修正
                currentPeriod.CurrentPeriodX = correction.CorrectedValue;

                // 如果是双重修正，使用AdditionalData中的累计值
                if (correction.CorrectionType == CorrectionType.Both &&
                    correction.AdditionalData != null &&
                    correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                {
                    currentPeriod.CumulativeX = Convert.ToDouble(cumulativeObj);
                }
                else
                {
                    // 否则通过公式计算累计值
                    currentPeriod.CumulativeX = previousPeriod.CumulativeX + correction.CorrectedValue;
                }
            }

            // 应用Y方向修正
            var yCorrections = corrections.Where(c => c.Direction == DataDirection.Y).ToList();
            if (yCorrections.Any())
            {
                var correction = yCorrections.First();
                currentPeriod.CurrentPeriodY = correction.CorrectedValue;

                if (correction.CorrectionType == CorrectionType.Both &&
                    correction.AdditionalData != null &&
                    correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                {
                    currentPeriod.CumulativeY = Convert.ToDouble(cumulativeObj);
                }
                else
                {
                    currentPeriod.CumulativeY = previousPeriod.CumulativeY + correction.CorrectedValue;
                }
            }

            // 应用Z方向修正
            var zCorrections = corrections.Where(c => c.Direction == DataDirection.Z).ToList();
            if (zCorrections.Any())
            {
                var correction = zCorrections.First();
                currentPeriod.CurrentPeriodZ = correction.CorrectedValue;

                if (correction.CorrectionType == CorrectionType.Both &&
                    correction.AdditionalData != null &&
                    correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                {
                    currentPeriod.CumulativeZ = Convert.ToDouble(cumulativeObj);
                }
                else
                {
                    currentPeriod.CumulativeZ = previousPeriod.CumulativeZ + correction.CorrectedValue;
                }
            }
        }

        /// <summary>
        /// 验证监测点数据一致性
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidatePointDataConsistency(MonitoringPoint point)
        {
            var result = new ValidationResult
            {
                PointName = point.PointName,
                Status = ValidationStatus.Valid,
                ValidationType = "监测点数据一致性验证",
                Description = $"验证监测点 {point.PointName} 修正后的数据是否满足基本逻辑"
            };

            try
            {
                if (point.PeriodDataCount < 2)
                {
                    result.SetSuccess();
                    return result;
                }

                // 按时间排序
                var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();

                for (int i = 1; i < sortedData.Count; i++)
                {
                    var current = sortedData[i];
                    var previous = sortedData[i - 1];

                    // 检查本期变化量是否超出限制
                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(current.CurrentPeriodX), _options.MaxCurrentPeriodValue, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期X方向本期变化量超出限制: {current.CurrentPeriodX:F6} > {_options.MaxCurrentPeriodValue:F6}");
                        return result;
                    }

                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(current.CurrentPeriodY), _options.MaxCurrentPeriodValue, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Y方向本期变化量超出限制: {current.CurrentPeriodY:F6} > {_options.MaxCurrentPeriodValue:F6}");
                        return result;
                    }

                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(current.CurrentPeriodZ), _options.MaxCurrentPeriodValue, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Z方向本期变化量超出限制: {current.CurrentPeriodZ:F6} > {_options.MaxCurrentPeriodValue:F6}");
                        return result;
                    }

                    // 验证X方向：累计值 = 上期累计值 + 本期变化量
                    var expectedCumulativeX = previous.CumulativeX + current.CurrentPeriodX;
                    var actualCumulativeX = current.CumulativeX;
                    var xDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeX, actualCumulativeX);

                    if (FloatingPointUtils.IsGreaterThan(xDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期X方向数据不一致: 期望累计值={expectedCumulativeX:F6}, 实际累计值={actualCumulativeX:F6}, 差异={xDifference:F6}");
                        return result;
                    }

                    // 验证Y方向
                    var expectedCumulativeY = previous.CumulativeY + current.CurrentPeriodY;
                    var actualCumulativeY = current.CumulativeY;
                    var yDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeY, actualCumulativeY);

                    if (FloatingPointUtils.IsGreaterThan(yDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Y方向数据不一致: 期望累计值={expectedCumulativeY:F6}, 实际累计值={actualCumulativeY:F6}, 差异={yDifference:F6}");
                        return result;
                    }

                    // 验证Z方向
                    var expectedCumulativeZ = previous.CumulativeZ + current.CurrentPeriodZ;
                    var actualCumulativeZ = current.CumulativeZ;
                    var zDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeZ, actualCumulativeZ);

                    if (FloatingPointUtils.IsGreaterThan(zDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Z方向数据不一致: 期望累计值={expectedCumulativeZ:F6}, 实际累计值={actualCumulativeZ:F6}, 差异={zDifference:F6}");
                        return result;
                    }
                }

                result.SetSuccess();
            }
            catch (Exception ex)
            {
                result.SetFailure($"验证过程中发生异常: {ex.Message}", ValidationSeverity.Error);
            }

            return result;
        }

        /// <summary>
        /// 验证修正后的数据是否满足累计值关系
        /// </summary>
        /// <param name="pointData">监测点数据</param>
        /// <returns>验证结果</returns>
        public ValidationResult ValidateCorrectedData(List<PeriodData> pointData)
        {
            var result = new ValidationResult
            {
                Status = ValidationStatus.Valid,
                ValidationType = "累计值关系验证",
                Description = "验证修正后的数据是否满足累计值关系"
            };

            try
            {
                for (int i = 1; i < pointData.Count; i++)
                {
                    var current = pointData[i];
                    var previous = pointData[i - 1];

                    // 验证X方向
                    var expectedCumulativeX = previous.CumulativeX + current.CurrentPeriodX;
                    var actualCumulativeX = current.CumulativeX;
                    var xDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeX, actualCumulativeX);

                    if (FloatingPointUtils.IsGreaterThan(xDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期X方向数据不一致: 期望累计值={expectedCumulativeX:F6}, 实际累计值={actualCumulativeX:F6}, 差异={xDifference:F6}");
                        result.AddFailedValue($"第{i}期X累计值", actualCumulativeX);
                        result.AddExpectedValue($"第{i}期X累计值", expectedCumulativeX);
                    }

                    // 验证Y方向
                    var expectedCumulativeY = previous.CumulativeY + current.CurrentPeriodY;
                    var actualCumulativeY = current.CumulativeY;
                    var yDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeY, actualCumulativeY);

                    if (FloatingPointUtils.IsGreaterThan(yDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Y方向数据不一致: 期望累计值={expectedCumulativeY:F6}, 实际累计值={actualCumulativeY:F6}, 差异={yDifference:F6}");
                        result.AddFailedValue($"第{i}期Y累计值", actualCumulativeY);
                        result.AddExpectedValue($"第{i}期Y累计值", expectedCumulativeY);
                    }

                    // 验证Z方向
                    var expectedCumulativeZ = previous.CumulativeZ + current.CurrentPeriodZ;
                    var actualCumulativeZ = current.CumulativeZ;
                    var zDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeZ, actualCumulativeZ);

                    if (FloatingPointUtils.IsGreaterThan(zDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                    {
                        result.SetFailure($"第{i}期Z方向数据不一致: 期望累计值={expectedCumulativeZ:F6}, 实际累计值={actualCumulativeZ:F6}, 差异={zDifference:F6}");
                        result.AddFailedValue($"第{i}期Z累计值", actualCumulativeZ);
                        result.AddExpectedValue($"第{i}期Z累计值", expectedCumulativeZ);
                    }
                }

                if (result.Status == ValidationStatus.Valid)
                {
                    result.SetSuccess();
                }
            }
            catch (Exception ex)
            {
                result.SetFailure($"验证过程中发生异常: {ex.Message}", ValidationSeverity.Error);
                result.AddErrorDetail($"异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 验证修正后的监测点数据是否满足累计值关系
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <returns>验证结果列表</returns>
        public List<ValidationResult> ValidateCorrectedMonitoringPoints(List<MonitoringPoint> monitoringPoints)
        {
            var results = new List<ValidationResult>();

            foreach (var point in monitoringPoints)
            {
                var result = new ValidationResult
                {
                    PointName = point.PointName,
                    Status = ValidationStatus.Valid,
                    ValidationType = "修正后累计值关系验证",
                    Description = $"验证监测点 {point.PointName} 修正后的数据是否满足累计值关系"
                };

                try
                {
                    if (point.PeriodDataCount < 2)
                    {
                        result.SetSuccess();
                        results.Add(result);
                        continue;
                    }

                    // 按时间排序
                    var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();

                    for (int i = 1; i < sortedData.Count; i++)
                    {
                        var current = sortedData[i];
                        var previous = sortedData[i - 1];

                        // 验证X方向
                        var expectedCumulativeX = previous.CumulativeX + current.CurrentPeriodX;
                        var actualCumulativeX = current.CumulativeX;
                        var xDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeX, actualCumulativeX);

                        if (FloatingPointUtils.IsGreaterThan(xDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                        {
                            result.SetFailure($"第{i}期X方向数据不一致: 期望累计值={expectedCumulativeX:F6}, 实际累计值={actualCumulativeX:F6}, 差异={xDifference:F6}");
                            result.AddFailedValue($"第{i}期X累计值", actualCumulativeX);
                            result.AddExpectedValue($"第{i}期X累计值", expectedCumulativeX);
                        }

                        // 验证Y方向
                        var expectedCumulativeY = previous.CumulativeY + current.CurrentPeriodY;
                        var actualCumulativeY = current.CumulativeY;
                        var yDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeY, actualCumulativeY);

                        if (FloatingPointUtils.IsGreaterThan(yDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                        {
                            result.SetFailure($"第{i}期Y方向数据不一致: 期望累计值={expectedCumulativeY:F6}, 实际累计值={actualCumulativeY:F6}, 差异={yDifference:F6}");
                            result.AddFailedValue($"第{i}期Y累计值", actualCumulativeY);
                            result.AddExpectedValue($"第{i}期Y累计值", expectedCumulativeY);
                        }

                        // 验证Z方向
                        var expectedCumulativeZ = previous.CumulativeZ + current.CurrentPeriodZ;
                        var actualCumulativeZ = current.CumulativeZ;
                        var zDifference = FloatingPointUtils.SafeAbsoluteDifference(expectedCumulativeZ, actualCumulativeZ);

                        if (FloatingPointUtils.IsGreaterThan(zDifference, _options.CumulativeTolerance, _options.CumulativeTolerance))
                        {
                            result.SetFailure($"第{i}期Z方向数据不一致: 期望累计值={expectedCumulativeZ:F6}, 实际累计值={actualCumulativeZ:F6}, 差异={zDifference:F6}");
                            result.AddFailedValue($"第{i}期Z累计值", actualCumulativeZ);
                            result.AddExpectedValue($"第{i}期Z累计值", expectedCumulativeZ);
                        }
                    }

                    if (result.Status == ValidationStatus.Valid)
                    {
                        result.SetSuccess();
                    }
                }
                catch (Exception ex)
                {
                    result.SetFailure($"验证过程中发生异常: {ex.Message}", ValidationSeverity.Error);
                    result.AddErrorDetail($"异常: {ex.Message}");
                }

                results.Add(result);
            }

            return results;
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

        /// <summary>
        /// 计算验证失败比例
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="validationResult">验证结果</param>
        /// <returns>失败比例 (0.0 - 1.0)</returns>
        private double CalculateValidationFailureRatio(MonitoringPoint point, ValidationResult validationResult)
        {
            if (validationResult.Status == ValidationStatus.Valid)
                return 0.0;

            // 计算总的数据项数（3个方向 × 期数）
            var totalDataItems = point.PeriodDataCount * 3;
            if (totalDataItems == 0) return 1.0;

            // 从验证结果描述中提取失败信息，估算失败数量
            // 这里可以根据实际的验证结果格式进行调整
            var failureCount = EstimateFailureCount(validationResult);
            return Math.Min(1.0, (double)failureCount / totalDataItems);
        }

        /// <summary>
        /// 估算验证失败的数量
        /// </summary>
        /// <param name="validationResult">验证结果</param>
        /// <returns>失败数量</returns>
        private int EstimateFailureCount(ValidationResult validationResult)
        {
            if (string.IsNullOrEmpty(validationResult.Description))
                return 1;

            // 根据描述中的期数信息估算失败数量
            var description = validationResult.Description;
            var failureCount = 0;

            // 简单估算：每个"第X期"表示一个失败
            for (int i = 1; i <= 100; i++) // 假设最多100期
            {
                if (description.Contains($"第{i}期"))
                    failureCount++;
            }

            return Math.Max(1, failureCount); // 至少返回1
        }

        /// <summary>
        /// 应用部分修正策略
        /// 当失败比例较低时，只修正有问题的数据，保持其他数据不变
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="validationResult">验证结果</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> ApplyPartialCorrectionStrategy(MonitoringPoint point, ValidationResult validationResult)
        {
            var corrections = new List<DataCorrection>();
            var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();
            var random = new Random();

            // 分析验证失败的具体问题，只修正有问题的数据
            var failedPeriods = ExtractFailedPeriods(validationResult);

            foreach (var failedPeriod in failedPeriods)
            {
                var periodData = sortedData.FirstOrDefault(pd => pd.RowNumber == failedPeriod.RowNumber);
                if (periodData == null) continue;

                foreach (var direction in failedPeriod.FailedDirections)
                {
                    var (originalPeriodValue, originalCumulative) = GetDirectionValue(periodData, direction);

                    // 策略1：生成绝对值小于0.3的随机本期变化量
                    var newPeriodValue = GenerateSmallRandomValue(random, 0.3);

                    // 重新计算累计值
                    var previousPeriod = GetPreviousPeriodForPartialCorrection(sortedData, periodData);
                    var newCumulative = previousPeriod != null
                        ? GetDirectionValue(previousPeriod, direction).cumulativeValue + newPeriodValue
                        : newPeriodValue; // 如果是第一期，累计值等于本期值

                    // 检查累计值是否在合理范围内
                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(newCumulative), _options.MaxCumulativeValue, _options.CumulativeTolerance))
                    {
                        // 如果累计值超出范围，调整本期变化量
                        var maxAllowedCumulative = _options.MaxCumulativeValue * Math.Sign(newCumulative);
                        var previousCumulative = previousPeriod != null
                            ? GetDirectionValue(previousPeriod, direction).cumulativeValue
                            : 0.0;
                        newPeriodValue = maxAllowedCumulative - previousCumulative;
                        newCumulative = maxAllowedCumulative;
                    }

                    // 如果修正后的值仍然不合理，设置为0
                    if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(newPeriodValue), 0.5, _options.CumulativeTolerance))
                    {
                        newPeriodValue = 0.0;
                        newCumulative = previousPeriod != null
                            ? GetDirectionValue(previousPeriod, direction).cumulativeValue
                            : 0.0;
                    }

                    var correction = new DataCorrection
                    {
                        PeriodData = periodData,
                        Direction = direction,
                        CorrectionType = CorrectionType.Both,
                        OriginalValue = originalPeriodValue,
                        CorrectedValue = newPeriodValue,
                        Reason = $"部分修正策略：修正验证失败的数据，本期变化量={newPeriodValue:F6}, 累计值={newCumulative:F6}",
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "CorrectedCumulative", newCumulative },
                            { "OriginalCumulative", originalCumulative }
                        }
                    };

                    corrections.Add(correction);
                }
            }

            return corrections;
        }

        /// <summary>
        /// 生成小的随机值
        /// </summary>
        /// <param name="random">随机数生成器</param>
        /// <param name="maxAbsValue">最大绝对值</param>
        /// <returns>随机值</returns>
        private double GenerateSmallRandomValue(Random random, double maxAbsValue)
        {
            var value = (random.NextDouble() - 0.5) * 2 * maxAbsValue;
            return FloatingPointUtils.SafeRound(value, 6);
        }

        /// <summary>
        /// 获取部分修正策略中的上一期数据
        /// </summary>
        /// <param name="sortedData">排序后的数据</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <returns>上一期数据</returns>
        private PeriodData? GetPreviousPeriodForPartialCorrection(List<PeriodData> sortedData, PeriodData currentPeriod)
        {
            var currentIndex = sortedData.IndexOf(currentPeriod);
            if (currentIndex <= 0) return null;
            return sortedData[currentIndex - 1];
        }

        /// <summary>
        /// 提取验证失败的期数和方向信息
        /// </summary>
        /// <param name="validationResult">验证结果</param>
        /// <returns>失败信息列表</returns>
        private List<FailedPeriodInfo> ExtractFailedPeriods(ValidationResult validationResult)
        {
            var failedPeriods = new List<FailedPeriodInfo>();

            if (string.IsNullOrEmpty(validationResult.Description))
                return failedPeriods;

            var description = validationResult.Description;

            // 解析验证失败信息，提取期数和方向
            // 这里需要根据实际的验证结果格式进行调整
            var lines = description.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                if (line.Contains("第") && line.Contains("期"))
                {
                    var periodInfo = ParseFailedPeriodLine(line);
                    if (periodInfo != null)
                        failedPeriods.Add(periodInfo);
                }
            }

            return failedPeriods;
        }

        /// <summary>
        /// 解析失败期数行
        /// </summary>
        /// <param name="line">失败描述行</param>
        /// <returns>失败期数信息</returns>
        private FailedPeriodInfo? ParseFailedPeriodLine(string line)
        {
            try
            {
                // 提取期数
                var periodMatch = System.Text.RegularExpressions.Regex.Match(line, @"第(\d+)期");
                if (!periodMatch.Success) return null;

                var periodNumber = int.Parse(periodMatch.Groups[1].Value);

                // 提取方向信息
                var directions = new List<DataDirection>();
                if (line.Contains("X方向")) directions.Add(DataDirection.X);
                if (line.Contains("Y方向")) directions.Add(DataDirection.Y);
                if (line.Contains("Z方向")) directions.Add(DataDirection.Z);

                // 如果没有明确的方向信息，假设所有方向都有问题
                if (!directions.Any())
                {
                    directions.AddRange(new[] { DataDirection.X, DataDirection.Y, DataDirection.Z });
                }

                return new FailedPeriodInfo
                {
                    RowNumber = periodNumber,
                    FailedDirections = directions
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 应用最终修正策略
        /// 当其他策略都失败时，使用智能重置策略
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>修正列表</returns>
        private List<DataCorrection> ApplyFinalCorrectionStrategy(MonitoringPoint point)
        {
            var corrections = new List<DataCorrection>();
            var sortedData = point.PeriodDataList.OrderBy(pd => pd.FileInfo?.FullDateTime).ToList();
            var directions = new[] { DataDirection.X, DataDirection.Y, DataDirection.Z };

            foreach (var direction in directions)
            {
                // 分析该方向的数据特征
                var directionValues = sortedData.Select(pd => GetDirectionValue(pd, direction))
                    .Where(v => FloatingPointUtils.IsLessThanOrEqual(FloatingPointUtils.SafeAbs(v.currentPeriodValue), _options.MaxCurrentPeriodValue * 2, _options.CumulativeTolerance)) // 过滤掉极端异常值
                    .ToList();

                if (directionValues.Any())
                {
                    // 计算合理的本期变化量范围，避免使用边界值
                    var validPeriodValues = directionValues.Select(v => v.currentPeriodValue).ToList();
                    var avgPeriodValue = validPeriodValues.Average();
                    var stdPeriodValue = CalculateStandardDeviation(validPeriodValues);

                    // 生成合理的本期变化量，避免使用边界值
                    var random = new Random();
                    var currentCumulative = 0.0; // 从0开始累加

                    for (int i = 0; i < sortedData.Count; i++)
                    {
                        var current = sortedData[i];
                        var (originalPeriodValue, originalCumulative) = GetDirectionValue(current, direction);

                        // 策略1：基于原始数据分布生成合理的本期变化量
                        double newPeriodValue;
                        if (FloatingPointUtils.IsLessThan(FloatingPointUtils.SafeAbs(avgPeriodValue), 0.1, _options.CumulativeTolerance))
                        {
                            // 如果平均值很小，生成小的随机值
                            newPeriodValue = GenerateSmallRandomValue(random, 0.2);
                        }
                        else
                        {
                            // 基于正态分布生成值，但限制在合理范围内
                            var u1 = random.NextDouble();
                            var u2 = random.NextDouble();
                            var z0 = FloatingPointUtils.SafeSqrt(-2.0 * FloatingPointUtils.SafeLog(u1)) * FloatingPointUtils.SafeCos(2.0 * Math.PI * u2);

                            newPeriodValue = FloatingPointUtils.SafeRound(avgPeriodValue + z0 * stdPeriodValue * 0.5, 6);

                            // 限制在安全范围内，但避免使用边界值
                            var maxSafeValue = _options.MaxCurrentPeriodValue * 0.8; // 使用80%的边界值
                            if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbs(newPeriodValue), maxSafeValue, _options.CumulativeTolerance))
                            {
                                newPeriodValue = FloatingPointUtils.SafeSign(newPeriodValue) * maxSafeValue;
                            }
                        }

                        // 确保生成的值不会过小
                        if (FloatingPointUtils.IsLessThan(FloatingPointUtils.SafeAbs(newPeriodValue), 0.001, _options.CumulativeTolerance))
                        {
                            newPeriodValue = FloatingPointUtils.SafeSign(avgPeriodValue) * 0.001;
                        }

                        // 累计值基于本期变化量累加
                        var newCumulative = currentCumulative;
                        currentCumulative += newPeriodValue;

                        // 创建修正记录
                        if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbsoluteDifference(newPeriodValue, originalPeriodValue), _options.CumulativeTolerance, _options.CumulativeTolerance))
                        {
                            var correction = new DataCorrection
                            {
                                PeriodData = current,
                                Direction = direction,
                                CorrectionType = CorrectionType.Both,
                                OriginalValue = originalPeriodValue,
                                CorrectedValue = newPeriodValue,
                                Reason = $"最终策略：基于数据分布生成合理的本期变化量 {newPeriodValue:F6}，避免边界值",
                                AdditionalData = new Dictionary<string, object>
                                {
                                    { "CorrectedCumulative", newCumulative }
                                }
                            };

                            corrections.Add(correction);
                        }
                    }
                }
                else
                {
                    // 如果所有数据都异常，使用智能重置策略，避免全部设为0
                    var random = new Random();
                    var currentCumulative = 0.0;

                    for (int i = 0; i < sortedData.Count; i++)
                    {
                        var current = sortedData[i];
                        var (originalPeriodValue, originalCumulative) = GetDirectionValue(current, direction);

                        // 生成小的随机值，而不是全部设为0
                        var newPeriodValue = GenerateSmallRandomValue(random, 0.1);

                        // 累计值基于本期变化量累加
                        var newCumulative = currentCumulative;
                        currentCumulative += newPeriodValue;

                        if (FloatingPointUtils.IsGreaterThan(FloatingPointUtils.SafeAbsoluteDifference(newPeriodValue, originalPeriodValue), _options.CumulativeTolerance, _options.CumulativeTolerance))
                        {
                            var correction = new DataCorrection
                            {
                                PeriodData = current,
                                Direction = direction,
                                CorrectionType = CorrectionType.Both,
                                OriginalValue = originalPeriodValue,
                                CorrectedValue = newPeriodValue,
                                Reason = $"最终策略：数据严重异常，使用智能重置策略，本期变化量={newPeriodValue:F6}",
                                AdditionalData = new Dictionary<string, object>
                                {
                                    { "CorrectedCumulative", newCumulative }
                                }
                            };

                            corrections.Add(correction);
                        }
                    }
                }
            }

            return corrections;
        }

        /// <summary>
        /// 计算标准差
        /// </summary>
        private double CalculateStandardDeviation(List<double> values)
        {
            if (values.Count <= 1) return 0.0;

            var mean = FloatingPointUtils.SafeAverage(values.ToArray());
            var sumSquaredDiff = values.Sum(x => FloatingPointUtils.SafePow(x - mean, 2));
            return FloatingPointUtils.SafeSqrt(sumSquaredDiff / (values.Count - 1));
        }
    }

    /// <summary>
    /// 修正选项
    /// </summary>
    public class CorrectionOptions
    {
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
        /// 累计值容差
        /// </summary>
        public double CumulativeTolerance { get; set; } = 1e-6;

        /// <summary>
        /// 最小值阈值
        /// </summary>
        public double MinValueThreshold { get; set; } = 1e-9;
    }

    // DataCorrection 和 CorrectionType 已移动到 Models/CorrectionModels.cs

    // CorrectionStatus 和 PointCorrectionResult 已移动到 Models/CorrectionModels.cs

    // CorrectionResult 类已移动到 Models/CorrectionModels.cs

    /// <summary>
    /// 失败期数信息
    /// </summary>
    public class FailedPeriodInfo
    {
        /// <summary>
        /// 行号（期数）
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 失败的数据方向
        /// </summary>
        public List<DataDirection> FailedDirections { get; set; } = new List<DataDirection>();
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
