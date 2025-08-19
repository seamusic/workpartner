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
                                //var perviousData1 = GetPreviousPeriodData(point, currentResult[0].FormattedTime);
                                // 检查是否可以修正
                                var canFix = currentResult.Any(x => x.PointName == point.PointName && x.CanAdjustment);
                                if (canFix)
                                {
                                    foreach (ValidationResult validationResult in currentResult)
                                    {
                                        foreach (PeriodData periodData in point.PeriodDataList)
                                        {
                                            if (periodData.PointName == validationResult.PointName && periodData.FormattedTime == validationResult.FormattedTime)
                                            {
                                                periodData.CanAdjustment = true;
                                            }
                                        }
                                    }
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

                // 检查是否再进行修正
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
            
            // 按时间顺序分析数据，识别可调整区间
            var adjustableSegments = FindAdjustableSegments(sortedData);
            
            foreach (var segment in adjustableSegments)
            {
                var segmentCorrections = CorrectSegmentWithCumulativeLogic(sortedData, segment, direction);
                corrections.AddRange(segmentCorrections);
            }
            
            return corrections;
        }
        
        /// <summary>
        /// 使用累计值逻辑修正区间
        /// </summary>
        private List<DataCorrection> CorrectSegmentWithCumulativeLogic(List<PeriodData> sortedData, AdjustableSegment segment, DataDirection direction)
        {
            var corrections = new List<DataCorrection>();
            var random = new Random();
            var maxRetryAttempts = 20; // 增加尝试次数
            var bestCorrections = new List<DataCorrection>();
            var bestError = double.MaxValue;
            
            // 获取区间边界值
            var previousCumulative = segment.PreviousFixedIndex >= 0 
                ? GetDirectionValue(sortedData[segment.PreviousFixedIndex], direction).cumulativeValue 
                : 0.0;
            
            var nextCumulative = segment.NextFixedIndex >= 0 
                ? GetDirectionValue(sortedData[segment.NextFixedIndex], direction).cumulativeValue 
                : double.NaN;
            
            // 尝试多次随机生成，选择最佳结果
            for (int attempt = 0; attempt < maxRetryAttempts; attempt++)
            {
                var attemptCorrections = new List<DataCorrection>();
                var currentCumulative = previousCumulative;
                
                // 如果区间内只有一期数据，随机生成本期变化量
                if (segment.StartIndex == segment.EndIndex)
                {
                    var current = sortedData[segment.StartIndex];
                    var (originalPeriodValue, originalCumulative) = GetDirectionValue(current, direction);
                    
                    // 计算期望的本期变化量
                    var expectedPeriodValue = !double.IsNaN(nextCumulative) 
                        ? nextCumulative - previousCumulative 
                        : originalPeriodValue;
                    
                    // 随机生成本期变化量，但优先考虑能改善累计值关系的值
                    var newPeriodValue = (random.NextDouble() - 0.5) * (2 * _options.RandomChangeRange); // 范围：-RandomChangeRange 到 +RandomChangeRange
                    
                    // 如果有下一期固定值，优先选择能改善累计值关系的值
                    if (!double.IsNaN(nextCumulative))
                    {
                        // 计算当前随机值的累计值差异
                        var randomCumulative = previousCumulative + newPeriodValue;
                        var randomError = Math.Abs(randomCumulative - nextCumulative);
                        
                        // 计算保持原值的累计值差异
                        var originalCumulativeValue = previousCumulative + originalPeriodValue;
                        var originalError = Math.Abs(originalCumulativeValue - nextCumulative);
                        
                        // 如果随机值能改善累计值关系，使用它；否则保持原值
                        if (randomError < originalError)
                        {
                            // 使用随机值，但确保累计值关系
                            newPeriodValue = expectedPeriodValue;
                        }
                        else
                        {
                            // 保持原值，只调整累计值
                            newPeriodValue = originalPeriodValue;
                        }
                    }
                    
                    var correctCumulative = previousCumulative + newPeriodValue;
                    
                    var correction = new DataCorrection
                    {
                        PeriodData = current,
                        Direction = direction,
                        CorrectionType = CorrectionType.Both,
                        OriginalValue = originalPeriodValue,
                        CorrectedValue = newPeriodValue,
                        Reason = $"第{attempt + 1}次尝试：智能选择本期变化量，确保累计值关系正确",
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "CorrectedCumulative", correctCumulative },
                            { "Attempt", attempt + 1 },
                            { "AdjustmentType", "SmartSelection" }
                        }
                    };
                    attemptCorrections.Add(correction);
                    currentCumulative = correctCumulative;
                }
                else
                {
                    // 多期数据，随机生成本期变化量
                    var segmentLength = segment.EndIndex - segment.StartIndex + 1;
                    var periodValues = new List<double>();
                    
                    for (int i = 0; i < segmentLength; i++)
                    {
                        var current = sortedData[segment.StartIndex + i];
                        var (originalPeriodValue, _) = GetDirectionValue(current, direction);
                        
                        // 随机生成本期变化量，保持波动性
                        var newPeriodValue = (random.NextDouble() - 0.5) * (2 * _options.RandomChangeRange); // 范围：-RandomChangeRange 到 +RandomChangeRange
                        
                        // 对于最后一期，如果有下一期固定值，调整确保累计值关系
                        if (i == segmentLength - 1 && !double.IsNaN(nextCumulative))
                        {
                            var remainingChange = nextCumulative - currentCumulative;
                            // 如果随机值在合理范围内，使用它；否则调整
                            if (Math.Abs(newPeriodValue - remainingChange) <= 0.5)
                            {
                                newPeriodValue = remainingChange;
                            }
                            else
                            {
                                // 保持随机性，但确保累计值关系
                                newPeriodValue = remainingChange;
                            }
                        }
                        
                        // 确保本期变化量在合理范围内
                        newPeriodValue = Math.Max(-1.0, Math.Min(1.0, newPeriodValue));
                        
                        periodValues.Add(newPeriodValue);
                        currentCumulative += newPeriodValue;
                    }
                    
                    // 应用修正
                    currentCumulative = previousCumulative;
                    for (int i = 0; i < periodValues.Count; i++)
                    {
                        var current = sortedData[segment.StartIndex + i];
                        var (originalPeriodValue, originalCumulative) = GetDirectionValue(current, direction);
                        var newPeriodValue = periodValues[i];
                        
                        currentCumulative += newPeriodValue;
                        
                        var correction = new DataCorrection
                        {
                            PeriodData = current,
                            Direction = direction,
                            CorrectionType = CorrectionType.Both,
                            OriginalValue = originalPeriodValue,
                            CorrectedValue = newPeriodValue,
                            Reason = $"第{attempt + 1}次尝试：随机生成本期变化量，确保累计值关系正确",
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "CorrectedCumulative", currentCumulative },
                                { "Attempt", attempt + 1 },
                                { "AdjustmentType", "RandomGeneration" }
                            }
                        };
                        attemptCorrections.Add(correction);
                    }
                }
                
                // 计算区间误差
                var segmentError = 0.0;
                if (!double.IsNaN(nextCumulative))
                {
                    // 有下一期固定值，计算与目标累计值的差异
                    segmentError = Math.Abs(currentCumulative - nextCumulative);
                }
                else
                {
                    // 没有下一期固定值，只检查区间内部的累计值关系
                    segmentError = 0.0;
                    for (int i = segment.StartIndex + 1; i <= segment.EndIndex; i++)
                    {
                        var current = sortedData[i];
                        var previous = sortedData[i - 1];
                        var (currentPeriodValue, _) = GetDirectionValue(current, direction);
                        var (_, prevCumulative) = GetDirectionValue(previous, direction);
                        
                        var expectedCumulative = prevCumulative + currentPeriodValue;
                        var actualCumulative = GetDirectionValue(current, direction).cumulativeValue;
                        segmentError += Math.Abs(expectedCumulative - actualCumulative);
                    }
                }
                
                // 如果这次尝试的结果更好，保存它
                if (segmentError < bestError)
                {
                    bestError = segmentError;
                    bestCorrections = attemptCorrections.ToList();
                }
                
                // 如果误差已经很小，可以提前退出
                if (segmentError < _options.CumulativeTolerance)
                {
                    break;
                }
            }
            
            corrections.AddRange(bestCorrections);
            return corrections;
        }
        
        /// <summary>
        /// 查找可调整的数据区间
        /// </summary>
        private List<AdjustableSegment> FindAdjustableSegments(List<PeriodData> sortedData)
        {
            var segments = new List<AdjustableSegment>();
            var startIndex = -1;
            
            for (int i = 0; i < sortedData.Count; i++)
            {
                var current = sortedData[i];
                
                if (current.CanAdjustment && startIndex == -1)
                {
                    // 找到可调整区间的开始
                    startIndex = i;
                }
                else if (!current.CanAdjustment && startIndex != -1)
                {
                    // 找到可调整区间的结束
                    segments.Add(new AdjustableSegment
                    {
                        StartIndex = startIndex,
                        EndIndex = i - 1,
                        PreviousFixedIndex = startIndex > 0 ? startIndex - 1 : -1,
                        NextFixedIndex = i
                    });
                    startIndex = -1;
                }
            }
            
            // 处理最后一个可调整区间
            if (startIndex != -1)
            {
                segments.Add(new AdjustableSegment
                {
                    StartIndex = startIndex,
                    EndIndex = sortedData.Count - 1,
                    PreviousFixedIndex = startIndex > 0 ? startIndex - 1 : -1,
                    NextFixedIndex = -1
                });
            }
            
            return segments;
        }
        
        /// <summary>
        /// 可调整数据区间的信息
        /// </summary>
        private class AdjustableSegment
        {
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
            public int PreviousFixedIndex { get; set; } // 前一个CanAdjustment=false的索引
            public int NextFixedIndex { get; set; }     // 后一个CanAdjustment=false的索引
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

                // 检查 PeriodData 是否可以修正，如果不可以则跳过
                //if (!correction.PeriodData.CanAdjustment)
                //{
                //    //_logger?.ConsoleInfo($"跳过修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 {correction.Direction} 方向，因为 CanAdjustment=false");
                //    continue;
                //}

                switch (correction.Direction)
                {
                    case DataDirection.X:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            correction.PeriodData.CurrentPeriodX = correction.CorrectedValue;
                            _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 X方向本期变化量: {correction.OriginalValue:F3} → {correction.CorrectedValue:F3}");
                        }
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            // 对于Both类型，需要从AdditionalData中获取累计值
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                var correctedCumulative = Convert.ToDouble(cumulativeObj);
                                correction.PeriodData.CumulativeX = correctedCumulative;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 X方向累计值: {correction.PeriodData.GetCumulativeValue(DataDirection.X):F3} → {correctedCumulative:F3}");
                            }
                            else
                            {
                                // 如果没有AdditionalData，使用CorrectedValue作为累计值
                                var originalCumulative = correction.PeriodData.CumulativeX;
                                correction.PeriodData.CumulativeX = correction.CorrectedValue;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 X方向累计值: {originalCumulative:F3} → {correction.CorrectedValue:F3}");
                            }
                        }
                        break;
                    case DataDirection.Y:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            correction.PeriodData.CurrentPeriodY = correction.CorrectedValue;
                            _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Y方向本期变化量: {correction.OriginalValue:F3} → {correction.CorrectedValue:F3}");
                        }
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                var correctedCumulative = Convert.ToDouble(cumulativeObj);
                                correction.PeriodData.CumulativeY = correctedCumulative;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Y方向累计值: {correction.PeriodData.GetCumulativeValue(DataDirection.Y):F3} → {correctedCumulative:F3}");
                            }
                            else
                            {
                                var originalCumulative = correction.PeriodData.CumulativeY;
                                correction.PeriodData.CumulativeY = correction.CorrectedValue;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Y方向累计值: {originalCumulative:F3} → {correction.CorrectedValue:F3}");
                            }
                        }
                        break;
                    case DataDirection.Z:
                        if (correction.CorrectionType == CorrectionType.CurrentPeriodValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            correction.PeriodData.CurrentPeriodZ = correction.CorrectedValue;
                            _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Z方向本期变化量: {correction.OriginalValue:F3} → {correction.CorrectedValue:F3}");
                        }
                        if (correction.CorrectionType == CorrectionType.CumulativeValue || correction.CorrectionType == CorrectionType.Both)
                        {
                            if (correction.CorrectionType == CorrectionType.Both && correction.AdditionalData != null &&
                                correction.AdditionalData.TryGetValue("CorrectedCumulative", out var cumulativeObj))
                            {
                                var correctedCumulative = Convert.ToDouble(cumulativeObj);
                                correction.PeriodData.CumulativeZ = correctedCumulative;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Z方向累计值: {correction.PeriodData.GetCumulativeValue(DataDirection.Z):F3} → {correctedCumulative:F3}");
                            }
                            else
                            {
                                var originalCumulative = correction.PeriodData.CumulativeZ;
                                correction.PeriodData.CumulativeZ = correction.CorrectedValue;
                                _logger?.ConsoleInfo($"修正 {correction.PeriodData.PointName} 第 {correction.PeriodData.RowNumber} 行 Z方向累计值: {originalCumulative:F3} → {correction.CorrectedValue:F3}");
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
        /// 生成小的随机值
        /// </summary>
        /// <param name="random">随机数生成器</param>
        /// <param name="maxAbsValue">最大绝对值</param>
        /// <returns>随机值</returns>
        private double GenerateSmallRandomValue(Random random, double maxAbsValue)
        {
            var value = (random.NextDouble() - 0.5) * 2 * maxAbsValue;

            // 确保生成的值不会过小，本期变化量最小值为正负0.01
            if (FloatingPointUtils.IsLessThan(FloatingPointUtils.SafeAbs(value), 0.01, 0.001))
            {
                value = 0.01 * Math.Sign(value);
            }

            return FloatingPointUtils.SafeRound(value, 6);
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
}
