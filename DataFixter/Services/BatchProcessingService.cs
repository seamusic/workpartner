using DataFixter.Models;
using DataFixter.Excel;
using DataFixter.Logging;
using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 批量处理服务，负责协调整个数据处理流程
    /// </summary>
    public class BatchProcessingService
    {
        private readonly DualLoggerService _logger;

        public BatchProcessingService(DualLoggerService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// 执行完整的批量处理流程
        /// </summary>
        /// <param name="processedDirectory">待处理目录</param>
        /// <param name="comparisonDirectory">对比目录</param>
        /// <returns>处理结果</returns>
        public ProcessingResult ProcessBatch(string processedDirectory, string comparisonDirectory)
        {
            var result = new ProcessingResult();

            try
            {
                _logger.LogOperationStart("批量处理", $"待处理目录：{processedDirectory}，对比目录：{comparisonDirectory}");

                // 步骤1: 读取Excel文件
                var (processedResults, comparisonResults) = ReadExcelFiles(processedDirectory, comparisonDirectory);

                // 步骤2: 数据标准化
                var (normalizedData, normalizedComparisonData) = NormalizeData(processedResults, comparisonResults);

                // 步骤2.5: 合并对比数据到待处理数据
                MergeComparisonData(normalizedData, normalizedComparisonData);

                // 步骤3: 数据分组和排序
                var monitoringPoints = GroupAndSortData(normalizedData);

                // 步骤4: 数据验证
                var validationResults = ValidateData(monitoringPoints, normalizedComparisonData);

                // 步骤5: 数据修正
                var correctionResult = CorrectData(monitoringPoints, validationResults);

                // 步骤5.5: 修正后重新验证数据
                var correctedValidationResults = ValidateCorrectedData(monitoringPoints);

                // 步骤5.6: 添加数值超限检查
                var limitValidationResults = ValidateValueLimits(monitoringPoints);

                // 步骤5.7: 生成数据对比报告
                GenerateDataComparisonReport(processedResults, comparisonResults, correctionResult, processedDirectory);

                if (false)
                {
                    _logger.LogOperationComplete("批量处理", "不成功",
                        $"修正后错误仍然有：{limitValidationResults.Count}条");
                }
                else
                {
                    // 步骤6: 生成输出文件
                    var (outputResult, reportResult) = GenerateOutputFiles(monitoringPoints, validationResults, correctionResult, processedDirectory);

                    // 更新结果
                    UpdateProcessingResult(result, processedResults, comparisonResults, monitoringPoints,
                        validationResults, correctionResult, outputResult, reportResult);

                    _logger.LogOperationComplete("批量处理", "成功完成",
                        $"处理文件：{processedResults.Count}个，监测点：{monitoringPoints.Count}个，修正记录：{correctionResult.AdjustmentRecords.Count}条");
                }
            }
            catch (Exception ex)
            {
                _logger.ShowError($"批量处理过程中发生异常: {ex.Message}");
                _logger.FileError(ex, "批量处理过程中发生异常");
                result.Status = ProcessingStatus.Error;
                result.Message = $"处理过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 步骤1: 读取Excel文件
        /// </summary>
        private (List<ExcelReadResult> processedResults, List<ExcelReadResult> comparisonResults) ReadExcelFiles(
            string processedDirectory, string comparisonDirectory)
        {
            _logger.ShowStep("步骤1: 读取Excel文件...");

            var excelReader = new ExcelBatchReader(processedDirectory, Log.ForContext<ExcelBatchReader>());
            var processedResults = excelReader.ReadAllFiles();

            var comparisonReader = new ExcelBatchReader(comparisonDirectory, Log.ForContext<ExcelBatchReader>());
            var comparisonResults = comparisonReader.ReadAllFiles();

            _logger.ShowComplete($"读取完成: 待处理文件 {processedResults.Count} 个, 对比文件 {comparisonResults.Count} 个");

            return (processedResults, comparisonResults);
        }

        /// <summary>
        /// 步骤2: 数据标准化
        /// </summary>
        private (List<PeriodData> normalizedData, List<PeriodData> normalizedComparisonData) NormalizeData(
            List<ExcelReadResult> processedResults, List<ExcelReadResult> comparisonResults)
        {
            _logger.ShowStep("步骤2: 数据标准化...");

            var normalizer = new DataNormalizer(Log.ForContext<DataNormalizer>());
            var normalizedData = normalizer.NormalizeData(processedResults);
            var normalizedComparisonData = normalizer.NormalizeData(comparisonResults);

            _logger.ShowComplete($"标准化完成: 待处理数据 {normalizedData.Count} 条, 对比数据 {normalizedComparisonData.Count} 条");

            return (normalizedData, normalizedComparisonData);
        }

        /// <summary>
        /// 步骤2.5: 合并对比数据到待处理数据
        /// 将对比数据中不为空的本期变化量和累计变化量覆盖到待处理数据中
        /// </summary>
        /// <param name="normalizedData">待处理数据</param>
        /// <param name="normalizedComparisonData">对比数据</param>
        private void MergeComparisonData(List<PeriodData> normalizedData, List<PeriodData> normalizedComparisonData)
        {
            _logger.ShowStep("步骤2.5: 合并对比数据到待处理数据...");

            var mergeCount = 0;
            var totalComparisonData = normalizedComparisonData.Count;

            // 创建对比数据的查找字典，以FormattedTime和PointName为键
            var comparisonDataLookup = normalizedComparisonData
                .Where(cd => !string.IsNullOrEmpty(cd.PointName))
                .ToLookup(cd => (cd.FormattedTime, cd.PointName));

            foreach (var data in normalizedData)
            {
                if (string.IsNullOrEmpty(data.PointName))
                    continue;

                var key = (data.FormattedTime, data.PointName);
                var matchingComparisonData = comparisonDataLookup[key].ToList();

                if (matchingComparisonData.Any())
                {
                    var comparisonData = matchingComparisonData.First();
                    var hasChanges = false;

                    // 检查并覆盖本期变化量（仅当对比数据不为空时）
                    if (IsValidValue(comparisonData.CurrentPeriodX))
                    {
                        data.CurrentPeriodX = comparisonData.CurrentPeriodX;
                        hasChanges = true;
                    }

                    if (IsValidValue(comparisonData.CurrentPeriodY))
                    {
                        data.CurrentPeriodY = comparisonData.CurrentPeriodY;
                        hasChanges = true;
                    }

                    if (IsValidValue(comparisonData.CurrentPeriodZ))
                    {
                        data.CurrentPeriodZ = comparisonData.CurrentPeriodZ;
                        hasChanges = true;
                    }

                    // 检查并覆盖累计变化量（仅当对比数据不为空时）
                    if (IsValidValue(comparisonData.CumulativeX))
                    {
                        data.CumulativeX = comparisonData.CumulativeX;
                        hasChanges = true;
                    }

                    if (IsValidValue(comparisonData.CumulativeY))
                    {
                        data.CumulativeY = comparisonData.CumulativeY;
                        hasChanges = true;
                    }

                    if (IsValidValue(comparisonData.CumulativeZ))
                    {
                        data.CumulativeZ = comparisonData.CumulativeZ;
                        hasChanges = true;
                    }

                    if (hasChanges)
                    {
                        mergeCount++;
                        //_logger.ConsoleInfo($"合并数据: {data.PointName} - {data.FormattedTime} - 行号:{data.RowNumber}");
                    }
                }
            }

            _logger.ShowComplete($"数据合并完成: 成功合并 {mergeCount} 条记录，对比数据总数 {totalComparisonData} 条");
        }

        /// <summary>
        /// 检查数值是否有效（不为空且不为默认值）
        /// </summary>
        /// <param name="value">要检查的数值</param>
        /// <returns>是否有效</returns>
        private bool IsValidValue(double value)
        {
            // 检查是否为NaN、无穷大或接近0的默认值
            return !double.IsNaN(value) && !double.IsInfinity(value) && Math.Abs(value) > 1e-10;
        }

        /// <summary>
        /// 步骤3: 数据分组和排序
        /// </summary>
        private List<MonitoringPoint> GroupAndSortData(List<PeriodData> normalizedData)
        {
            _logger.ShowStep("步骤3: 数据分组和排序...");

            var groupingService = new DataGroupingService(Log.ForContext<DataGroupingService>());
            var monitoringPoints = groupingService.GroupByPointName(normalizedData);
            groupingService.SortAllPointsByTime(monitoringPoints);

            _logger.ShowComplete($"分组完成: 监测点 {monitoringPoints.Count} 个");

            return monitoringPoints;
        }

        /// <summary>
        /// 步骤4: 数据验证
        /// </summary>
        private List<ValidationResult> ValidateData(List<MonitoringPoint> monitoringPoints, List<PeriodData> normalizedComparisonData)
        {
            _logger.ShowStep("步骤4: 数据验证...");

            // 创建配置服务并获取验证选项
            var configService = new ConfigurationService(Log.ForContext<ConfigurationService>());
            var validationOptions = configService.GetValidationOptions();

            var validationService = new DataValidationService(Log.ForContext<DataValidationService>(), validationOptions);

            // 添加超时保护和进度监控
            var validationResults = ValidateDataWithTimeout(validationService, monitoringPoints, normalizedComparisonData);

            // 将验证结果与监测点关联起来，为步骤5.5的精确验证做准备
            AssociateValidationResultsWithMonitoringPoints(monitoringPoints, validationResults);

            var validCount = validationResults.Count(v => v.Status == ValidationStatus.Valid);
            var invalidCount = validationResults.Count(v => v.Status == ValidationStatus.Invalid);
            var needsAdjustmentCount = validationResults.Count(v => v.Status == ValidationStatus.NeedsAdjustment);
            var canAdjustmentCount = validationResults.Count(v => v.CanAdjustment);

            _logger.ShowComplete($"验证完成: 通过 {validCount} 条, 失败 {invalidCount} 条, 需要修正 {needsAdjustmentCount} 条, 可以修正 {canAdjustmentCount} 条");

            return validationResults;
        }

        /// <summary>
        /// 带超时保护的数据验证（优化版本）
        /// </summary>
        /// <param name="validationService">验证服务</param>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <param name="normalizedComparisonData">对比数据</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateDataWithTimeout(DataValidationService validationService,
            List<MonitoringPoint> monitoringPoints, List<PeriodData> normalizedComparisonData)
        {
            var validationResults = new List<ValidationResult>();
            var totalPoints = monitoringPoints.Count;
            var startTime = DateTime.Now;

            // 获取验证选项
            var configService = new ConfigurationService(Log.ForContext<ConfigurationService>());
            var validationOptions = configService.GetValidationOptions();

            var maxProcessingTime = TimeSpan.FromMinutes(validationOptions.MaxProcessingTimeMinutes);
            var batchSize = validationOptions.BatchSize;
            var enableMemoryCleanup = validationOptions.EnableMemoryCleanup;
            var memoryCleanupFrequency = validationOptions.MemoryCleanupFrequency;

            _logger.ConsoleInfo($"开始验证 {totalPoints} 个监测点，最大处理时间: {maxProcessingTime.TotalMinutes} 分钟，批处理大小: {batchSize}");

            try
            {
                // 性能优化：预分配结果列表容量，减少动态扩容
                var estimatedResultsCount = totalPoints * 3; // 每个点平均3个方向
                validationResults = new List<ValidationResult>(estimatedResultsCount);

                // 分批处理监测点，避免长时间阻塞
                var processedPoints = 0;

                for (int i = 0; i < totalPoints; i += batchSize)
                {
                    // 检查是否超时
                    if (DateTime.Now - startTime > maxProcessingTime)
                    {
                        _logger.ShowWarning($"数据验证超时，已处理 {processedPoints}/{totalPoints} 个监测点");
                        break;
                    }

                    var batchEnd = Math.Min(i + batchSize, totalPoints);
                    var batch = monitoringPoints.Skip(i).Take(batchSize).ToList();

                    _logger.ConsoleInfo($"处理批次 {i / batchSize + 1}: 监测点 {i + 1} 到 {batchEnd}");

                    // 性能优化：使用快速验证模式
                    var batchResults = ValidateBatchWithOptimization(batch, normalizedComparisonData, validationService);
                    validationResults.AddRange(batchResults);

                    processedPoints += batch.Count;
                    var elapsed = DateTime.Now - startTime;
                    var estimatedTotal = elapsed.TotalSeconds * totalPoints / processedPoints;
                    var remaining = estimatedTotal - elapsed.TotalSeconds;

                    _logger.ConsoleInfo($"批次完成: {processedPoints}/{totalPoints} 个监测点，已用时: {elapsed:mm\\:ss}，预计剩余: {TimeSpan.FromSeconds(remaining):mm\\:ss}");

                    // 根据配置决定是否执行内存清理
                    if (enableMemoryCleanup && processedPoints % memoryCleanupFrequency == 0)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        _logger.ConsoleInfo("执行内存清理");
                    }
                }

                var totalElapsed = DateTime.Now - startTime;
                _logger.ConsoleInfo($"数据验证完成，总用时: {totalElapsed:mm\\:ss}，处理监测点: {processedPoints}/{totalPoints}");
            }
            catch (Exception ex)
            {
                _logger.ShowError($"数据验证过程中发生异常: {ex.Message}");
                _logger.FileError(ex, "数据验证过程中发生异常");

                // 如果验证失败，返回空的验证结果，避免后续步骤失败
                return new List<ValidationResult>();
            }

            return validationResults;
        }

        /// <summary>
        /// 批量验证优化版本
        /// </summary>
        /// <param name="batch">监测点批次</param>
        /// <param name="comparisonData">对比数据</param>
        /// <param name="validationService">验证服务</param>
        /// <returns>验证结果</returns>
        private List<ValidationResult> ValidateBatchWithOptimization(List<MonitoringPoint> batch,
            List<PeriodData> comparisonData, DataValidationService validationService)
        {
            var results = new List<ValidationResult>();

            // 性能优化：预分配容量
            var estimatedCapacity = batch.Count * 3;
            results = new List<ValidationResult>(estimatedCapacity);

            // 性能优化：创建对比数据的查找字典，避免重复LINQ查询
            var comparisonDataLookup = comparisonData
                .Where(cd => !string.IsNullOrEmpty(cd.PointName))
                .ToLookup(cd => (cd.PointName, cd.FormattedTime));

            // 并行处理批次内的监测点
            var lockObject = new object();
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount,
                CancellationToken = CancellationToken.None
            };

            Parallel.ForEach(batch, parallelOptions, point =>
            {
                try
                {
                    // 快速验证：只验证关键逻辑，跳过复杂的交叉验证
                    var pointResults = ValidateSinglePointOptimized(point, comparisonDataLookup);

                    lock (lockObject)
                    {
                        results.AddRange(pointResults);
                    }
                }
                catch (Exception ex)
                {
                    // 异常处理：创建错误结果
                    var errorResult = new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证过程中发生异常: {ex.Message}")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Critical
                    };

                    lock (lockObject)
                    {
                        results.Add(errorResult);
                    }
                }
            });

            return results;
        }

        /// <summary>
        /// 优化的单个监测点验证
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="comparisonDataLookup">对比数据查找表</param>
        /// <returns>验证结果</returns>
        private List<ValidationResult> ValidateSinglePointOptimized(MonitoringPoint point,
            ILookup<(string PointName, string FormattedTime), PeriodData> comparisonDataLookup)
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

                // 性能优化：快速验证累计变化量计算逻辑
                var cumulativeValidationResults = ValidateCumulativeCalculationOptimized(point);
                var limitValidationResults = ValidateValueLimits(new List<MonitoringPoint>() { point });
                if (limitValidationResults.Any())
                {
                    cumulativeValidationResults.AddRange(limitValidationResults);
                }

                // 如果直接校验正确，都不需要校验其它了，直接当这行记录是不需要修改的
                if (cumulativeValidationResults.Count == 0)
                {
                    results.Add(new ValidationResult(ValidationStatus.Valid, "数据验证", "累计变化量计算逻辑规则通过，不需要考虑其它处理了")
                    {
                        PointName = point.PointName,
                        Severity = ValidationSeverity.Info
                    });
                    return results;
                }

                // 性能优化：使用预构建的查找表进行交叉验证
                foreach (ValidationResult result in cumulativeValidationResults)
                {
                    var key = (point.PointName, result.FormattedTime);
                    var comparisonPointData = comparisonDataLookup[key].ToList();

                    // 取上一期数据，如果找不到，也可以修改
                    var previousData = GetPreviousPeriodDataOptimized(point, result.FormattedTime);

                    // 如果找不到，表示可以修改
                    if (comparisonPointData.Count == 0 || previousData == null)
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
                results.Add(new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证过程中发生异常: {ex.Message}")
                {
                    PointName = point.PointName,
                    Severity = ValidationSeverity.Error
                });
            }

            return results;
        }

        /// <summary>
        /// 优化的累计变化量验证
        /// </summary>
        /// <param name="point">监测点</param>
        /// <returns>验证结果</returns>
        private List<ValidationResult> ValidateCumulativeCalculationOptimized(MonitoringPoint point)
        {
            var results = new List<ValidationResult>();

            try
            {
                if (point.PeriodDataCount < 2) return results;

                // 性能优化：预分配容量
                var estimatedCapacity = (point.PeriodDataCount - 1) * 3; // 每期3个方向
                results = new List<ValidationResult>(estimatedCapacity);

                // 性能优化：使用数组而不是LINQ排序，减少内存分配
                var sortedData = point.PeriodDataList.ToArray();
                Array.Sort(sortedData, (a, b) =>
                {
                    var timeA = a.FileInfo?.FullDateTime ?? DateTime.MinValue;
                    var timeB = b.FileInfo?.FullDateTime ?? DateTime.MinValue;
                    return timeA.CompareTo(timeB);
                });

                // 性能优化：使用快速浮点数比较，避免复杂的decimal转换
                const double tolerance = 2.0; // 使用固定容差，避免配置查询

                // 逐期验证累计变化量计算逻辑
                for (int i = 1; i < sortedData.Length; i++)
                {
                    var previousPeriod = sortedData[i - 1];
                    var currentPeriod = sortedData[i];

                    // 验证X方向
                    var xValidation = ValidateCumulativeDirectionOptimized(
                        point.PointName, currentPeriod, previousPeriod.CumulativeX,
                        currentPeriod.CurrentPeriodX, currentPeriod.CumulativeX,
                        DataDirection.X, tolerance);
                    if (xValidation != null)
                    {
                        xValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(xValidation);
                    }

                    // 验证Y方向
                    var yValidation = ValidateCumulativeDirectionOptimized(
                        point.PointName, currentPeriod, previousPeriod.CumulativeY,
                        currentPeriod.CurrentPeriodY, currentPeriod.CumulativeY,
                        DataDirection.Y, tolerance);
                    if (yValidation != null)
                    {
                        yValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(yValidation);
                    }

                    // 验证Z方向
                    var zValidation = ValidateCumulativeDirectionOptimized(
                        point.PointName, currentPeriod, previousPeriod.CumulativeZ,
                        currentPeriod.CurrentPeriodZ, currentPeriod.CumulativeZ,
                        DataDirection.Z, tolerance);
                    if (zValidation != null)
                    {
                        zValidation.FormattedTime = currentPeriod.FormattedTime;
                        results.Add(zValidation);
                    }
                }
            }
            catch (Exception ex)
            {
                // 异常处理：记录错误但不中断验证
                results.Add(new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证累计变化量时发生异常: {ex.Message}")
                {
                    PointName = point.PointName,
                    Severity = ValidationSeverity.Error
                });
            }

            return results;
        }

        /// <summary>
        /// 优化的单个方向累计变化量验证
        /// </summary>
        /// <param name="pointName">点名</param>
        /// <param name="currentPeriod">当前期数据</param>
        /// <param name="previousCumulative">上一期累计值</param>
        /// <param name="currentPeriodValue">当前期变化量</param>
        /// <param name="currentCumulative">当前期累计值</param>
        /// <param name="direction">数据方向</param>
        /// <param name="tolerance">容差</param>
        /// <returns>验证结果</returns>
        private ValidationResult? ValidateCumulativeDirectionOptimized(
            string pointName, PeriodData currentPeriod, double previousCumulative,
            double currentPeriodValue, double currentCumulative, DataDirection direction, double tolerance)
        {
            try
            {
                // 性能优化：使用直接的双精度浮点数计算，避免decimal转换
                var expectedCumulative = previousCumulative + currentPeriodValue;
                var difference = Math.Abs(currentCumulative - expectedCumulative);

                // 检查是否在容差范围内
                if (difference <= tolerance)
                {
                    return null; // 验证通过
                }

                // 计算验证失败的程度
                var severity = difference > 5.0 ? ValidationSeverity.Critical :
                              difference > 3.0 ? ValidationSeverity.Error :
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

                // 性能优化：减少字符串格式化操作
                result.AddErrorDetail($"期望累计值: {expectedCumulative:F3}");
                result.AddErrorDetail($"实际累计值: {currentCumulative:F3}");
                result.AddErrorDetail($"差异: {difference:F3}");
                result.AddErrorDetail($"上一期累计值: {previousCumulative:F3}");
                result.AddErrorDetail($"本期变化量: {currentPeriodValue:F3}");

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
                // 异常处理：返回错误结果
                return new ValidationResult(ValidationStatus.Invalid, "验证异常", $"验证{direction}方向累计变化量时发生异常: {ex.Message}")
                {
                    PointName = pointName,
                    Severity = ValidationSeverity.Error
                };
            }
        }

        /// <summary>
        /// 优化的获取上一期数据方法
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="formattedTime">目标时间</param>
        /// <returns>上一期数据</returns>
        private PeriodData? GetPreviousPeriodDataOptimized(MonitoringPoint point, string formattedTime)
        {
            // 将 formattedTime 转换为 DateTime 进行比较
            if (!DateTime.TryParse(formattedTime, out DateTime targetTime))
            {
                return null;
            }

            // 性能优化：使用数组和循环，避免LINQ查询
            var dataArray = point.PeriodDataList.ToArray();
            PeriodData? previousPeriod = null;
            DateTime? latestTime = null;

            for (int i = 0; i < dataArray.Length; i++)
            {
                var pd = dataArray[i];
                if (pd.FileInfo?.FullDateTime != null && pd.FileInfo.FullDateTime < targetTime)
                {
                    if (latestTime == null || pd.FileInfo.FullDateTime > latestTime)
                    {
                        latestTime = pd.FileInfo.FullDateTime;
                        previousPeriod = pd;
                    }
                }
            }

            return previousPeriod;
        }

        /// <summary>
        /// 将验证结果与监测点关联起来
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <param name="validationResults">验证结果列表</param>
        private void AssociateValidationResultsWithMonitoringPoints(List<MonitoringPoint> monitoringPoints, List<ValidationResult> validationResults)
        {
            // 按点名分组验证结果
            var validationByPoint = validationResults
                .Where(v => !string.IsNullOrEmpty(v.PointName))
                .GroupBy(v => v.PointName)
                .ToDictionary(g => g.Key!, g => g.ToList());

            foreach (var point in monitoringPoints)
            {
                if (string.IsNullOrEmpty(point.PointName))
                    continue;

                if (validationByPoint.TryGetValue(point.PointName, out var pointValidations))
                {
                    // 检查该监测点是否有需要修正的验证结果
                    var hasInvalidAndAdjustable = pointValidations.Any(v =>
                        v.Status == ValidationStatus.Invalid && v.CanAdjustment);

                    // 检查该监测点是否有验证通过的记录
                    var hasValid = pointValidations.Any(v => v.Status == ValidationStatus.Valid);

                    // 如果监测点有验证通过的记录，且没有需要修正的验证结果，则标记为不需要验证
                    if (hasValid && !hasInvalidAndAdjustable)
                    {
                        point.ValidationStatus = ValidationStatus.Valid;
                    }
                    else if (hasInvalidAndAdjustable)
                    {
                        point.ValidationStatus = ValidationStatus.Invalid;
                    }
                    else
                    {
                        point.ValidationStatus = ValidationStatus.NotValidated;
                    }
                }
                else
                {
                    point.ValidationStatus = ValidationStatus.NotValidated;
                }
            }
        }

        /// <summary>
        /// 步骤5: 数据修正
        /// </summary>
        private CorrectionResult CorrectData(List<MonitoringPoint> monitoringPoints, List<ValidationResult> validationResults)
        {
            _logger.ShowStep("步骤5: 数据修正...");

            // 获取修正选项
            var configService = new ConfigurationService(Log.ForContext<ConfigurationService>());
            var correctionOptions = configService.GetCorrectionOptions();

            var correctionService = new DataCorrectionService(Log.ForContext<DataCorrectionService>(), correctionOptions);
            var correctionResult = correctionService.CorrectAllPoints(monitoringPoints, validationResults);

            _logger.ShowComplete($"修正完成: 修正 {correctionResult.AdjustmentRecords.Count} 条记录");

            return correctionResult;
        }

        /// <summary>
        /// 步骤5.5: 修正后重新验证数据
        /// </summary>
        private List<ValidationResult> ValidateCorrectedData(List<MonitoringPoint> monitoringPoints)
        {
            _logger.ShowStep("步骤5.5: 修正后重新验证数据...");

            var configService = new ConfigurationService(Log.ForContext<ConfigurationService>());
            var correctionOptions = configService.GetCorrectionOptions();
            var correctionService = new DataCorrectionService(Log.ForContext<DataCorrectionService>(), correctionOptions);

            //// 只验证那些在步骤4中被标记为无效且需要修正的监测点
            //var pointsToValidate = monitoringPoints.Where(p =>
            //    p.ValidationStatus == ValidationStatus.Invalid
            //).ToList();

            //if (!pointsToValidate.Any())
            //{
            //    _logger.ShowComplete("修正后验证: 没有需要验证的监测点");
            //    return new List<ValidationResult>();
            //}

            //var skippedPoints = monitoringPoints.Count - pointsToValidate.Count;
            //_logger.ConsoleInfo($"修正后验证: 只验证 {pointsToValidate.Count} 个需要修正的监测点，跳过 {skippedPoints} 个已通过或无需验证的监测点");

            var correctedValidationResults = correctionService.ValidateCorrectedMonitoringPoints(monitoringPoints);

            var correctedValidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Valid);
            var correctedInvalidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Invalid);

            _logger.ShowComplete($"修正后验证: 通过 {correctedValidCount} 条, 失败 {correctedInvalidCount} 条 ");

            return correctedValidationResults;
        }

        /// <summary>
        /// 验证数值超限检查
        /// 检查本期变化量不能超过2，累计变化量不能超过5
        /// </summary>
        /// <param name="monitoringPoints">监测点列表</param>
        /// <returns>验证结果列表</returns>
        private List<ValidationResult> ValidateValueLimits(List<MonitoringPoint> monitoringPoints)
        {
            var validationResults = new List<ValidationResult>();
            const double currentPeriodLimit = 2.0;  // 本期变化量限制
            const double cumulativeLimit = 5.0;     // 累计变化量限制
            foreach (var point in monitoringPoints)
            {
                foreach (var periodData in point.PeriodDataList)
                {
                    // 检查本期变化量
                    if (Math.Abs(periodData.CurrentPeriodX) > currentPeriodLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.X,
                            "本期变化量超限",
                            periodData.CurrentPeriodX,
                            currentPeriodLimit,
                            $"X方向本期变化量 {periodData.CurrentPeriodX:F3} 超过限制 {currentPeriodLimit}");
                        validationResults.Add(result);
                    }

                    if (Math.Abs(periodData.CurrentPeriodY) > currentPeriodLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.Y,
                            "本期变化量超限",
                            periodData.CurrentPeriodY,
                            currentPeriodLimit,
                            $"Y方向本期变化量 {periodData.CurrentPeriodY:F3} 超过限制 {currentPeriodLimit}");
                        validationResults.Add(result);
                    }

                    if (Math.Abs(periodData.CurrentPeriodZ) > currentPeriodLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.Z,
                            "本期变化量超限",
                            periodData.CurrentPeriodZ,
                            currentPeriodLimit,
                            $"Z方向本期变化量 {periodData.CurrentPeriodZ:F3} 超过限制 {currentPeriodLimit}");
                        validationResults.Add(result);
                    }

                    // 检查累计变化量
                    if (Math.Abs(periodData.CumulativeX) > cumulativeLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.X,
                            "累计变化量超限",
                            periodData.CumulativeX,
                            cumulativeLimit,
                            $"X方向累计变化量 {periodData.CumulativeX:F3} 超过限制 {cumulativeLimit}");
                        validationResults.Add(result);
                    }

                    if (Math.Abs(periodData.CumulativeY) > cumulativeLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.Y,
                            "累计变化量超限",
                            periodData.CumulativeY,
                            cumulativeLimit,
                            $"Y方向累计变化量 {periodData.CumulativeY:F3} 超过限制 {cumulativeLimit}");
                        validationResults.Add(result);
                    }

                    if (Math.Abs(periodData.CumulativeZ) > cumulativeLimit)
                    {
                        var result = CreateLimitValidationResult(
                            point, periodData, DataDirection.Z,
                            "累计变化量超限",
                            periodData.CumulativeZ,
                            cumulativeLimit,
                            $"Z方向累计变化量 {periodData.CumulativeZ:F3} 超过限制 {cumulativeLimit}");
                        validationResults.Add(result);
                    }
                }
            }

            if (validationResults.Any())
            {
                _logger.ShowWarning($"数值超限检查发现 {validationResults.Count} 个问题");
                // 显示所有的原因
                foreach (var result in validationResults)
                {
                    _logger.ConsoleInfo($"{result.FileName} {result.PointName} {result.DataDirection} {result.RowNumber} 原因: {result.Description}");
                }
            }

            return validationResults;
        }

        /// <summary>
        /// 创建数值超限验证结果
        /// </summary>
        /// <param name="point">监测点</param>
        /// <param name="periodData">期数据</param>
        /// <param name="direction">数据方向</param>
        /// <param name="validationType">验证类型</param>
        /// <param name="actualValue">实际值</param>
        /// <param name="limitValue">限制值</param>
        /// <param name="errorMessage">错误消息</param>
        /// <returns>验证结果</returns>
        private ValidationResult CreateLimitValidationResult(
            MonitoringPoint point,
            PeriodData periodData,
            DataDirection direction,
            string validationType,
            double actualValue,
            double limitValue,
            string errorMessage)
        {
            return new ValidationResult
            {
                Status = ValidationStatus.Invalid,
                CanAdjustment = false, // 超限数据不能修正
                ValidationType = validationType,
                Description = errorMessage,
                FileName = periodData.FileInfo?.OriginalFileName,
                PointName = point.PointName,
                RowNumber = periodData.RowNumber,
                DataDirection = direction,
                ErrorDetails = new List<string> { errorMessage },
                FailedValues = new Dictionary<string, object>
                {
                    { "实际值", actualValue },
                    { "限制值", limitValue },
                    { "超出量", Math.Abs(actualValue) - limitValue }
                },
                ExpectedValues = new Dictionary<string, object>
                {
                    { "期望范围", $"[-{limitValue:F3}, {limitValue:F3}]" }
                },
                ValidationRule = $"数值变化量不能超过限制：本期变化量≤2.000，累计变化量≤5.000",
                Severity = ValidationSeverity.Error
            };
        }

        /// <summary>
        /// 步骤6: 生成输出文件
        /// </summary>
        private (OutputResult outputResult, ReportResult reportResult) GenerateOutputFiles(
            List<MonitoringPoint> monitoringPoints, List<ValidationResult> validationResults,
            CorrectionResult correctionResult, string processedDirectory)
        {
            _logger.ShowStep("步骤6: 生成输出文件...");

            var outputDirectory = Path.Combine(processedDirectory, "修正后");

            // 获取输出选项
            var configService = new ConfigurationService(Log.ForContext<ConfigurationService>());
            var outputOptions = configService.GetOutputOptions();

            var outputService = new ExcelOutputService(Log.ForContext<ExcelOutputService>(), outputOptions);

            var outputResult = outputService.GenerateCorrectedExcelFiles(monitoringPoints, outputDirectory, processedDirectory);
            var reportResult = outputService.GenerateCorrectionReport(correctionResult, validationResults, outputDirectory);

            _logger.ShowComplete($"输出完成: 生成文件 {outputResult.FileResults.Count} 个, 报告 {reportResult.Status}");

            return (outputResult, reportResult);
        }

        /// <summary>
        /// 步骤7: 生成数据对比报告
        /// </summary>
        private void GenerateDataComparisonReport(List<ExcelReadResult> processedResults,
            List<ExcelReadResult> comparisonResults, CorrectionResult correctionResult, string processedDirectory)
        {
            _logger.ShowStep("步骤7: 生成数据对比报告...");

            try
            {
                // 统计processedResults中的数据行数
                var processedDataRowCount = processedResults.Sum(r => r.DataRows.Count);
                var processedFileCount = processedResults.Count;

                // 统计comparisonResults中的数据行数
                var comparisonDataRowCount = comparisonResults.Sum(r => r.DataRows.Count);
                var comparisonFileCount = comparisonResults.Count;

                // 计算数据增加情况
                var dataIncreaseCount = processedDataRowCount - comparisonDataRowCount;
                var dataIncreaseRatio = comparisonDataRowCount > 0 ? (double)dataIncreaseCount / comparisonDataRowCount * 100 : 0;

                // 统计修正的数据量
                var correctedDataCount = correctionResult.AdjustmentRecords.Count();
                var correctedDataRatio = processedDataRowCount > 0 ? (double)correctedDataCount / processedDataRowCount * 100 : 0;

                // 按修正类型统计
                var correctionTypeStats = correctionResult.AdjustmentRecords
                    .GroupBy(r => r.AdjustmentType)
                    .ToDictionary(g => g.Key.ToString(), g => g.Count());

                // 按数据方向统计
                var directionStats = correctionResult.AdjustmentRecords
                    .GroupBy(r => r.DataDirection)
                    .ToDictionary(g => g.Key.ToString(), g => g.Count());

                // 生成报告内容
                var reportContent = new List<string>
                {
                    new string('=', 80),
                    "数据对比分析报告",
                    new string('=', 80),
                    "",
                    $"报告生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                    "",
                    "1. 文件数量对比",
                    $"   待处理文件数量: {processedFileCount} 个",
                    $"   对比文件数量: {comparisonFileCount} 个",
                    $"   文件数量差异: {processedFileCount - comparisonFileCount} 个",
                    "",
                    "2. 数据行数对比",
                    $"   待处理数据行数: {processedDataRowCount} 行",
                    $"   对比数据行数: {comparisonDataRowCount} 行",
                    $"   数据行数增加: {dataIncreaseCount} 行",
                    $"   数据行数增加比例: {dataIncreaseRatio:F2}%",
                    "",
                    "3. 数据修正统计",
                    $"   修正记录总数: {correctedDataCount} 条",
                    $"   修正数据比例: {correctedDataRatio:F2}%",
                    "",
                    "4. 按修正类型统计",
                };

                foreach (var stat in correctionTypeStats)
                {
                    reportContent.Add($"   {stat.Key}: {stat.Value} 条");
                }

                reportContent.Add("");
                reportContent.Add("5. 按数据方向统计");
                foreach (var stat in directionStats)
                {
                    reportContent.Add($"   {stat.Key}: {stat.Value} 条");
                }

                reportContent.Add("");
                reportContent.Add("6. 详细修正记录");
                reportContent.Add(new string('-', 60));

                foreach (var record in correctionResult.AdjustmentRecords)
                {
                    reportContent.Add($"文件名: {record.FileName ?? "未知"}");
                    reportContent.Add($"点名: {record.PointName ?? "未知"}");
                    reportContent.Add($"行号: {record.RowNumber}");
                    reportContent.Add($"修正类型: {record.AdjustmentType}");
                    reportContent.Add($"数据方向: {record.DataDirection}");
                    reportContent.Add($"原始值: {record.OriginalValue:F6}");
                    reportContent.Add($"修正值: {record.AdjustedValue:F6}");
                    reportContent.Add($"修正幅度: {record.AdjustmentAmount:F6}");
                    reportContent.Add($"修正原因: {record.Reason ?? "未指定"}");
                    reportContent.Add(new string('-', 40));
                }

                // 保存报告到文件
                var outputDirectory = Path.Combine(processedDirectory, "修正后");
                var reportFilePath = Path.Combine(outputDirectory, "数据对比分析报告.txt");

                // 确保目录存在
                Directory.CreateDirectory(outputDirectory);

                File.WriteAllLines(reportFilePath, reportContent, System.Text.Encoding.UTF8);

                _logger.ShowComplete($"数据对比报告生成完成: {reportFilePath}");
                _logger.ConsoleInfo($"数据增加统计: {dataIncreaseCount} 行 ({dataIncreaseRatio:F2}%)");
                _logger.ConsoleInfo($"数据修正统计: {correctedDataCount} 条 ({correctedDataRatio:F2}%)");
            }
            catch (Exception ex)
            {
                _logger.ShowError($"生成数据对比报告时发生异常: {ex.Message}");
                _logger.FileError(ex, "生成数据对比报告时发生异常");
            }
        }

        /// <summary>
        /// 更新处理结果
        /// </summary>
        private void UpdateProcessingResult(ProcessingResult result, List<ExcelReadResult> processedResults,
            List<ExcelReadResult> comparisonResults, List<MonitoringPoint> monitoringPoints,
            List<ValidationResult> validationResults, CorrectionResult correctionResult,
            OutputResult outputResult, ReportResult reportResult)
        {
            result.Status = ProcessingStatus.Success;
            result.Message = "批量处理完成";
            result.ProcessedFiles = processedResults.Count;
            result.ComparisonFiles = comparisonResults.Count;
            result.MonitoringPoints = monitoringPoints.Count;
            result.ValidationResults = validationResults;
            result.CorrectionResult = correctionResult;
            result.OutputResult = outputResult;
            result.ReportResult = reportResult;
        }
    }
}
