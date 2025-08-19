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

                // 步骤6: 生成输出文件
                var (outputResult, reportResult) = GenerateOutputFiles(monitoringPoints, validationResults, correctionResult, processedDirectory);

                // 步骤7: 生成数据对比报告
                GenerateDataComparisonReport(processedResults, comparisonResults, correctionResult, processedDirectory);

                // 更新结果
                UpdateProcessingResult(result, processedResults, comparisonResults, monitoringPoints, 
                    validationResults, correctionResult, outputResult, reportResult);

                _logger.LogOperationComplete("批量处理", "成功完成", 
                    $"处理文件：{processedResults.Count}个，监测点：{monitoringPoints.Count}个，修正记录：{correctionResult.AdjustmentRecords.Count}条");
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
                        _logger.ConsoleInfo($"合并数据: {data.PointName} - {data.FormattedTime} - 行号:{data.RowNumber}");
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
            var validationResults = validationService.ValidateAllPoints(monitoringPoints, normalizedComparisonData);

            var validCount = validationResults.Count(v => v.Status == ValidationStatus.Valid);
            var invalidCount = validationResults.Count(v => v.Status == ValidationStatus.Invalid);
            var needsAdjustmentCount = validationResults.Count(v => v.Status == ValidationStatus.NeedsAdjustment);
            var canAdjustmentCount = validationResults.Count(v => v.CanAdjustment);

            _logger.ShowComplete($"验证完成: 通过 {validCount} 条, 失败 {invalidCount} 条, 需要修正 {needsAdjustmentCount} 条, 可以修正 {canAdjustmentCount} 条");

            return validationResults;
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
            
            var correctedValidationResults = correctionService.ValidateCorrectedMonitoringPoints(monitoringPoints);
            
            // 添加数值超限检查
            var limitValidationResults = ValidateValueLimits(monitoringPoints);
            correctedValidationResults.AddRange(limitValidationResults);
            
            var correctedValidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Valid);
            var correctedInvalidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Invalid);
            
            _logger.ShowComplete($"修正后验证: 通过 {correctedValidCount} 条, 失败 {correctedInvalidCount} 条");

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
                _logger.ShowWarning($"数值超限检查发现 {validationResults.Count} 个问题，修正失败，不能保存");
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
                var correctedDataCount = correctionResult.AdjustmentRecords.Count;
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
