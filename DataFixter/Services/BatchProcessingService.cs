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
            
            var correctedValidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Valid);
            var correctedInvalidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Invalid);
            
            _logger.ShowComplete($"修正后验证: 通过 {correctedValidCount} 条, 失败 {correctedInvalidCount} 条");

            return correctedValidationResults;
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
