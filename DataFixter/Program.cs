using DataFixter.Models;
using DataFixter.Services;
using DataFixter.Excel;
using Microsoft.Extensions.Logging;

namespace DataFixter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== DataFixter 数据修正工具 ===");
            Console.WriteLine("用于修复监测数据中累计变化量计算错误的工具");
            Console.WriteLine();

            // 添加测试修正算法的调用
            TestCorrectionAlgorithm();
            Console.WriteLine();

            try
            {
                // 检查命令行参数
                if (args.Length != 2)
                {
                    ShowUsage();
                    return;
                }

                var processedDirectory = args[0];
                var comparisonDirectory = args[1];

                // 验证目录是否存在
                if (!Directory.Exists(processedDirectory))
                {
                    Console.WriteLine($"错误: 待处理目录不存在: {processedDirectory}");
                    return;
                }

                if (!Directory.Exists(comparisonDirectory))
                {
                    Console.WriteLine($"错误: 对比目录不存在: {comparisonDirectory}");
                    return;
                }

                Console.WriteLine($"待处理目录: {processedDirectory}");
                Console.WriteLine($"对比目录: {comparisonDirectory}");
                Console.WriteLine();

                // 创建日志工厂
                var loggerFactory = LoggerFactory.Create(builder =>
                {
                    builder.AddConsole();
                });

                // 执行完整的批量处理流程
                var result = ExecuteBatchProcessing(loggerFactory, processedDirectory, comparisonDirectory);

                // 输出处理结果
                Console.WriteLine("=== 处理完成 ===");
                Console.WriteLine(result.GetSummary());
                Console.WriteLine();
                Console.WriteLine("详细报告已生成到输出目录");
                Console.WriteLine();
                Console.WriteLine("=== 验证修正结果 ===");
                Console.WriteLine("1. 检查 '修正后' 目录中的Excel文件");
                Console.WriteLine("2. 对比原始文件和修正后文件的数据");
                Console.WriteLine("3. 查看 '修正详细报告.txt' 了解具体修正内容");
                Console.WriteLine("4. 如果文件大小相同，这是正常的Excel格式特性");
                Console.WriteLine("5. 重点检查列3-8的数据是否已按修正逻辑更新");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"DataFixter 运行时发生致命错误: {ex.Message}");
                Console.WriteLine($"详细错误信息: {ex}");
            }
        }

        /// <summary>
        /// 显示使用说明
        /// </summary>
        static void ShowUsage()
        {
            Console.WriteLine("使用方法:");
            Console.WriteLine("DataFixter <待处理目录> <对比目录>");
            Console.WriteLine();
            Console.WriteLine("参数说明:");
            Console.WriteLine("  待处理目录: 包含需要修正的Excel文件的目录");
            Console.WriteLine("  对比目录: 包含对比数据的Excel文件的目录");
            Console.WriteLine();
            Console.WriteLine("示例:");
            Console.WriteLine("DataFixter \"E:\\data\\processed\" \"E:\\data\\comparison\"");
        }

        /// <summary>
        /// 执行完整的批量处理流程
        /// </summary>
        /// <param name="loggerFactory">日志工厂</param>
        /// <param name="processedDirectory">待处理目录</param>
        /// <param name="comparisonDirectory">对比目录</param>
        /// <returns>处理结果</returns>
        static ProcessingResult ExecuteBatchProcessing(ILoggerFactory loggerFactory, string processedDirectory, string comparisonDirectory)
        {
            var logger = loggerFactory.CreateLogger<Program>();
            var result = new ProcessingResult();

            try
            {
                logger.LogInformation("开始批量处理...");

                // 步骤1: 读取Excel文件
                Console.WriteLine("步骤1: 读取Excel文件...");
                var excelReader = new ExcelBatchReader(processedDirectory, loggerFactory.CreateLogger<ExcelBatchReader>());
                var processedResults = excelReader.ReadAllFiles();
                
                var comparisonReader = new ExcelBatchReader(comparisonDirectory, loggerFactory.CreateLogger<ExcelBatchReader>());
                var comparisonResults = comparisonReader.ReadAllFiles();

                Console.WriteLine($"  读取完成: 待处理文件 {processedResults.Count} 个, 对比文件 {comparisonResults.Count} 个");

                // 步骤2: 数据标准化
                Console.WriteLine("步骤2: 数据标准化...");
                var normalizer = new DataNormalizer(loggerFactory.CreateLogger<DataNormalizer>());
                var normalizedData = normalizer.NormalizeData(processedResults);
                var normalizedComparisonData = normalizer.NormalizeData(comparisonResults);

                Console.WriteLine($"  标准化完成: 待处理数据 {normalizedData.Count} 条, 对比数据 {normalizedComparisonData.Count} 条");

                // 步骤3: 数据分组和排序
                Console.WriteLine("步骤3: 数据分组和排序...");
                var groupingService = new DataGroupingService(loggerFactory.CreateLogger<DataGroupingService>());
                var monitoringPoints = groupingService.GroupByPointName(normalizedData);
                groupingService.SortAllPointsByTime(monitoringPoints);

                Console.WriteLine($"  分组完成: 监测点 {monitoringPoints.Count} 个");

                // 步骤4: 数据验证
                Console.WriteLine("步骤4: 数据验证...");
                
                // 创建配置服务并获取验证选项
                var configService = new ConfigurationService(loggerFactory.CreateLogger<ConfigurationService>());
                var validationOptions = configService.GetValidationOptions();
                
                var validationService = new DataValidationService(loggerFactory.CreateLogger<DataValidationService>(), validationOptions);
                var validationResults = validationService.ValidateAllPoints(monitoringPoints, normalizedComparisonData);

                var validCount = validationResults.Count(v => v.Status == ValidationStatus.Valid);
                var invalidCount = validationResults.Count(v => v.Status == ValidationStatus.Invalid);
                var needsAdjustmentCount = validationResults.Count(v => v.Status == ValidationStatus.NeedsAdjustment);

                Console.WriteLine($"  验证完成: 通过 {validCount} 条, 失败 {invalidCount} 条, 需要修正 {needsAdjustmentCount} 条");

                // 步骤5: 数据修正
                Console.WriteLine("步骤5: 数据修正...");
                
                // 获取修正选项
                var correctionOptions = configService.GetCorrectionOptions();
                
                var correctionService = new DataCorrectionService(loggerFactory.CreateLogger<DataCorrectionService>(), correctionOptions);
                var correctionResult = correctionService.CorrectAllPoints(monitoringPoints, validationResults);

                Console.WriteLine($"  修正完成: 修正 {correctionResult.AdjustmentRecords.Count} 条记录");

                // 步骤5.5: 修正后重新验证数据
                Console.WriteLine("步骤5.5: 修正后重新验证数据...");
                var correctedValidationResults = correctionService.ValidateCorrectedMonitoringPoints(monitoringPoints);
                
                var correctedValidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Valid);
                var correctedInvalidCount = correctedValidationResults.Count(v => v.Status == ValidationStatus.Invalid);
                
                Console.WriteLine($"  修正后验证: 通过 {correctedValidCount} 条, 失败 {correctedInvalidCount} 条");

                // 调试信息：检查修正前后的数据
                Console.WriteLine("=== 调试信息：数据修正检查 ===");
                var samplePoint = monitoringPoints.FirstOrDefault();
                if (samplePoint != null)
                {
                    Console.WriteLine($"示例监测点: {samplePoint.PointName}");
                    var sampleData = samplePoint.PeriodDataList.FirstOrDefault();
                    if (sampleData != null)
                    {
                        Console.WriteLine($"  修正前 - X本期: {sampleData.CurrentPeriodX}, X累计: {sampleData.CumulativeX}");
                        Console.WriteLine($"  修正前 - Y本期: {sampleData.CurrentPeriodY}, Y累计: {sampleData.CumulativeY}");
                        Console.WriteLine($"  修正前 - Z本期: {sampleData.CurrentPeriodZ}, Z累计: {sampleData.CumulativeZ}");
                    }
                }

                // 步骤6: 生成输出文件
                Console.WriteLine("步骤6: 生成输出文件...");
                var outputDirectory = Path.Combine(processedDirectory, "修正后");
                
                // 获取输出选项
                var outputOptions = configService.GetOutputOptions();
                
                var outputService = new ExcelOutputService(loggerFactory.CreateLogger<ExcelOutputService>(), outputOptions);

                var outputResult = outputService.GenerateCorrectedExcelFiles(monitoringPoints, outputDirectory, processedDirectory);
                var reportResult = outputService.GenerateCorrectionReport(correctionResult, validationResults, outputDirectory);

                Console.WriteLine($"  输出完成: 生成文件 {outputResult.FileResults.Count} 个, 报告 {reportResult.Status}");

                // 更新结果
                result.Status = ProcessingStatus.Success;
                result.Message = "批量处理完成";
                result.ProcessedFiles = processedResults.Count;
                result.ComparisonFiles = comparisonResults.Count;
                result.MonitoringPoints = monitoringPoints.Count;
                result.ValidationResults = validationResults;
                result.CorrectionResult = correctionResult;
                result.OutputResult = outputResult;
                result.ReportResult = reportResult;

                logger.LogInformation("批量处理完成");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "批量处理过程中发生异常");
                result.Status = ProcessingStatus.Error;
                result.Message = $"处理过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 测试修正算法
        /// </summary>
        private static void TestCorrectionAlgorithm()
        {
            Console.WriteLine("=== 测试修正算法 ===");
            
            // 创建测试数据 - 包含需要双重修正的情况
            var testPoint = new MonitoringPoint
            {
                PointName = "测试点1",
                PeriodDataList = new List<PeriodData>
                {
                    // 第0期：基准期
                    new PeriodData
                    {
                        RowNumber = 1,
                        CurrentPeriodX = 0.0,
                        CurrentPeriodY = 0.0,
                        CurrentPeriodZ = 0.0,
                        CumulativeX = 0.0,
                        CumulativeY = 0.0,
                        CumulativeZ = 0.0,
                        FileInfo = new DataFixter.Models.FileInfo("test_2025.1.1-00测试.xls", 1024, DateTime.Now.AddDays(-2))
                    },
                    // 第1期：本期变化量和累计值都有问题，需要双重修正
                    new PeriodData
                    {
                        RowNumber = 2,
                        CurrentPeriodX = 2.5, // 超出1.0限制
                        CurrentPeriodY = 1.8, // 超出1.0限制
                        CurrentPeriodZ = 1.2, // 超出1.0限制
                        CumulativeX = 1.2,   // 错误：应该是 0.0 + 2.5 = 2.5
                        CumulativeY = 0.8,   // 错误：应该是 0.0 + 1.8 = 1.8
                        CumulativeZ = 0.5,   // 错误：应该是 0.0 + 1.2 = 1.2
                        FileInfo = new DataFixter.Models.FileInfo("test_2025.1.2-00测试.xls", 1024, DateTime.Now.AddDays(-1))
                    },
                    // 第2期：本期变化量和累计值都有问题，需要双重修正
                    new PeriodData
                    {
                        RowNumber = 3,
                        CurrentPeriodX = 1.8, // 超出1.0限制
                        CurrentPeriodY = 1.5, // 超出1.0限制
                        CurrentPeriodZ = 1.1, // 超出1.0限制
                        CumulativeX = 2.1,   // 错误：应该是 2.5 + 1.8 = 4.3
                        CumulativeY = 1.4,   // 错误：应该是 1.8 + 1.5 = 3.3
                        CumulativeZ = 0.9,   // 错误：应该是 1.2 + 1.1 = 2.3
                        FileInfo = new DataFixter.Models.FileInfo("test_2025.1.3-00测试.xls", 1024, DateTime.Now)
                    }
                }
            };
            
            var testPoints = new List<MonitoringPoint> { testPoint };

            // 创建配置选项 - 使用配置文件中的值，而不是硬编码的严格值
            var options = new Services.CorrectionOptions
            {
                CumulativeTolerance = 0.1,  // 使用配置文件中的值
                MaxCurrentPeriodValue = 5.0, // 使用配置文件中的值
                MaxCumulativeValue = 10.0    // 使用配置文件中的值
            };

            // 创建数据修正服务
            var logger = LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<Services.DataCorrectionService>();
            var correctionService = new Services.DataCorrectionService(logger, options);

            Console.WriteLine("修正前数据:");
            for (int i = 0; i < testPoint.PeriodDataCount; i++)
            {
                var data = testPoint.PeriodDataList[i];
                Console.WriteLine($"第{i}期: X本期={data.CurrentPeriodX:F3}, X累计={data.CumulativeX:F3}");
                Console.WriteLine($"        Y本期={data.CurrentPeriodY:F3}, Y累计={data.CumulativeY:F3}");
                Console.WriteLine($"        Z本期={data.CurrentPeriodZ:F3}, Z累计={data.CumulativeZ:F3}");
            }

            // 创建模拟的验证结果（标记需要修正的数据）
            var validationResults = new List<ValidationResult>
            {
                new ValidationResult
                {
                    PointName = "测试点1",
                    Status = ValidationStatus.Invalid,
                    ValidationType = "累计值关系验证",
                    Description = "第1期本期变化量与累计值不一致"
                },
                new ValidationResult
                {
                    PointName = "测试点1",
                    Status = ValidationStatus.Invalid,
                    ValidationType = "累计值关系验证",
                    Description = "第2期累计值计算错误"
                }
            };

            // 执行修正 - 使用正确的方法
            var result = correctionService.CorrectAllPoints(testPoints, validationResults);
            Console.WriteLine($"\n修正结果: {result.Status}, 修正数量: {result.AdjustmentRecords.Count}");

            Console.WriteLine("\n修正后数据:");
            for (int i = 0; i < testPoint.PeriodDataCount; i++)
            {
                var data = testPoint.PeriodDataList[i];
                Console.WriteLine($"第{i}期: X本期={data.CurrentPeriodX:F3}, X累计={data.CumulativeX:F3}");
                Console.WriteLine($"        Y本期={data.CurrentPeriodY:F3}, Y累计={data.CumulativeY:F3}");
                Console.WriteLine($"        Z本期={data.CurrentPeriodZ:F3}, Z累计={data.CumulativeZ:F3}");
            }

            // 验证修正后的数据
            var validationResult = correctionService.ValidateCorrectedData(testPoint.PeriodDataList);
            Console.WriteLine($"\n验证结果: {validationResult.Status}");
            if (validationResult.Status == ValidationStatus.Invalid)
            {
                Console.WriteLine("验证失败，错误详情:");
                foreach (var error in validationResult.ErrorDetails)
                {
                    Console.WriteLine($"  - {error}");
                }
            }
            else
            {
                Console.WriteLine("验证通过！所有数据都满足累计值关系。");
            }
            
            // 输出修正详情
            Console.WriteLine("\n=== 修正详情 ===");
            foreach (var pointResult in result.PointResults)
            {
                Console.WriteLine($"监测点: {pointResult.PointName} - {pointResult.Status} - {pointResult.Message}");
                if (pointResult.Corrections.Any())
                {
                    foreach (var correction in pointResult.Corrections)
                    {
                        Console.WriteLine($"  修正: {correction.Direction}方向 {correction.CorrectionType} - {correction.OriginalValue:F6} -> {correction.CorrectedValue:F6}");
                        Console.WriteLine($"  原因: {correction.Reason}");
                    }
                }
            }
        }

        // 转换方法已移除，使用DataNormalizer.NormalizeData替代

        /// <summary>
        /// 测试数据验证服务
        /// </summary>
        /// <param name="loggerFactory">日志工厂</param>
        static void TestDataValidationService(ILoggerFactory loggerFactory)
        {
            var logger = loggerFactory.CreateLogger<Program>();
            logger.LogInformation("开始测试数据验证服务...");

            try
            {
                // 创建测试数据
                var monitoringPoints = CreateTestMonitoringPoints();
                var comparisonData = CreateTestComparisonData();

                // 创建数据验证服务
                var validationService = new DataValidationService(loggerFactory.CreateLogger<DataValidationService>());

                // 执行验证
                var validationResults = validationService.ValidateAllPoints(monitoringPoints, comparisonData);

                // 获取统计信息
                var statistics = validationService.GetValidationStatistics(validationResults);

                // 输出结果
                Console.WriteLine("验证完成，结果统计:");
                Console.WriteLine(statistics.GetSummary());
                Console.WriteLine(statistics.GetDetailedInfo());

                // 输出验证结果详情
                foreach (var result in validationResults.Take(5)) // 只显示前5个结果
                {
                    Console.WriteLine($"验证结果: {result.GetSummary()}");
                }

                logger.LogInformation("数据验证服务测试完成");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "测试数据验证服务时发生异常");
            }
        }

        /// <summary>
        /// 创建测试用的监测点数据
        /// </summary>
        /// <returns>监测点列表</returns>
        private static List<MonitoringPoint> CreateTestMonitoringPoints()
        {
            var points = new List<MonitoringPoint>();

            // 创建第一个监测点（有错误的累计变化量计算）
            var point1 = new MonitoringPoint("测试点1", 100.50);
            
            var period1 = new PeriodData
            {
                FileInfo = new Models.FileInfo("test1.xls", 1024, DateTime.Now),
                RowNumber = 5,
                PointName = "测试点1",
                Mileage = 100.50,
                CurrentPeriodX = 0.001,
                CurrentPeriodY = 0.002,
                CurrentPeriodZ = 0.003,
                CumulativeX = 0.001,
                CumulativeY = 0.002,
                CumulativeZ = 0.003,
                DailyX = 0.0001,
                DailyY = 0.0002,
                DailyZ = 0.0003
            };

            var period2 = new PeriodData
            {
                FileInfo = new Models.FileInfo("test2.xls", 1024, DateTime.Now.AddDays(1)),
                RowNumber = 5,
                PointName = "测试点1",
                Mileage = 100.50,
                CurrentPeriodX = 0.002,
                CurrentPeriodY = 0.003,
                CurrentPeriodZ = 0.004,
                CumulativeX = 0.005, // 错误：应该是 0.001 + 0.002 = 0.003
                CumulativeY = 0.005, // 正确：0.002 + 0.003 = 0.005
                CumulativeZ = 0.010, // 错误：应该是 0.003 + 0.004 = 0.007
                DailyX = 0.0002,
                DailyY = 0.0003,
                DailyZ = 0.0004
            };

            point1.AddPeriodData(period1);
            point1.AddPeriodData(period2);
            points.Add(point1);

            // 创建第二个监测点（完全正确）
            var point2 = new MonitoringPoint("测试点2", 200.75);
            
            var period3 = new PeriodData
            {
                FileInfo = new Models.FileInfo("test1.xls", 1024, DateTime.Now),
                RowNumber = 6,
                PointName = "测试点2",
                Mileage = 200.75,
                CurrentPeriodX = 0.004,
                CurrentPeriodY = 0.005,
                CurrentPeriodZ = 0.006,
                CumulativeX = 0.004,
                CumulativeY = 0.005,
                CumulativeZ = 0.006,
                DailyX = 0.0004,
                DailyY = 0.0005,
                DailyZ = 0.0006
            };

            point2.AddPeriodData(period3);
            points.Add(point2);

            return points;
        }

        /// <summary>
        /// 测试数据修正服务
        /// </summary>
        /// <param name="loggerFactory">日志工厂</param>
        static void TestDataCorrectionService(ILoggerFactory loggerFactory)
        {
            var logger = loggerFactory.CreateLogger<Program>();
            logger.LogInformation("开始测试数据修正服务...");

            try
            {
                // 创建测试数据
                var monitoringPoints = CreateTestMonitoringPoints();
                var comparisonData = CreateTestComparisonData();

                // 创建数据验证服务
                var validationService = new DataValidationService(loggerFactory.CreateLogger<DataValidationService>());

                // 执行验证
                var validationResults = validationService.ValidateAllPoints(monitoringPoints, comparisonData);

                // 创建数据修正服务
                var correctionService = new DataCorrectionService(loggerFactory.CreateLogger<DataCorrectionService>());

                // 执行修正
                var correctionResult = correctionService.CorrectAllPoints(monitoringPoints, validationResults);

                // 获取统计信息
                var statistics = correctionService.GetCorrectionStatistics();

                // 输出结果
                Console.WriteLine("\n=== 数据修正结果 ===");
                Console.WriteLine(correctionResult.GetSummary());

                // 输出修正详情
                foreach (var pointResult in correctionResult.PointResults.Take(5)) // 只显示前5个结果
                {
                    Console.WriteLine($"修正结果: {pointResult.PointName} - {pointResult.Status} - {pointResult.Message}");
                    if (pointResult.Corrections.Any())
                    {
                        foreach (var correction in pointResult.Corrections.Take(3)) // 只显示前3个修正
                        {
                            Console.WriteLine($"  修正: {correction.Direction}方向 {correction.CorrectionType} - {correction.OriginalValue:F6} -> {correction.CorrectedValue:F6}");
                        }
                    }
                }

                // 输出修正统计
                Console.WriteLine($"\n修正统计: 总计{statistics.TotalAdjustments}次修正, 涉及{statistics.TotalPoints}个监测点, {statistics.TotalFiles}个文件");

                logger.LogInformation("数据修正服务测试完成");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "测试数据修正服务时发生异常");
            }
        }

        /// <summary>
        /// 创建测试用的对比数据
        /// </summary>
        /// <returns>对比数据列表</returns>
        private static List<PeriodData> CreateTestComparisonData()
        {
            var data = new List<PeriodData>();

            // 对比数据1
            data.Add(new PeriodData
            {
                FileInfo = new Models.FileInfo("comparison1.xls", 1024, DateTime.Now),
                RowNumber = 5,
                PointName = "测试点1",
                Mileage = 100.50,
                CurrentPeriodX = 0.001,
                CurrentPeriodY = 0.002,
                CurrentPeriodZ = 0.003,
                CumulativeX = 0.001,
                CumulativeY = 0.002,
                CumulativeZ = 0.003,
                DailyX = 0.0001,
                DailyY = 0.0002,
                DailyZ = 0.0003
            });

            // 对比数据2
            data.Add(new PeriodData
            {
                FileInfo = new Models.FileInfo("comparison2.xls", 1024, DateTime.Now),
                RowNumber = 6,
                PointName = "测试点2",
                Mileage = 200.75,
                CurrentPeriodX = 0.004,
                CurrentPeriodY = 0.005,
                CurrentPeriodZ = 0.006,
                CumulativeX = 0.004,
                CumulativeY = 0.005,
                CumulativeZ = 0.006,
                DailyX = 0.0004,
                DailyY = 0.0005,
                DailyZ = 0.0006
            });

            // 对比数据3（空数据，用于测试对比数据为空的情况）
            data.Add(new PeriodData
            {
                FileInfo = new Models.FileInfo("comparison3.xls", 1024, DateTime.Now),
                RowNumber = 7,
                PointName = "测试点3",
                Mileage = 300.00,
                CurrentPeriodX = 0.0,
                CurrentPeriodY = 0.0,
                CurrentPeriodZ = 0.0,
                CumulativeX = 0.0,
                CumulativeY = 0.0,
                CumulativeZ = 0.0,
                DailyX = 0.0,
                DailyY = 0.0,
                DailyZ = 0.0
            });

            return data;
        }
    }
}
