using DataFixter.Models;
using DataFixter.Services;
using Serilog;

namespace DataFixter
{
    /// <summary>
    /// 测试程序，用于测试各种功能模块
    /// </summary>
    public class TestProgram
    {
        /// <summary>
        /// 测试修正算法
        /// </summary>
        public static void TestCorrectionAlgorithm()
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
                        FileInfo = new ExcelFileInfo("test_2025.1.1-00测试.xls", 1024, DateTime.Now.AddDays(-2))
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
                        FileInfo = new ExcelFileInfo("test_2025.1.2-00测试.xls", 1024, DateTime.Now.AddDays(-1))
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
                        FileInfo = new ExcelFileInfo("test_2025.1.3-00测试.xls", 1024, DateTime.Now)
                    }
                }
            };
            
            var testPoints = new List<MonitoringPoint> { testPoint };

            // 创建配置选项 - 使用更严格的值来测试修正逻辑
            var options = new DataFixter.Services.CorrectionOptions
            {
                CumulativeTolerance = 0.1,  // 使用配置文件中的值
                MaxCurrentPeriodValue = 1.0, // 使用更严格的值，使测试数据能够触发修正
                MaxCumulativeValue = 10.0    // 使用配置文件中的值
            };

            // 创建数据修正服务
            var logger = Log.ForContext<DataCorrectionService>();
            var correctionService = new DataCorrectionService(logger, options);

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

        /// <summary>
        /// 测试数据验证服务
        /// </summary>
        public static void TestDataValidationService()
        {
            var logger = Log.ForContext<TestProgram>();
            logger.Information("开始测试数据验证服务...");

            try
            {
                // 创建测试数据
                var monitoringPoints = CreateTestMonitoringPoints();
                var comparisonData = CreateTestComparisonData();

                // 创建数据验证服务
                var validationService = new DataValidationService(Log.ForContext<DataValidationService>());

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

                logger.Information("数据验证服务测试完成");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "测试数据验证服务时发生异常");
            }
        }

        /// <summary>
        /// 测试数据修正服务
        /// </summary>
        public static void TestDataCorrectionService()
        {
            var logger = Log.ForContext<TestProgram>();
            logger.Information("开始测试数据修正服务...");

            try
            {
                // 创建测试数据
                var monitoringPoints = CreateTestMonitoringPoints();
                var comparisonData = CreateTestComparisonData();

                // 创建数据验证服务
                var validationService = new DataValidationService(Log.ForContext<DataValidationService>());

                // 执行验证
                var validationResults = validationService.ValidateAllPoints(monitoringPoints, comparisonData);

                // 创建数据修正服务
                var correctionService = new DataCorrectionService(Log.ForContext<DataCorrectionService>());

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

                logger.Information("数据修正服务测试完成");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "测试数据修正服务时发生异常");
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
                FileInfo = new ExcelFileInfo("test1.xls", 1024, DateTime.Now),
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
                FileInfo = new ExcelFileInfo("test2.xls", 1024, DateTime.Now.AddDays(1)),
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
                FileInfo = new ExcelFileInfo("test1.xls", 1024, DateTime.Now),
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
        /// 创建测试用的对比数据
        /// </summary>
        /// <returns>对比数据列表</returns>
        private static List<PeriodData> CreateTestComparisonData()
        {
            var data = new List<PeriodData>();

            // 对比数据1
            data.Add(new PeriodData
            {
                FileInfo = new ExcelFileInfo("comparison1.xls", 1024, DateTime.Now),
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
                FileInfo = new ExcelFileInfo("comparison2.xls", 1024, DateTime.Now),
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
                FileInfo = new ExcelFileInfo("comparison3.xls", 1024, DateTime.Now),
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

        /// <summary>
        /// 运行所有测试
        /// </summary>
        public static void RunAllTests()
        {
            Console.WriteLine("=== 开始运行所有测试 ===");
            Console.WriteLine();

            try
            {
                // 测试修正算法
                TestCorrectionAlgorithm();
                Console.WriteLine();

                // 测试数据验证服务
                TestDataValidationService();
                Console.WriteLine();

                // 测试数据修正服务
                TestDataCorrectionService();
                Console.WriteLine();

                Console.WriteLine("=== 所有测试完成 ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"测试过程中发生异常: {ex.Message}");
                Console.WriteLine($"详细错误信息: {ex}");
            }
        }
    }
} 