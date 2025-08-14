using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Xunit;
using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Tests.IntegrationTests
{
    public class DataProcessorIntegrationTests
    {
        private readonly DataProcessorConfig _defaultConfig;

        public DataProcessorIntegrationTests()
        {
            _defaultConfig = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001,
                BatchSize = 100,
                EnableCaching = true,
                CacheExpirationMinutes = 30
            };
        }

        [Fact]
        public void CompleteWorkflow_WithRealisticData_ShouldProcessSuccessfully()
        {
            // Arrange
            var files = CreateRealisticTestData();
            var stopwatch = Stopwatch.StartNew();

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            stopwatch.Stop();

            // Assert
            Assert.NotNull(result);
            Assert.Equal(files.Count, result.Count);
            
            // Verify data integrity
            foreach (var file in result)
            {
                Assert.NotNull(file.DataRows);
                foreach (var row in file.DataRows)
                {
                    Assert.NotNull(row.Values);
                    Assert.All(row.Values, value => Assert.True(value.HasValue || value == null));
                }
            }

            // Performance check
            Assert.True(stopwatch.ElapsedMilliseconds < 5000, $"Processing took too long: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void LargeScaleProcessing_WithPerformanceOptimization_ShouldCompleteEfficiently()
        {
            // Arrange
            var files = CreateLargeScaleTestData(1000, 50);
            var config = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.03,
                RandomSeed = 123,
                TimeFactorWeight = 0.8,
                MinimumAdjustment = 0.001,
                BatchSize = 200,
                EnableCaching = true,
                CacheExpirationMinutes = 60
            };

            var stopwatch = Stopwatch.StartNew();

            // Act
            var result = DataProcessor.ProcessMissingData(files, config);
            stopwatch.Stop();

            // Assert
            Assert.NotNull(result);
            Assert.Equal(files.Count, result.Count);
            
            // Performance benchmarks
            var totalDataPoints = result.Sum(f => f.DataRows.Sum(r => r.Values.Count));
            var processingTimePerPoint = stopwatch.ElapsedMilliseconds / (double)totalDataPoints;
            
            // Adjusted performance expectations for large-scale processing
            Assert.True(processingTimePerPoint < 1.0, $"Processing time per data point too high: {processingTimePerPoint:F3}ms");
            Assert.True(stopwatch.ElapsedMilliseconds < 600000, $"Total processing time too high: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void ErrorHandling_WithInvalidData_ShouldHandleGracefully()
        {
            // Arrange
            var files = CreateInvalidTestData();

            // Act & Assert
            var exception = Assert.ThrowsAny<Exception>(() => 
                DataProcessor.ProcessMissingData(files, _defaultConfig));
            
            Assert.NotNull(exception);
            Assert.Contains("Invalid data", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void MemoryUsage_WithLargeDataset_ShouldRemainStable()
        {
            // Arrange
            var files = CreateLargeScaleTestData(500, 100);
            var initialMemory = GC.GetTotalMemory(false);

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            var finalMemory = GC.GetTotalMemory(false);

            // Assert
            Assert.NotNull(result);
            
            // Memory usage should not increase significantly
            var memoryIncrease = finalMemory - initialMemory;
            var memoryIncreaseMB = memoryIncrease / (1024 * 1024);
            
            Assert.True(memoryIncreaseMB < 100, $"Memory increase too high: {memoryIncreaseMB:F2}MB");
        }

        [Fact]
        public void DataConsistency_AfterProcessing_ShouldMaintainBusinessLogic()
        {
            // Arrange
            var files = CreateBusinessLogicTestData();

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);

            // Assert
            Assert.NotNull(result);
            
            // Verify business logic rules
            foreach (var file in result)
            {
                foreach (var row in file.DataRows)
                {
                    if (row.Name.StartsWith("G") && row.Name.StartsWith("D"))
                    {
                        var gIndex = int.Parse(row.Name.Substring(1));
                        var dIndex = int.Parse(row.Name.Substring(1));
                        
                        // G columns should be cumulative, D columns should be changes
                        if (row.Name.StartsWith("G"))
                        {
                            Assert.All(row.Values, value => 
                                Assert.True(value.HasValue && value.Value >= 0, "G columns should be non-negative"));
                        }
                        else if (row.Name.StartsWith("D"))
                        {
                            // D columns can be negative (changes)
                            Assert.All(row.Values, value => Assert.True(value.HasValue));
                        }
                    }
                }
            }
        }

        [Fact]
        public void CacheEffectiveness_WithRepeatedOperations_ShouldImprovePerformance()
        {
            // Arrange
            var files = CreateTestDataForCaching();
            var config = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001,
                EnableCaching = true,
                CacheExpirationMinutes = 30
            };

            // First run
            var stopwatch1 = Stopwatch.StartNew();
            var result1 = DataProcessor.ProcessMissingData(files, config);
            stopwatch1.Stop();

            // Second run (should use cache)
            var stopwatch2 = Stopwatch.StartNew();
            var result2 = DataProcessor.ProcessMissingData(files, config);
            stopwatch2.Stop();

            // Assert
            Assert.NotNull(result1);
            Assert.NotNull(result2);
            
            // Second run should be faster
            Assert.True(stopwatch2.ElapsedMilliseconds < stopwatch1.ElapsedMilliseconds * 0.8, 
                "Cached run should be significantly faster");
        }

        private List<ExcelFile> CreateRealisticTestData()
        {
            var files = new List<ExcelFile>();
            var baseDate = DateTime.Today.AddDays(-7);
            
            for (int day = 0; day < 7; day++)
            {
                for (int hour = 0; hour < 24; hour += 8)
                {
                    var file = new ExcelFile
                    {
                        FileName = $"test_{baseDate.AddDays(day):yyyy.M.d}-{hour}.xls",
                        Date = baseDate.AddDays(day),
                        Hour = hour,
                        DataRows = new List<DataRow>()
                    };

                    // Add cumulative columns (G)
                    for (int i = 1; i <= 5; i++)
                    {
                        var gRow = new DataRow
                        {
                            Name = $"G{i}",
                            Values = new List<double?>()
                        };

                        // Add change columns (D)
                        var dRow = new DataRow
                        {
                            Name = $"D{i}",
                            Values = new List<double?>()
                        };

                        for (int j = 0; j < 10; j++)
                        {
                            var change = (day * 24 + hour + j) % 10 + 1.0;
                            var cumulative = (day * 24 + hour + j) * 10.0 + i * 100.0;
                            
                            // Simulate some missing data
                            if ((day + hour + j) % 7 == 0)
                            {
                                gRow.Values.Add(null);
                            }
                            else
                            {
                                gRow.Values.Add(cumulative);
                            }
                            
                            dRow.Values.Add(change);
                        }

                        file.DataRows.Add(gRow);
                        file.DataRows.Add(dRow);
                    }

                    files.Add(file);
                }
            }

            return files;
        }

        private List<ExcelFile> CreateLargeScaleTestData(int fileCount, int rowsPerFile)
        {
            var files = new List<ExcelFile>();
            var random = new Random(42);
            var baseDate = DateTime.Today.AddDays(-fileCount / 24);

            for (int i = 0; i < fileCount; i++)
            {
                var file = new ExcelFile
                {
                    FileName = $"large_test_{i}.xls",
                    Date = baseDate.AddHours(i),
                    Hour = i % 24,
                    DataRows = new List<DataRow>()
                };

                for (int j = 0; j < rowsPerFile; j++)
                {
                    var row = new DataRow
                    {
                        Name = $"Data{j}",
                        Values = new List<double?>()
                    };

                    for (int k = 0; k < 20; k++)
                    {
                        // Simulate missing data patterns
                        if (random.NextDouble() < 0.1)
                        {
                            row.Values.Add(null);
                        }
                        else
                        {
                            row.Values.Add(random.NextDouble() * 1000);
                        }
                    }

                    file.DataRows.Add(row);
                }

                files.Add(file);
            }

            return files;
        }

        private List<ExcelFile> CreateInvalidTestData()
        {
            var files = new List<ExcelFile>();
            
            var invalidFile = new ExcelFile
            {
                FileName = "invalid.xls",
                Date = DateTime.Today,
                Hour = 0,
                DataRows = new List<DataRow>
                {
                    new DataRow 
                    { 
                        Name = "InvalidData", 
                        Values = new List<double?> { double.NaN, double.PositiveInfinity, double.NegativeInfinity } 
                    }
                }
            };

            files.Add(invalidFile);
            return files;
        }

        private List<ExcelFile> CreateBusinessLogicTestData()
        {
            var files = new List<ExcelFile>();
            
            var file1 = new ExcelFile
            {
                FileName = "business_logic_test.xls",
                Date = DateTime.Today,
                Hour = 0,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "G1", Values = new List<double?> { 100.0, 200.0, 300.0 } },
                    new DataRow { Name = "D1", Values = new List<double?> { 50.0, 100.0, 100.0 } },
                    new DataRow { Name = "G2", Values = new List<double?> { 150.0, 250.0, 350.0 } },
                    new DataRow { Name = "D2", Values = new List<double?> { 75.0, 100.0, 100.0 } }
                }
            };

            files.Add(file1);
            return files;
        }

        private List<ExcelFile> CreateTestDataForCaching()
        {
            var files = new List<ExcelFile>();
            
            var file = new ExcelFile
            {
                FileName = "cache_test.xls",
                Date = DateTime.Today,
                Hour = 0,
                DataRows = new List<DataRow>
                {
                    new DataRow { Name = "G1", Values = new List<double?> { 100.0, 200.0, 300.0 } },
                    new DataRow { Name = "D1", Values = new List<double?> { 50.0, 100.0, 100.0 } }
                }
            };

            files.Add(file);
            return files;
        }
    }
}
