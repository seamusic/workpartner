using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Tests.PerformanceTests
{
    public class DataProcessorPerformanceTests
    {
        private readonly DataProcessorConfig _defaultConfig;

        public DataProcessorPerformanceTests()
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
        public void PerformanceTest_SmallDataset_ShouldCompleteWithinTimeLimit()
        {
            // Arrange
            var files = CreatePerformanceTestData(100, 20);
            var stopwatch = Stopwatch.StartNew();

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            stopwatch.Stop();

            // Assert
            Assert.NotNull(result);
            Assert.Equal(files.Count, result.Count);
            
            // Performance assertion for small dataset
            Assert.True(stopwatch.ElapsedMilliseconds < 2000, 
                $"Small dataset processing took too long: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void PerformanceTest_MediumDataset_ShouldCompleteWithinTimeLimit()
        {
            // Arrange
            var files = CreatePerformanceTestData(500, 50);
            var stopwatch = Stopwatch.StartNew();

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            stopwatch.Stop();

            // Assert
            Assert.NotNull(result);
            Assert.Equal(files.Count, result.Count);
            
            // Performance assertion for medium dataset (500 files × 50 rows = 25,000 data points)
            // Allow more time for complex data processing operations
            Assert.True(stopwatch.ElapsedMilliseconds < 120000, 
                $"Medium dataset processing took too long: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void PerformanceTest_LargeDataset_ShouldCompleteWithinTimeLimit()
        {
            // Arrange
            var files = CreatePerformanceTestData(1000, 100);
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
            
            // Performance assertion for large dataset
            Assert.True(stopwatch.ElapsedMilliseconds < 30000, 
                $"Large dataset processing took too long: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void CachePerformanceTest_WithAndWithoutCache_ShouldShowImprovement()
        {
            // Arrange
            var files = CreatePerformanceTestData(200, 30);
            
            // Test without cache
            var configWithoutCache = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001,
                EnableCaching = false
            };

            var stopwatch1 = Stopwatch.StartNew();
            var result1 = DataProcessor.ProcessMissingData(files, configWithoutCache);
            stopwatch1.Stop();

            // Test with cache
            var configWithCache = new DataProcessorConfig
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

            var stopwatch2 = Stopwatch.StartNew();
            var result2 = DataProcessor.ProcessMissingData(files, configWithCache);
            stopwatch2.Stop();

            // Assert
            Assert.NotNull(result1);
            Assert.NotNull(result2);
            
            // Cache should provide performance improvement
            var improvement = (double)stopwatch1.ElapsedMilliseconds / stopwatch2.ElapsedMilliseconds;
            Assert.True(improvement > 1.0, $"Cache should improve performance. Improvement ratio: {improvement:F2}");
        }

        [Fact]
        public void MemoryUsageTest_LargeDataset_ShouldRemainWithinLimits()
        {
            // Arrange
            var files = CreatePerformanceTestData(800, 80);
            var initialMemory = GC.GetTotalMemory(false);

            // Act
            var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
            var finalMemory = GC.GetTotalMemory(false);

            // Assert
            Assert.NotNull(result);
            
            // Memory usage should be reasonable
            var memoryIncrease = finalMemory - initialMemory;
            var memoryIncreaseMB = memoryIncrease / (1024 * 1024);
            
            Assert.True(memoryIncreaseMB < 200, $"Memory increase too high: {memoryIncreaseMB:F2}MB");
        }

        [Fact]
        public void BatchProcessingTest_DifferentBatchSizes_ShouldShowOptimalPerformance()
        {
            // Arrange
            var files = CreatePerformanceTestData(600, 60);
            var batchSizes = new[] { 50, 100, 200, 500 };
            var results = new Dictionary<int, long>();

            foreach (var batchSize in batchSizes)
            {
                var config = new DataProcessorConfig
                {
                    CumulativeColumnPrefix = "G",
                    ChangeColumnPrefix = "D",
                    AdjustmentRange = 0.05,
                    RandomSeed = 42,
                    TimeFactorWeight = 1.0,
                    MinimumAdjustment = 0.001,
                    BatchSize = batchSize,
                    EnableCaching = true,
                    CacheExpirationMinutes = 30
                };

                var stopwatch = Stopwatch.StartNew();
                var result = DataProcessor.ProcessMissingData(files, config);
                stopwatch.Stop();

                Assert.NotNull(result);
                results[batchSize] = stopwatch.ElapsedMilliseconds;
            }

            // Assert that batch processing provides reasonable performance
            Assert.All(results.Values, time => Assert.True(time < 15000, $"Batch processing took too long: {time}ms"));
        }

        [Fact]
        public void StressTest_ConcurrentProcessing_ShouldHandleMultipleRequests()
        {
            // Arrange
            var testDataSets = new List<List<ExcelFile>>();
            for (int i = 0; i < 5; i++)
            {
                testDataSets.Add(CreatePerformanceTestData(100, 20));
            }

            var stopwatch = Stopwatch.StartNew();
            var tasks = new List<Task<List<ExcelFile>>>();

            // Act - Process multiple datasets concurrently
            foreach (var dataSet in testDataSets)
            {
                var task = Task.Run(() => DataProcessor.ProcessMissingData(dataSet, _defaultConfig));
                tasks.Add(task);
            }

            // Wait for all tasks to complete
            Task.WaitAll(tasks.ToArray());
            stopwatch.Stop();

            // Assert
            Assert.All(tasks, task => Assert.True(task.IsCompletedSuccessfully));
            Assert.All(tasks, task => Assert.NotNull(task.Result));
            
            // Concurrent processing should complete within reasonable time
            Assert.True(stopwatch.ElapsedMilliseconds < 10000, 
                $"Concurrent processing took too long: {stopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void PerformanceRegressionTest_CompareWithBaseline_ShouldNotDegrade()
        {
            // Arrange
            var files = CreatePerformanceTestData(300, 40);
            var baselineConfig = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001,
                BatchSize = 50,
                EnableCaching = false
            };

            var optimizedConfig = new DataProcessorConfig
            {
                CumulativeColumnPrefix = "G",
                ChangeColumnPrefix = "D",
                AdjustmentRange = 0.05,
                RandomSeed = 42,
                TimeFactorWeight = 1.0,
                MinimumAdjustment = 0.001,
                BatchSize = 200,
                EnableCaching = false,
                CacheExpirationMinutes = 30
            };

            // Baseline performance
            var baselineStopwatch = Stopwatch.StartNew();
            var baselineResult = DataProcessor.ProcessMissingData(files, baselineConfig);
            baselineStopwatch.Stop();

            // Optimized performance
            var optimizedStopwatch = Stopwatch.StartNew();
            var optimizedResult = DataProcessor.ProcessMissingData(files, optimizedConfig);
            optimizedStopwatch.Stop();

            // Assert
            Assert.NotNull(baselineResult);
            Assert.NotNull(optimizedResult);
            
            // Optimized version should not be slower than baseline
            Assert.True(optimizedStopwatch.ElapsedMilliseconds <= baselineStopwatch.ElapsedMilliseconds * 1.1, 
                $"Optimized version should not be significantly slower than baseline. " +
                $"Baseline: {baselineStopwatch.ElapsedMilliseconds}ms, " +
                $"Optimized: {optimizedStopwatch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void ScalabilityTest_IncreasingDataSize_ShouldShowLinearGrowth()
        {
            // Arrange
            var dataSizes = new[] { 100, 200, 400, 800 };
            var results = new Dictionary<int, long>();

            foreach (var size in dataSizes)
            {
                var files = CreatePerformanceTestData(size, size / 10);
                var stopwatch = Stopwatch.StartNew();
                
                var result = DataProcessor.ProcessMissingData(files, _defaultConfig);
                stopwatch.Stop();

                Assert.NotNull(result);
                results[size] = stopwatch.ElapsedMilliseconds;
            }

            // Assert that performance scales reasonably with data size
            var previousTime = results[dataSizes[0]];
            for (int i = 1; i < dataSizes.Length; i++)
            {
                var currentTime = results[dataSizes[i]];
                var dataSizeRatio = (double)dataSizes[i] / dataSizes[i - 1];
                var timeRatio = (double)currentTime / previousTime;
                
                // Time increase should not be more than 3x the data size increase
                Assert.True(timeRatio <= dataSizeRatio * 3, 
                    $"Performance does not scale well. Data size ratio: {dataSizeRatio:F2}, Time ratio: {timeRatio:F2}");
                
                previousTime = currentTime;
            }
        }

        private List<ExcelFile> CreatePerformanceTestData(int fileCount, int rowsPerFile)
        {
            var files = new List<ExcelFile>();
            var random = new Random(42);
            var baseDate = DateTime.Today.AddDays(-fileCount / 24);

            for (int i = 0; i < fileCount; i++)
            {
                var file = new ExcelFile
                {
                    FileName = $"perf_test_{i}.xls",
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

                    for (int k = 0; k < 15; k++)
                    {
                        // Simulate realistic missing data patterns
                        if (random.NextDouble() < 0.15)
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
    }
}
