using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataFixter.Excel;
using DataFixter.Models;
using DataFixter.Services;
using Microsoft.Extensions.Logging;
using Xunit;

namespace DataFixter.Tests
{
    /// <summary>
    /// 数据处理测试类
    /// </summary>
    public class DataProcessingTests
    {
        private readonly ILogger<DataProcessingTests> _logger;

        public DataProcessingTests()
        {
            // 创建测试用的日志记录器
            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder.AddConsole();
            });
            _logger = loggerFactory.CreateLogger<DataProcessingTests>();
        }

        [Fact]
        public void TestDataNormalization()
        {
            // 创建测试数据
            var testData = CreateTestData();
            
            // 创建数据标准化器
            var normalizer = new DataNormalizer(_logger);
            
            // 执行标准化
            var normalizedData = normalizer.NormalizeData(testData);
            
            // 验证结果
            Assert.NotNull(normalizedData);
            Assert.Equal(3, normalizedData.Count); // 应该有3行有效数据
            
            // 验证第一行数据
            var firstRow = normalizedData.First();
            Assert.Equal("测试点1", firstRow.PointName);
            Assert.Equal(100.50, firstRow.Mileage);
            Assert.Equal(0.001, firstRow.CurrentPeriodX, 3);
            Assert.Equal(0.002, firstRow.CurrentPeriodY, 3);
            Assert.Equal(0.003, firstRow.CurrentPeriodZ, 3);
        }

        [Fact]
        public void TestDataGrouping()
        {
            // 创建测试数据
            var testData = CreateTestPeriodData();
            
            // 创建数据分组服务
            var groupingService = new DataGroupingService(_logger);
            
            // 执行分组
            var monitoringPoints = groupingService.GroupByPointName(testData);
            
            // 验证结果
            Assert.NotNull(monitoringPoints);
            Assert.Equal(2, monitoringPoints.Count); // 应该有2个监测点
            
            // 验证第一个监测点
            var firstPoint = monitoringPoints.First();
            Assert.Equal("测试点1", firstPoint.PointName);
            Assert.Equal(100.50, firstPoint.Mileage);
            Assert.Equal(2, firstPoint.PeriodDataCount); // 应该有2期数据
            
            // 验证第二个监测点
            var secondPoint = monitoringPoints.Last();
            Assert.Equal("测试点2", secondPoint.PointName);
            Assert.Equal(200.75, secondPoint.Mileage);
            Assert.Equal(1, secondPoint.PeriodDataCount); // 应该有1期数据
        }

        [Fact]
        public void TestDataIntegrityCheck()
        {
            // 创建测试数据
            var testData = CreateTestPeriodData();
            var groupingService = new DataGroupingService(_logger);
            var monitoringPoints = groupingService.GroupByPointName(testData);
            
            // 执行完整性检查
            var integrityReport = groupingService.CheckDataIntegrity(monitoringPoints);
            
            // 验证结果
            Assert.NotNull(integrityReport);
            Assert.Equal(2, integrityReport.TotalPoints);
            Assert.Equal(3, integrityReport.TotalPeriods);
            Assert.Equal(0, integrityReport.CriticalIssueCount); // 应该没有严重问题
        }

        [Fact]
        public void TestDataSorting()
        {
            // 创建测试数据
            var testData = CreateTestPeriodData();
            var groupingService = new DataGroupingService(_logger);
            var monitoringPoints = groupingService.GroupByPointName(testData);
            
            // 执行时间排序
            groupingService.SortAllPointsByTime(monitoringPoints);
            
            // 验证排序结果
            foreach (var point in monitoringPoints)
            {
                if (point.PeriodDataCount > 1)
                {
                    // 验证时间顺序
                    for (int i = 1; i < point.PeriodDataCount; i++)
                    {
                        var prevTime = point.PeriodDataList[i - 1].FileInfo?.FullDateTime;
                        var currTime = point.PeriodDataList[i].FileInfo?.FullDateTime;
                        
                        if (prevTime.HasValue && currTime.HasValue)
                        {
                            Assert.True(prevTime.Value <= currTime.Value, 
                                $"时间排序错误: {prevTime.Value} 应该小于等于 {currTime.Value}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 创建测试用的Excel读取结果
        /// </summary>
        /// <returns>测试数据</returns>
        private List<ExcelReadResult> CreateTestData()
        {
            var results = new List<ExcelReadResult>();
            
            // 创建第一个文件的结果
            var firstResult = new ExcelReadResult
            {
                FilePath = "test1.xls",
                IsSuccess = true,
                SheetName = "Sheet1",
                DataRows = new List<ExcelDataRow>
                {
                    new ExcelDataRow
                    {
                        RowNumber = 5,
                        FileInfo = new FileInfo("test1.xls", 1024, DateTime.Now),
                        PointName = "测试点1",
                        Mileage = 100.50,
                        CurrentPeriodX = 0.001,
                        CurrentPeriodY = 0.002,
                        CurrentPeriodZ = 0.003,
                        CumulativeX = 1.001,
                        CumulativeY = 1.002,
                        CumulativeZ = 1.003,
                        DailyX = 0.0001,
                        DailyY = 0.0002,
                        DailyZ = 0.0003
                    },
                    new ExcelDataRow
                    {
                        RowNumber = 6,
                        FileInfo = new FileInfo("test1.xls", 1024, DateTime.Now),
                        PointName = "测试点2",
                        Mileage = 200.75,
                        CurrentPeriodX = 0.004,
                        CurrentPeriodY = 0.005,
                        CurrentPeriodZ = 0.006,
                        CumulativeX = 2.004,
                        CumulativeY = 2.005,
                        CumulativeZ = 2.006,
                        DailyX = 0.0004,
                        DailyY = 0.0005,
                        DailyZ = 0.0006
                    }
                }
            };
            
            // 创建第二个文件的结果
            var secondResult = new ExcelReadResult
            {
                FilePath = "test2.xls",
                IsSuccess = true,
                SheetName = "Sheet1",
                DataRows = new List<ExcelDataRow>
                {
                    new ExcelDataRow
                    {
                        RowNumber = 5,
                        FileInfo = new FileInfo("test2.xls", 1024, DateTime.Now.AddDays(1)),
                        PointName = "测试点1",
                        Mileage = 100.50,
                        CurrentPeriodX = 0.007,
                        CurrentPeriodY = 0.008,
                        CurrentPeriodZ = 0.009,
                        CumulativeX = 1.008,
                        CumulativeY = 1.009,
                        CumulativeZ = 1.010,
                        DailyX = 0.0007,
                        DailyY = 0.0008,
                        DailyZ = 0.0009
                    }
                }
            };
            
            results.Add(firstResult);
            results.Add(secondResult);
            
            return results;
        }

        /// <summary>
        /// 创建测试用的期数据
        /// </summary>
        /// <returns>测试数据</returns>
        private List<PeriodData> CreateTestPeriodData()
        {
            var data = new List<PeriodData>();
            
            // 第一个监测点的第一期数据
            data.Add(new PeriodData
            {
                FileInfo = new FileInfo("test1.xls", 1024, DateTime.Now),
                RowNumber = 5,
                PointName = "测试点1",
                Mileage = 100.50,
                CurrentPeriodX = 0.001,
                CurrentPeriodY = 0.002,
                CurrentPeriodZ = 0.003,
                CumulativeX = 1.001,
                CumulativeY = 1.002,
                CumulativeZ = 1.003,
                DailyX = 0.0001,
                DailyY = 0.0002,
                DailyZ = 0.0003
            });
            
            // 第一个监测点的第二期数据
            data.Add(new PeriodData
            {
                FileInfo = new FileInfo("test2.xls", 1024, DateTime.Now.AddDays(1)),
                RowNumber = 5,
                PointName = "测试点1",
                Mileage = 100.50,
                CurrentPeriodX = 0.007,
                CurrentPeriodY = 0.008,
                CurrentPeriodZ = 0.009,
                CumulativeX = 1.008,
                CumulativeY = 1.009,
                CumulativeZ = 1.010,
                DailyX = 0.0007,
                DailyY = 0.0008,
                DailyZ = 0.0009
            });
            
            // 第二个监测点的数据
            data.Add(new PeriodData
            {
                FileInfo = new FileInfo("test1.xls", 1024, DateTime.Now),
                RowNumber = 6,
                PointName = "测试点2",
                Mileage = 200.75,
                CurrentPeriodX = 0.004,
                CurrentPeriodY = 0.005,
                CurrentPeriodZ = 0.006,
                CumulativeX = 2.004,
                CumulativeY = 2.005,
                CumulativeZ = 2.006,
                DailyX = 0.0004,
                DailyY = 0.0005,
                DailyZ = 0.0006
            });
            
            return data;
        }
    }
}
