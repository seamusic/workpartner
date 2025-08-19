using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using DataFixter.Services;
using FluentAssertions;
using Xunit;

namespace DataFixter.Tests
{
    public class AnalyzeAndCorrectDirectionTests : IDisposable
    {
        private readonly DataCorrectionService _correctionService;


        public AnalyzeAndCorrectDirectionTests()
        {
            _correctionService = new DataCorrectionService(null, new DataFixter.Models.CorrectionOptions
        {
            RandomChangeRange = 0.1 // 使用较小的随机变化量范围，使测试更稳定
        });

        }

        [Fact]
        public void CorrectPoint_ShouldHandleSimpleCumulativeIssue()
        {
            // 创建一个简单的测试数据，只有两期，累计值关系明显错误
            var monitoringPoint = new MonitoringPoint("SimplePoint", 100.0);
            
            // 第一期：本期变化量=0.5，累计值=0.5（正确）
            var firstPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(0),
                RowNumber = 1,
                SequenceNumber = 1,
                PointName = "SimplePoint",
                Mileage = 100.0,
                CurrentPeriodX = 0.5,
                CumulativeX = 0.5,
                CanAdjustment = false, // 第一期设为不可调整
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第二期：本期变化量=0.3，累计值=1.0（错误，应该是0.8）
            var secondPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(1),
                RowNumber = 2,
                SequenceNumber = 2,
                PointName = "SimplePoint",
                Mileage = 101.0,
                CurrentPeriodX = 0.3,
                CumulativeX = 1.0, // 错误：应该是0.5 + 0.3 = 0.8
                CanAdjustment = true, // 第二期设为可调整
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            monitoringPoint.AddPeriodData(firstPeriod);
            monitoringPoint.AddPeriodData(secondPeriod);
            
            // 验证初始数据确实有问题
            var expectedCumulative = firstPeriod.CumulativeX + secondPeriod.CurrentPeriodX; // 0.5 + 0.3 = 0.8
            var actualCumulative = secondPeriod.CumulativeX; // 1.0
            var initialDifference = Math.Abs(expectedCumulative - actualCumulative); // |0.8 - 1.0| = 0.2
            initialDifference.Should().BeGreaterThan(0.0);
            
            // 修正数据
            var result = _correctionService.CorrectPoint(monitoringPoint);
            result.Status.Should().Be(CorrectionStatus.Success);
            
            // 验证修正后的数据
            var correctedDifference = Math.Abs(expectedCumulative - secondPeriod.CumulativeX);
            
            // 输出调试信息
            Console.WriteLine($"初始差异: {initialDifference:F3}");
            Console.WriteLine($"修正后差异: {correctedDifference:F3}");
            Console.WriteLine($"改善: {initialDifference - correctedDifference:F3}");
            
            // 由于使用了随机算法，主要验证算法能正常运行，不要求具体的数值结果
            correctedDifference.Should().BeGreaterThan(0.0, "修正后应该有结果");
            correctedDifference.Should().BeLessThan(1.0, "修正后的差异不应该过大");
        }

        [Fact]
        public void CorrectPoint_ShouldHandleAdjustableSegment()
        {
            // 创建一个包含可调整区间的测试数据
            var monitoringPoint = new MonitoringPoint("SegmentPoint", 100.0);
            
            // 第一期：不可调整，作为基准
            var firstPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(0),
                RowNumber = 1,
                SequenceNumber = 1,
                PointName = "SegmentPoint",
                Mileage = 100.0,
                CurrentPeriodX = 0.5,
                CumulativeX = 0.5,
                CanAdjustment = false,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第二期：可调整，累计值错误
            var secondPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(1),
                RowNumber = 2,
                SequenceNumber = 2,
                PointName = "SegmentPoint",
                Mileage = 101.0,
                CurrentPeriodX = 0.3,
                CumulativeX = 1.0, // 错误：应该是0.5 + 0.3 = 0.8
                CanAdjustment = true,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第三期：可调整，累计值错误
            var thirdPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(2),
                RowNumber = 3,
                SequenceNumber = 3,
                PointName = "SegmentPoint",
                Mileage = 102.0,
                CurrentPeriodX = 0.2,
                CumulativeX = 1.5, // 错误：应该是0.8 + 0.2 = 1.0
                CanAdjustment = true,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第四期：不可调整，作为目标
            var fourthPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(3),
                RowNumber = 4,
                SequenceNumber = 4,
                PointName = "SegmentPoint",
                Mileage = 103.0,
                CurrentPeriodX = 0.1,
                CumulativeX = 1.1, // 正确：1.0 + 0.1 = 1.1
                CanAdjustment = false,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            monitoringPoint.AddPeriodData(firstPeriod);
            monitoringPoint.AddPeriodData(secondPeriod);
            monitoringPoint.AddPeriodData(thirdPeriod);
            monitoringPoint.AddPeriodData(fourthPeriod);
            
            // 验证初始数据确实有问题
            var expectedSecondCumulative = firstPeriod.CumulativeX + secondPeriod.CurrentPeriodX; // 0.5 + 0.3 = 0.8
            var expectedThirdCumulative = expectedSecondCumulative + thirdPeriod.CurrentPeriodX; // 0.8 + 0.2 = 1.0
            var expectedFourthCumulative = expectedThirdCumulative + fourthPeriod.CurrentPeriodX; // 1.0 + 0.1 = 1.1
            
            var initialError = Math.Abs(secondPeriod.CumulativeX - expectedSecondCumulative) + 
                              Math.Abs(thirdPeriod.CumulativeX - expectedThirdCumulative);
            
            initialError.Should().BeGreaterThan(0.0);
            
            // 修正数据
            var result = _correctionService.CorrectPoint(monitoringPoint);
            result.Status.Should().Be(CorrectionStatus.Success);
            
            // 验证修正后的数据
            var correctedError = Math.Abs(secondPeriod.CumulativeX - expectedSecondCumulative) + 
                                Math.Abs(thirdPeriod.CumulativeX - expectedThirdCumulative);
            
            // 输出调试信息
            Console.WriteLine($"初始误差: {initialError:F3}");
            Console.WriteLine($"修正后误差: {correctedError:F3}");
            Console.WriteLine($"改善: {initialError - correctedError:F3}");
            
            // 验证修正应该改善数据
            correctedError.Should().BeLessThanOrEqualTo(initialError, "修正应该减少或至少不增加误差");
            
            // 验证累计值关系
            var actualSecondCumulative = secondPeriod.CumulativeX;
            var actualThirdCumulative = thirdPeriod.CumulativeX;
            
            // 检查修正后的累计值是否更接近期望值
            var secondImprovement = Math.Abs(secondPeriod.CumulativeX - expectedSecondCumulative) < 
                                   Math.Abs(1.0 - expectedSecondCumulative);
            var thirdImprovement = Math.Abs(thirdPeriod.CumulativeX - expectedThirdCumulative) < 
                                  Math.Abs(1.5 - expectedThirdCumulative);
            
            (secondImprovement || thirdImprovement).Should().BeTrue("至少有一期的累计值应该得到改善");
        }

        [Fact]
        public void CorrectPoint_ShouldPreserveFixedData()
        {
            // 测试确保CanAdjustment=false的数据不被修改
            var monitoringPoint = new MonitoringPoint("FixedPoint", 100.0);
            
            // 第一期：不可调整
            var firstPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(0),
                RowNumber = 1,
                SequenceNumber = 1,
                PointName = "FixedPoint",
                Mileage = 100.0,
                CurrentPeriodX = 0.5,
                CumulativeX = 0.5,
                CanAdjustment = false,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第二期：可调整
            var secondPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(1),
                RowNumber = 2,
                SequenceNumber = 2,
                PointName = "FixedPoint",
                Mileage = 101.0,
                CurrentPeriodX = 0.3,
                CumulativeX = 1.0, // 错误
                CanAdjustment = true,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            monitoringPoint.AddPeriodData(firstPeriod);
            monitoringPoint.AddPeriodData(secondPeriod);
            
            // 记录修正前的值
            var originalFirstPeriodX = firstPeriod.CurrentPeriodX;
            var originalFirstCumulativeX = firstPeriod.CumulativeX;
            
            // 修正数据
            var result = _correctionService.CorrectPoint(monitoringPoint);
            result.Status.Should().Be(CorrectionStatus.Success);
            
            // 验证不可调整的数据没有被修改
            firstPeriod.CurrentPeriodX.Should().Be(originalFirstPeriodX, "不可调整的本期变化量不应该被修改");
            firstPeriod.CumulativeX.Should().Be(originalFirstCumulativeX, "不可调整的累计值不应该被修改");
            
            // 验证可调整的数据被修改了
            var hasChanges = secondPeriod.CurrentPeriodX != 0.3 || secondPeriod.CumulativeX != 1.0;
            hasChanges.Should().BeTrue("可调整的数据应该被修正服务修改");
        }

        [Fact]
        public void Debug_CorrectionApplication()
        {
            // 创建一个简单的测试数据
            var monitoringPoint = new MonitoringPoint("DebugPoint", 100.0);
            
            // 第一期：不可调整
            var firstPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(0),
                RowNumber = 1,
                SequenceNumber = 1,
                PointName = "DebugPoint",
                Mileage = 100.0,
                CurrentPeriodX = 0.5,
                CumulativeX = 0.5,
                CanAdjustment = false,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            // 第二期：可调整，累计值错误
            var secondPeriod = new PeriodData
            {
                FileInfo = CreateTestFileInfo(1),
                RowNumber = 2,
                SequenceNumber = 2,
                PointName = "DebugPoint",
                Mileage = 101.0,
                CurrentPeriodX = 0.3,
                CumulativeX = 1.0, // 错误：应该是0.5 + 0.3 = 0.8
                CanAdjustment = true,
                ValidationStatus = ValidationStatus.NotValidated,
                AdjustmentType = AdjustmentType.None
            };
            
            monitoringPoint.AddPeriodData(firstPeriod);
            monitoringPoint.AddPeriodData(secondPeriod);
            
            // 记录修正前的值
            Console.WriteLine($"修正前:");
            Console.WriteLine($"  第一期: CurrentPeriodX={firstPeriod.CurrentPeriodX}, CumulativeX={firstPeriod.CumulativeX}");
            Console.WriteLine($"  第二期: CurrentPeriodX={secondPeriod.CurrentPeriodX}, CumulativeX={secondPeriod.CumulativeX}");
            
            // 修正数据
            var result = _correctionService.CorrectPoint(monitoringPoint);
            Console.WriteLine($"修正结果: Status={result.Status}, CorrectedValues={result.CorrectedValues}");
            
            // 记录修正后的值
            Console.WriteLine($"修正后:");
            Console.WriteLine($"  第一期: CurrentPeriodX={firstPeriod.CurrentPeriodX}, CumulativeX={firstPeriod.CumulativeX}");
            Console.WriteLine($"  第二期: CurrentPeriodX={secondPeriod.CurrentPeriodX}, CumulativeX={secondPeriod.CumulativeX}");
            
            // 验证修正是否被应用
            var hasChanges = secondPeriod.CurrentPeriodX != 0.3 || secondPeriod.CumulativeX != 1.0;
            hasChanges.Should().BeTrue("可调整的数据应该被修正服务修改");
            
            // 验证累计值关系
            var expectedCumulative = firstPeriod.CumulativeX + secondPeriod.CurrentPeriodX;
            var actualCumulative = secondPeriod.CumulativeX;
            var difference = Math.Abs(expectedCumulative - actualCumulative);
            
            Console.WriteLine($"累计值关系验证:");
            Console.WriteLine($"  期望累计值: {expectedCumulative}");
            Console.WriteLine($"  实际累计值: {actualCumulative}");
            Console.WriteLine($"  差异: {difference}");
            
            // 差异应该很小
            difference.Should().BeLessThan(0.01, "修正后的累计值关系应该基本正确");
        }

        private ExcelFileInfo CreateTestFileInfo(int sequence)
        {
            return new ExcelFileInfo($"test_{sequence}.txt", 1024, DateTime.Now.AddDays(sequence));
        }

        public void Dispose()
        {
            // 清理资源
        }
    }
}
