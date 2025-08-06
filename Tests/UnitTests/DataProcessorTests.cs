using Xunit;
using FluentAssertions;
using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Tests.UnitTests
{
    /// <summary>
    /// 数据处理器单元测试
    /// </summary>
    public class DataProcessorTests
    {
        [Fact]
        public void ProcessMissingData_EmptyList_ShouldReturnEmptyList()
        {
            // Arrange
            var files = new List<ExcelFile>();

            // Act
            var result = DataProcessor.ProcessMissingData(files);

            // Assert
            result.Should().NotBeNull();
            result.Should().BeEmpty();
        }

        [Fact]
        public void ProcessMissingData_NullList_ShouldReturnEmptyList()
        {
            // Act
            var result = DataProcessor.ProcessMissingData(null);

            // Assert
            result.Should().NotBeNull();
            result.Should().BeEmpty();
        }

        [Fact]
        public void ProcessMissingData_FilesWithMissingData_ShouldFillMissingValues()
        {
            // Arrange
            var files = CreateTestFilesWithMissingData();

            // Act
            var result = DataProcessor.ProcessMissingData(files);

            // Assert
            result.Should().HaveCount(3);
            
            // 检查中间文件的缺失数据是否被补充
            var middleFile = result.FirstOrDefault(f => f.Hour == 8);
            middleFile.Should().NotBeNull();
            
            var testDataRow = middleFile!.DataRows.FirstOrDefault(r => r.Name == "测试数据");
            testDataRow.Should().NotBeNull();
            testDataRow!.Values[0].Should().BeApproximately(15.0, 0.01); // (10 + 20) / 2
        }

        [Fact]
        public void ProcessMissingData_SameDayComplement_ShouldUseAverageFromSameDay()
        {
            // Arrange
            var files = CreateTestFilesForSameDayComplement();

            // Act
            var result = DataProcessor.ProcessMissingData(files);

            // Assert
            var targetFile = result.FirstOrDefault(f => f.Hour == 8);
            targetFile.Should().NotBeNull();
            
            var testDataRow = targetFile!.DataRows.FirstOrDefault(r => r.Name == "温度");
            testDataRow.Should().NotBeNull();
            testDataRow!.Values[0].Should().BeApproximately(25.0, 0.01); // (20 + 30) / 2
        }

        [Fact]
        public void ProcessMissingData_SingleNearestValue_ShouldUseSingleValue()
        {
            // Arrange
            var files = CreateTestFilesForSingleNearestValue();

            // Act
            var result = DataProcessor.ProcessMissingData(files);

            // Assert
            var targetFile = result.FirstOrDefault(f => f.Hour == 8);
            targetFile.Should().NotBeNull();
            
            var testDataRow = targetFile!.DataRows.FirstOrDefault(r => r.Name == "湿度");
            testDataRow.Should().NotBeNull();
            testDataRow!.Values[0].Should().Be(75.0); // 使用前一个有效值
        }

        [Fact]
        public void CheckCompleteness_EmptyList_ShouldReturnEmptyResult()
        {
            // Arrange
            var files = new List<ExcelFile>();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IncompleteDates.Should().BeEmpty();
            result.DateCompleteness.Should().BeEmpty();
            result.IsAllComplete.Should().BeTrue();
        }

        [Fact]
        public void CheckCompleteness_CompleteDays_ShouldShowFullCompleteness()
        {
            // Arrange
            var files = CreateCompleteTestFiles();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IncompleteDates.Should().BeEmpty();
            result.IsAllComplete.Should().BeTrue();
            result.DateCompleteness.Should().HaveCount(1);
            result.DateCompleteness[0].IsComplete.Should().BeTrue();
        }

        [Fact]
        public void CheckCompleteness_IncompleteDays_ShouldIdentifyMissingFiles()
        {
            // Arrange
            var files = CreateIncompleteTestFiles();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IncompleteDates.Should().NotBeEmpty();
            result.IsAllComplete.Should().BeFalse();
            
            // 应该识别出缺失的时间点
            var dateCompleteness = result.DateCompleteness.FirstOrDefault();
            dateCompleteness.Should().NotBeNull();
            dateCompleteness!.IsComplete.Should().BeFalse();
            dateCompleteness.MissingHours.Should().Contain(0); // 缺少0点文件
        }

        [Fact]
        public void GenerateSupplementFiles_CompleteDays_ShouldReturnEmptyList()
        {
            // Arrange
            var files = CreateCompleteTestFiles();

            // Act
            var result = DataProcessor.GenerateSupplementFiles(files);

            // Assert
            result.Should().BeEmpty();
        }

        [Fact]
        public void GenerateSupplementFiles_IncompleteDays_ShouldGenerateSupplementFiles()
        {
            // Arrange
            var files = CreateIncompleteTestFiles();

            // Act
            var result = DataProcessor.GenerateSupplementFiles(files);

            // Assert
            result.Should().NotBeEmpty();
            
            var supplementFile = result.FirstOrDefault();
            supplementFile.Should().NotBeNull();
            supplementFile!.SourceFile.Should().NotBeNull();
            supplementFile.TargetFileName.Should().NotBeNullOrEmpty();
        }

        [Fact]
        public void ValidateDataQuality_ValidData_ShouldReturnGoodQuality()
        {
            // Arrange
            var files = CreateTestFilesWithValidData();

            // Act
            var result = DataProcessor.ValidateDataQuality(files);

            // Assert
            result.Should().NotBeNull();
            result.OverallCompleteness.Should().BeGreaterThan(80.0);
            result.FileQuality.Should().NotBeEmpty();
        }

        [Fact]
        public void ValidateDataQuality_PoorData_ShouldIdentifyIssues()
        {
            // Arrange
            var files = CreateTestFilesWithPoorData();

            // Act
            var result = DataProcessor.ValidateDataQuality(files);

            // Assert
            result.Should().NotBeNull();
            result.OverallCompleteness.Should().BeLessThan(50.0);
            result.MissingRows.Should().BeGreaterThan(0);
        }

        #region Helper Methods

        private List<ExcelFile> CreateTestFilesWithMissingData()
        {
            var baseDate = new DateTime(2025, 4, 18);
            
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 0,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "测试数据",
                            RowIndex = 5,
                            Values = new List<double?> { 10.0, 5.0, 15.0, 20.0, 25.0, 30.0 }
                        }
                    }
                },
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 8,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "测试数据",
                            RowIndex = 5,
                            Values = new List<double?> { null, 8.0, 18.0, 22.0, 28.0, 32.0 } // 第一个值缺失
                        }
                    }
                },
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "测试数据",
                            RowIndex = 5,
                            Values = new List<double?> { 20.0, 12.0, 22.0, 25.0, 30.0, 35.0 }
                        }
                    }
                }
            };
        }

        private List<ExcelFile> CreateTestFilesForSameDayComplement()
        {
            var baseDate = new DateTime(2025, 4, 18);
            
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 0,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "温度",
                            RowIndex = 5,
                            Values = new List<double?> { 20.0, 22.0, 24.0, 26.0, 28.0, 30.0 }
                        }
                    }
                },
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 8,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "温度",
                            RowIndex = 5,
                            Values = new List<double?> { null, 23.0, 25.0, 27.0, 29.0, 31.0 } // 第一个值缺失
                        }
                    }
                },
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "温度",
                            RowIndex = 5,
                            Values = new List<double?> { 30.0, 32.0, 34.0, 36.0, 38.0, 40.0 }
                        }
                    }
                }
            };
        }

        private List<ExcelFile> CreateTestFilesForSingleNearestValue()
        {
            var baseDate = new DateTime(2025, 4, 18);
            
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 0,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "湿度",
                            RowIndex = 6,
                            Values = new List<double?> { 75.0, 70.0, 65.0, 60.0, 55.0, 50.0 }
                        }
                    }
                },
                new ExcelFile
                {
                    Date = baseDate,
                    Hour = 8,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "湿度",
                            RowIndex = 6,
                            Values = new List<double?> { null, 68.0, 63.0, 58.0, 53.0, 48.0 } // 第一个值缺失
                        }
                    }
                }
            };
        }

        private List<ExcelFile> CreateCompleteTestFiles()
        {
            var baseDate = new DateTime(2025, 4, 18);
            
            return new List<ExcelFile>
            {
                new ExcelFile { Date = baseDate, Hour = 0, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = baseDate, Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = baseDate, Hour = 16, ProjectName = "测试项目.xlsx" }
            };
        }

        private List<ExcelFile> CreateIncompleteTestFiles()
        {
            var baseDate = new DateTime(2025, 4, 18);
            
            return new List<ExcelFile>
            {
                new ExcelFile { Date = baseDate, Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = baseDate, Hour = 16, ProjectName = "测试项目.xlsx" }
                // 缺少 0 点的文件
            };
        }

        private List<ExcelFile> CreateTestFilesWithValidData()
        {
            var files = CreateCompleteTestFiles();
            foreach (var file in files)
            {
                file.DataRows = new List<DataRow>
                {
                    new DataRow
                    {
                        Name = "有效数据",
                        RowIndex = 5,
                        Values = new List<double?> { 10.0, 20.0, 30.0, 40.0, 50.0, 60.0 }
                    }
                };
            }
            return files;
        }

        private List<ExcelFile> CreateTestFilesWithPoorData()
        {
            var files = CreateCompleteTestFiles();
            foreach (var file in files)
            {
                file.DataRows = new List<DataRow>
                {
                    new DataRow
                    {
                        Name = "缺失数据",
                        RowIndex = 5,
                        Values = new List<double?> { null, null, null, null, null, null }
                    }
                };
            }
            return files;
        }

        #endregion
    }
}