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
        public void ProcessMissingData_LargeDataset_ShouldProcessEfficiently()
        {
            // Arrange - 创建大量测试数据
            var files = new List<ExcelFile>();
            var random = new Random(42); // 固定种子以确保可重复性
            
            // 创建100个文件，每个文件有10行数据，每行有50个值
            for (int i = 0; i < 100; i++)
            {
                var file = new ExcelFile
                {
                    Date = new DateTime(2025, 4, 18).AddDays(i % 30),
                    Hour = (i % 3) * 8, // 0, 8, 16
                    ProjectName = "测试项目.xlsx",
                    FileName = $"2025.4.{18 + (i % 30)}-{(i % 3) * 8:D2}测试项目.xlsx",
                    FilePath = $@"C:\Test\2025.4.{18 + (i % 30)}-{(i % 3) * 8:D2}测试项目.xlsx"
                };

                var dataRows = new List<DataRow>();
                for (int row = 0; row < 10; row++)
                {
                    var values = new List<double?>();
                    for (int col = 0; col < 50; col++)
                    {
                        // 随机设置一些值为null（缺失数据）
                        if (random.Next(100) < 20) // 20%的概率为null
                        {
                            values.Add(null);
                        }
                        else
                        {
                            values.Add(random.NextDouble() * 100);
                        }
                    }
                    
                    dataRows.Add(new DataRow
                    {
                        Name = $"数据行{row + 1}",
                        Values = values
                    });
                }
                
                file.DataRows = dataRows;
                files.Add(file);
            }

            // Act
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var result = DataProcessor.ProcessMissingData(files);
            stopwatch.Stop();

            // Assert
            result.Should().NotBeNull();
            result.Should().HaveCount(100);
            
            // 验证处理时间在合理范围内（对于100个文件，应该在几秒内完成）
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(10000); // 10秒内
            
            // 验证所有缺失值都被补充了
            var totalValues = result.SelectMany(f => f.DataRows).Sum(r => r.Values.Count);
            var nullValues = result.SelectMany(f => f.DataRows).Sum(r => r.Values.Count(v => !v.HasValue));
            nullValues.Should().Be(0); // 所有缺失值都应该被补充
        }

        [Fact]
        public void ProcessMissingData_AllEmptyDataRows_ShouldUseAdjacentRowsAverage()
        {
            // Arrange - 创建测试数据，模拟您提供的场景
            var files = new List<ExcelFile>();
            
            // 创建3个文件，模拟不同时间点的数据
            for (int i = 0; i < 3; i++)
            {
                var file = new ExcelFile
                {
                    Date = new DateTime(2025, 4, 18),
                    Hour = i * 8, // 0, 8, 16
                    ProjectName = "测试项目.xlsx",
                    FileName = $"2025.4.18-{i * 8:D2}测试项目.xlsx",
                    FilePath = $@"C:\Test\2025.4.18-{i * 8:D2}测试项目.xlsx"
                };

                var dataRows = new List<DataRow>();
                
                // 模拟您提供的数据结构
                // 02P19Z033 - 有数据
                dataRows.Add(new DataRow
                {
                    Name = "02P19Z033",
                    Values = new List<double?> { -0.22, -0.15, 0.15, -0.82, -2.09, 0.61 }
                });
                
                // 02P19Z034 - 有数据
                dataRows.Add(new DataRow
                {
                    Name = "02P19Z034",
                    Values = new List<double?> { 0.27, 0.29, -0.02, -0.36, -1.10, -1.77 }
                });
                
                // 02P19Z041 - 所有文件都为空（这是我们要测试的）
                dataRows.Add(new DataRow
                {
                    Name = "02P19Z041",
                    Values = new List<double?> { null, null, null, null, null, null }
                });
                
                // 02P19Z042 - 有数据
                dataRows.Add(new DataRow
                {
                    Name = "02P19Z042",
                    Values = new List<double?> { -0.08, 0.15, -0.01, -1.52, -0.15, 1.60 }
                });
                
                // 02P19Z043 - 有数据
                dataRows.Add(new DataRow
                {
                    Name = "02P19Z043",
                    Values = new List<double?> { 0.08, -0.05, 0.08, -1.52, 1.69, 2.54 }
                });
                
                file.DataRows = dataRows;
                files.Add(file);
            }

            // Act
            var result = DataProcessor.ProcessMissingData(files);

            // Assert
            result.Should().NotBeNull();
            result.Should().HaveCount(3);
            
            // 验证02P19Z041行的数据已经被补充
            foreach (var file in result)
            {
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == "02P19Z041");
                dataRow.Should().NotBeNull();
                
                // 验证所有值都已经被补充（不再是null）
                dataRow!.Values.Should().NotContainNulls();
                
                // 验证补充的值是前一行和后一行的平均值
                // 前一行02P19Z034的平均值：(0.27 + 0.29 + (-0.02) + (-0.36) + (-1.10) + (-1.77)) / 6 = -0.448
                // 后一行02P19Z042的平均值：(-0.08 + 0.15 + (-0.01) + (-1.52) + (-0.15) + 1.60) / 6 = 0.000
                // 期望的平均值：(-0.448 + 0.000) / 2 = -0.224
                
                var expectedAverage = -0.224; // 计算出的期望平均值
                var actualAverage = dataRow.Values.Average();
                actualAverage.Should().BeApproximately(expectedAverage, 0.01);
            }
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

        [Fact]
        public void GetPreviousObservationTime_WithPreviousFile_ShouldReturnCorrectTime()
        {
            // Arrange
            var allFiles = CreateTestFilesForA2Update();
            var supplementFile = new SupplementFileInfo
            {
                TargetDate = new DateTime(2025, 4, 16),
                TargetHour = 8,
                ProjectName = "测试项目"
            };

            // Act
            var result = DataProcessor.GetPreviousObservationTime(supplementFile, allFiles);

            // Assert
            result.Should().Be("2025-4-16 00:00");
        }

        [Fact]
        public void GetPreviousObservationTime_FirstFile_ShouldReturnSameTime()
        {
            // Arrange
            var allFiles = CreateTestFilesForA2Update();
            var supplementFile = new SupplementFileInfo
            {
                TargetDate = new DateTime(2025, 4, 15),
                TargetHour = 0,
                ProjectName = "测试项目"
            };

            // Act
            var result = DataProcessor.GetPreviousObservationTime(supplementFile, allFiles);

            // Assert
            result.Should().Be("2025-4-15 00:00");
        }

        [Fact]
        public void GetPreviousObservationTime_NoPreviousFile_ShouldReturnSameTime()
        {
            // Arrange
            var allFiles = new List<ExcelFile>();
            var supplementFile = new SupplementFileInfo
            {
                TargetDate = new DateTime(2025, 4, 15),
                TargetHour = 0,
                ProjectName = "测试项目"
            };

            // Act
            var result = DataProcessor.GetPreviousObservationTime(supplementFile, allFiles);

            // Assert
            result.Should().Be("2025-4-15 00:00");
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

        private List<ExcelFile> CreateTestFilesForA2Update()
        {
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = new DateTime(2025, 4, 15),
                    Hour = 16,
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
                    Date = new DateTime(2025, 4, 16),
                    Hour = 0,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "测试数据",
                            RowIndex = 5,
                            Values = new List<double?> { 12.0, 6.0, 16.0, 21.0, 26.0, 31.0 }
                        }
                    }
                },
                new ExcelFile
                {
                    Date = new DateTime(2025, 4, 16),
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = "测试数据",
                            RowIndex = 5,
                            Values = new List<double?> { 14.0, 7.0, 17.0, 22.0, 27.0, 32.0 }
                        }
                    }
                }
            };
        }

        #endregion
    }
}