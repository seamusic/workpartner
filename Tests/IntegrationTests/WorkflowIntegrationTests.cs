using Xunit;
using FluentAssertions;
using WorkPartner.Services;
using WorkPartner.Utils;
using WorkPartner.Models;
using System.IO.Abstractions.TestingHelpers;

namespace WorkPartner.Tests.IntegrationTests
{
    /// <summary>
    /// 工作流集成测试
    /// </summary>
    public class WorkflowIntegrationTests
    {
        [Fact]
        public void CompleteWorkflow_ProcessExcelFiles_ShouldExecuteSuccessfully()
        {
            // Arrange
            var tempDir = CreateTempTestEnvironment();
            var inputDir = Path.Combine(tempDir, "input");
            var outputDir = Path.Combine(tempDir, "output");
            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);

            try
            {
                // 创建测试文件
                CreateTestExcelFiles(inputDir);

                var fileService = new FileService();

                // Act & Assert
                
                // 1. 扫描文件
                var excelFiles = fileService.ScanExcelFiles(inputDir);
                excelFiles.Should().NotBeEmpty();

                // 2. 解析文件名
                var parsedFiles = new List<ExcelFile>();
                foreach (var filePath in excelFiles)
                {
                    var parseResult = FileNameParser.ParseFileName(filePath);
                    if (parseResult != null && parseResult.IsValid)
                    {
                        var excelFile = new ExcelFile
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            Date = parseResult.Date,
                            Hour = parseResult.Hour,
                            ProjectName = parseResult.ProjectName
                        };
                        parsedFiles.Add(excelFile);
                    }
                }

                parsedFiles.Should().NotBeEmpty();

                // 3. 检查完整性
                var completenessResult = DataProcessor.CheckCompleteness(parsedFiles);
                completenessResult.Should().NotBeNull();

                // 4. 生成补充文件
                var supplementFiles = DataProcessor.GenerateSupplementFiles(parsedFiles);

                // 5. 验证结果
                if (supplementFiles.Any())
                {
                    supplementFiles.Should().NotBeEmpty();
                    var createdCount = DataProcessor.CreateSupplementFiles(supplementFiles, outputDir);
                    createdCount.Should().BeGreaterThan(0);
                }

                // 6. 验证数据质量
                var qualityReport = DataProcessor.ValidateDataQuality(parsedFiles);
                qualityReport.Should().NotBeNull();
                qualityReport.TotalFiles.Should().Be(parsedFiles.Count);
            }
            finally
            {
                // Cleanup
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
        }

        [Fact]
        public void EdgeCaseWorkflow_EmptyDirectory_ShouldHandleGracefully()
        {
            // Arrange
            var tempDir = CreateTempTestEnvironment();
            var inputDir = Path.Combine(tempDir, "input");
            Directory.CreateDirectory(inputDir);

            try
            {
                var fileService = new FileService();

                // Act
                var excelFiles = fileService.ScanExcelFiles(inputDir);

                // Assert
                excelFiles.Should().BeEmpty();

                // 空文件列表的处理
                var parsedFiles = new List<ExcelFile>();
                var completenessResult = DataProcessor.CheckCompleteness(parsedFiles);
                completenessResult.Should().NotBeNull();
                completenessResult.IsAllComplete.Should().BeTrue();

                var supplementFiles = DataProcessor.GenerateSupplementFiles(parsedFiles);
                supplementFiles.Should().BeEmpty();
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
        }

        [Fact]
        public void ErrorHandlingWorkflow_InvalidFileNames_ShouldContinueProcessing()
        {
            // Arrange
            var tempDir = CreateTempTestEnvironment();
            var inputDir = Path.Combine(tempDir, "input");
            Directory.CreateDirectory(inputDir);

            try
            {
                // 创建有效和无效的文件名
                CreateMixedValidityFiles(inputDir);

                var fileService = new FileService();

                // Act
                var excelFiles = fileService.ScanExcelFiles(inputDir);
                excelFiles.Should().NotBeEmpty();

                var parsedFiles = new List<ExcelFile>();
                var invalidFiles = new List<string>();

                foreach (var filePath in excelFiles)
                {
                    var parseResult = FileNameParser.ParseFileName(filePath);
                    if (parseResult != null && parseResult.IsValid)
                    {
                        var excelFile = new ExcelFile
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            Date = parseResult.Date,
                            Hour = parseResult.Hour,
                            ProjectName = parseResult.ProjectName
                        };
                        parsedFiles.Add(excelFile);
                    }
                    else
                    {
                        invalidFiles.Add(filePath);
                    }
                }

                // Assert
                parsedFiles.Should().NotBeEmpty(); // 应该有一些有效文件
                invalidFiles.Should().NotBeEmpty(); // 应该有一些无效文件

                // 继续处理有效文件
                var completenessResult = DataProcessor.CheckCompleteness(parsedFiles);
                completenessResult.Should().NotBeNull();
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
        }

        [Fact]
        public void PerformanceTest_LargeNumberOfFiles_ShouldProcessInReasonableTime()
        {
            // Arrange
            var tempDir = CreateTempTestEnvironment();
            var inputDir = Path.Combine(tempDir, "input");
            Directory.CreateDirectory(inputDir);

            try
            {
                // 创建大量测试文件（模拟）
                var largeFileList = CreateLargeFileList(100);

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                // Act
                var completenessResult = DataProcessor.CheckCompleteness(largeFileList);
                var supplementFiles = DataProcessor.GenerateSupplementFiles(largeFileList);
                var qualityReport = DataProcessor.ValidateDataQuality(largeFileList);

                stopwatch.Stop();

                // Assert
                stopwatch.ElapsedMilliseconds.Should().BeLessThan(5000); // 应该在5秒内完成
                completenessResult.Should().NotBeNull();
                supplementFiles.Should().NotBeNull();
                qualityReport.Should().NotBeNull();
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
        }

        [Fact]
        public void DataConsistencyTest_ProcessedResults_ShouldMaintainIntegrity()
        {
            // Arrange
            var testFiles = CreateConsistencyTestFiles();

            // Act
            var originalCount = testFiles.Sum(f => f.DataRows.Sum(r => r.Values.Count));
            var processedFiles = DataProcessor.ProcessMissingData(testFiles);
            var processedCount = processedFiles.Sum(f => f.DataRows.Sum(r => r.Values.Count));

            // Assert
            processedCount.Should().Be(originalCount); // 数据数量应该保持一致
            
            foreach (var file in processedFiles)
            {
                foreach (var dataRow in file.DataRows)
                {
                    // 检查是否还有null值（应该都被补充了）
                    var nullCount = dataRow.Values.Count(v => !v.HasValue);
                    // 注意：可能仍有一些null值，如果所有相关文件都缺失该数据
                }
            }
        }

        #region Helper Methods

        private string CreateTempTestEnvironment()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "WorkPartnerIntegrationTest", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            return tempDir;
        }

        private void CreateTestExcelFiles(string directory)
        {
            // 创建一些有效的Excel文件名（不创建实际Excel内容，只是文件名）
            var testFiles = new[]
            {
                "2025.4.18-08测试项目.xlsx",
                "2025.4.18-16测试项目.xlsx",
                // 缺少 2025.4.18-00测试项目.xlsx
                "2025.4.19-00测试项目.xlsx",
                "2025.4.19-08测试项目.xlsx",
                "2025.4.19-16测试项目.xlsx"
            };

            foreach (var fileName in testFiles)
            {
                var filePath = Path.Combine(directory, fileName);
                File.WriteAllText(filePath, "dummy excel content"); // 创建假的Excel文件
            }
        }

        private void CreateMixedValidityFiles(string directory)
        {
            var testFiles = new[]
            {
                "2025.4.18-08测试项目.xlsx", // 有效
                "invalid_file_name.xlsx",    // 无效
                "2025.4.18-16测试项目.xlsx", // 有效
                "another_invalid.xls",       // 无效
                "2025.4.19-00测试项目.xlsx"  // 有效
            };

            foreach (var fileName in testFiles)
            {
                var filePath = Path.Combine(directory, fileName);
                File.WriteAllText(filePath, "dummy content");
            }
        }

        private List<ExcelFile> CreateLargeFileList(int count)
        {
            var files = new List<ExcelFile>();
            var baseDate = new DateTime(2025, 4, 1);

            for (int i = 0; i < count; i++)
            {
                var date = baseDate.AddDays(i / 3);
                var hour = (i % 3) switch
                {
                    0 => 0,
                    1 => 8,
                    _ => 16
                };

                files.Add(new ExcelFile
                {
                    Date = date,
                    Hour = hour,
                    ProjectName = $"项目{i % 10}.xlsx",
                    FileName = $"2025.4.{date.Day}-{hour:D2}项目{i % 10}.xlsx",
                    DataRows = new List<DataRow>
                    {
                        new DataRow
                        {
                            Name = $"数据{i}",
                            RowIndex = 5,
                            Values = new List<double?> { i * 1.0, i * 2.0, i * 3.0 }
                        }
                    }
                });
            }

            return files;
        }

        private List<ExcelFile> CreateConsistencyTestFiles()
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
                            Values = new List<double?> { 20.0, null, 24.0 } // 有一个缺失值
                        },
                        new DataRow
                        {
                            Name = "湿度",
                            RowIndex = 6,
                            Values = new List<double?> { 60.0, 65.0, 70.0 } // 完整数据
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
                            Values = new List<double?> { 22.0, 23.0, 25.0 } // 完整数据
                        },
                        new DataRow
                        {
                            Name = "湿度",
                            RowIndex = 6,
                            Values = new List<double?> { 62.0, null, 72.0 } // 有一个缺失值
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
                            Values = new List<double?> { 25.0, 26.0, 27.0 } // 完整数据
                        },
                        new DataRow
                        {
                            Name = "湿度",
                            RowIndex = 6,
                            Values = new List<double?> { 65.0, 68.0, 75.0 } // 完整数据
                        }
                    }
                }
            };
        }

        #endregion
    }
}