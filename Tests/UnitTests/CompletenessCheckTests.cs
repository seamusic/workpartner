using Xunit;
using FluentAssertions;
using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Tests.UnitTests
{
    /// <summary>
    /// 完整性检查单元测试
    /// </summary>
    public class CompletenessCheckTests
    {
        [Fact]
        public void CheckCompleteness_SingleCompleteDay_ShouldReturnComplete()
        {
            // Arrange
            var files = CreateCompleteDay();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IsAllComplete.Should().BeTrue();
            result.IncompleteDates.Should().BeEmpty();
            result.DateCompleteness.Should().HaveCount(1);
            
            var dayCompleteness = result.DateCompleteness[0];
            dayCompleteness.IsComplete.Should().BeTrue();
            dayCompleteness.ExistingHours.Should().BeEquivalentTo(new[] { 0, 8, 16 });
            dayCompleteness.MissingHours.Should().BeEmpty();
        }

        [Fact]
        public void CheckCompleteness_MultipleDaysWithOneMissing_ShouldIdentifyIncompleteDay()
        {
            // Arrange
            var files = CreateMultipleDaysWithMissing();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IsAllComplete.Should().BeFalse();
            result.IncompleteDates.Should().HaveCount(1);
            result.DateCompleteness.Should().HaveCount(2);

            var incompleteDay = result.DateCompleteness.FirstOrDefault(d => !d.IsComplete);
            incompleteDay.Should().NotBeNull();
            incompleteDay!.MissingHours.Should().Contain(0); // 缺少0点
            incompleteDay.ExistingHours.Should().BeEquivalentTo(new[] { 8, 16 });
        }

        [Fact]
        public void CheckCompleteness_OnlyOneHourPerDay_ShouldIdentifyMissingHours()
        {
            // Arrange
            var files = CreateSingleHourPerDay();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IsAllComplete.Should().BeFalse();
            result.IncompleteDates.Should().HaveCount(3);

            foreach (var dayCompleteness in result.DateCompleteness)
            {
                dayCompleteness.IsComplete.Should().BeFalse();
                dayCompleteness.ExistingHours.Should().HaveCount(1);
                dayCompleteness.MissingHours.Should().HaveCount(2);
            }
        }

        [Fact]
        public void CheckCompleteness_MixedScenario_ShouldCorrectlyClassifyDays()
        {
            // Arrange
            var files = CreateMixedScenario();

            // Act
            var result = DataProcessor.CheckCompleteness(files);

            // Assert
            result.Should().NotBeNull();
            result.IsAllComplete.Should().BeFalse();
            result.DateCompleteness.Should().HaveCount(3);

            // 第一天完整
            var day1 = result.DateCompleteness.FirstOrDefault(d => d.Date.Day == 18);
            day1.Should().NotBeNull();
            day1!.IsComplete.Should().BeTrue();

            // 第二天不完整
            var day2 = result.DateCompleteness.FirstOrDefault(d => d.Date.Day == 19);
            day2.Should().NotBeNull();
            day2!.IsComplete.Should().BeFalse();
            day2.MissingHours.Should().Contain(0);

            // 第三天不完整
            var day3 = result.DateCompleteness.FirstOrDefault(d => d.Date.Day == 20);
            day3.Should().NotBeNull();
            day3!.IsComplete.Should().BeFalse();
            day3.MissingHours.Should().BeEquivalentTo(new[] { 8, 16 });
        }

        [Fact]
        public void GenerateSupplementFiles_CompleteDay_ShouldReturnEmpty()
        {
            // Arrange
            var files = CreateCompleteDay();

            // Act
            var result = DataProcessor.GenerateSupplementFiles(files);

            // Assert
            result.Should().BeEmpty();
        }

        [Fact]
        public void GenerateSupplementFiles_MissingOneHour_ShouldGenerateOneFile()
        {
            // Arrange
            var files = CreateDayMissingOneHour();

            // Act
            var result = DataProcessor.GenerateSupplementFiles(files);

            // Assert
            result.Should().HaveCount(1);
            
            var supplementFile = result[0];
            supplementFile.TargetHour.Should().Be(0); // 缺少0点
            supplementFile.TargetDate.Should().Be(new DateTime(2025, 4, 18));
            supplementFile.SourceFile.Should().NotBeNull();
            supplementFile.TargetFileName.Should().Contain("2025.4.18-00");
        }

        [Fact]
        public void GenerateSupplementFiles_MissingMultipleHours_ShouldGenerateMultipleFiles()
        {
            // Arrange
            var files = CreateDayMissingMultipleHours();

            // Act
            var result = DataProcessor.GenerateSupplementFiles(files);

            // Assert
            result.Should().HaveCount(2);
            
            var missingHours = result.Select(s => s.TargetHour).ToList();
            missingHours.Should().BeEquivalentTo(new[] { 0, 8 });
        }

        [Fact]
        public void CreateSupplementFiles_ValidInput_ShouldProcessSuccessfully()
        {
            // Arrange
            var files = CreateDayMissingOneHour();
            var supplementFiles = DataProcessor.GenerateSupplementFiles(files);
            var tempOutputDir = Path.Combine(Path.GetTempPath(), "WorkPartnerTest", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempOutputDir);

            try
            {
                // Act
                var result = DataProcessor.CreateSupplementFiles(supplementFiles, tempOutputDir);

                // Assert
                result.Should().BeGreaterThan(0);
                
                // 验证文件是否创建成功
                var createdFiles = Directory.GetFiles(tempOutputDir, "*.xls*");
                createdFiles.Should().NotBeEmpty();
            }
            finally
            {
                // Cleanup
                if (Directory.Exists(tempOutputDir))
                {
                    Directory.Delete(tempOutputDir, true);
                }
            }
        }

        #region Helper Methods

        private List<ExcelFile> CreateCompleteDay()
        {
            var date = new DateTime(2025, 4, 18);
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = date,
                    Hour = 0,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-00测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-00测试项目.xlsx"
                },
                new ExcelFile
                {
                    Date = date,
                    Hour = 8,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-08测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-08测试项目.xlsx"
                },
                new ExcelFile
                {
                    Date = date,
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-16测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-16测试项目.xlsx"
                }
            };
        }

        private List<ExcelFile> CreateMultipleDaysWithMissing()
        {
            var day1 = new DateTime(2025, 4, 18);
            var day2 = new DateTime(2025, 4, 19);
            
            return new List<ExcelFile>
            {
                // 第一天完整
                new ExcelFile { Date = day1, Hour = 0, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = day1, Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = day1, Hour = 16, ProjectName = "测试项目.xlsx" },
                
                // 第二天缺少0点
                new ExcelFile { Date = day2, Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = day2, Hour = 16, ProjectName = "测试项目.xlsx" }
            };
        }

        private List<ExcelFile> CreateSingleHourPerDay()
        {
            return new List<ExcelFile>
            {
                new ExcelFile { Date = new DateTime(2025, 4, 18), Hour = 0, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = new DateTime(2025, 4, 19), Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = new DateTime(2025, 4, 20), Hour = 16, ProjectName = "测试项目.xlsx" }
            };
        }

        private List<ExcelFile> CreateMixedScenario()
        {
            return new List<ExcelFile>
            {
                // 第一天完整
                new ExcelFile { Date = new DateTime(2025, 4, 18), Hour = 0, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = new DateTime(2025, 4, 18), Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = new DateTime(2025, 4, 18), Hour = 16, ProjectName = "测试项目.xlsx" },
                
                // 第二天缺少0点
                new ExcelFile { Date = new DateTime(2025, 4, 19), Hour = 8, ProjectName = "测试项目.xlsx" },
                new ExcelFile { Date = new DateTime(2025, 4, 19), Hour = 16, ProjectName = "测试项目.xlsx" },
                
                // 第三天只有0点
                new ExcelFile { Date = new DateTime(2025, 4, 20), Hour = 0, ProjectName = "测试项目.xlsx" }
            };
        }

        private List<ExcelFile> CreateDayMissingOneHour()
        {
            var date = new DateTime(2025, 4, 18);
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = date,
                    Hour = 8,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-08测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-08测试项目.xlsx"
                },
                new ExcelFile
                {
                    Date = date,
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-16测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-16测试项目.xlsx"
                }
                // 缺少0点文件
            };
        }

        private List<ExcelFile> CreateDayMissingMultipleHours()
        {
            var date = new DateTime(2025, 4, 18);
            return new List<ExcelFile>
            {
                new ExcelFile
                {
                    Date = date,
                    Hour = 16,
                    ProjectName = "测试项目.xlsx",
                    FileName = "2025.4.18-16测试项目.xlsx",
                    FilePath = @"C:\Test\2025.4.18-16测试项目.xlsx"
                }
                // 缺少0点和8点文件
            };
        }

        #endregion
    }
}