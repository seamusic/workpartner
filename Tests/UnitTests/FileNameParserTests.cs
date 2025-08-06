using Xunit;
using FluentAssertions;
using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Tests.UnitTests
{
    /// <summary>
    /// 文件名解析器单元测试
    /// </summary>
    public class FileNameParserTests
    {
        [Theory]
        [InlineData("2025.4.18-8云港城项目4#地块.xls", true, "2025/4/18", 8, "云港城项目4#地块.xls")]
        [InlineData("2025.4.20-16云港城项目4#地块.xlsx", true, "2025/4/20", 16, "云港城项目4#地块.xlsx")]
        [InlineData("2025.4.15-0测试项目.xls", true, "2025/4/15", 0, "测试项目.xls")]
        [InlineData("2025.12.31-23项目名称.xlsx", true, "2025/12/31", 23, "项目名称.xlsx")]
        public void ParseFileName_ValidFileNames_ShouldReturnCorrectResults(
            string fileName, bool expectedValid, string expectedDate, int expectedHour, string expectedProjectName)
        {
            // Act
            var result = FileNameParser.ParseFileName(fileName);

            // Assert
            result.Should().NotBeNull();
            result!.IsValid.Should().Be(expectedValid);
            
            if (expectedValid)
            {
                result.Date.Should().Be(DateTime.Parse(expectedDate));
                result.Hour.Should().Be(expectedHour);
                result.ProjectName.Should().Be(expectedProjectName);
                result.OriginalFileName.Should().Be(fileName);
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("invalid_file_name.xlsx")]
        [InlineData("2025-4-18-8云港城项目.xlsx")]  // 错误的日期格式
        [InlineData("2025.4.18_8云港城项目.xlsx")]   // 错误的分隔符
        [InlineData("2025.4.18-云港城项目.xlsx")]    // 缺少时间
        [InlineData("2025.13.18-8云港城项目.xlsx")]  // 无效的月份
        [InlineData("2025.4.32-8云港城项目.xlsx")]   // 无效的日期
        [InlineData("2025.4.18-25云港城项目.xlsx")]  // 无效的小时
        [InlineData("2025.4.18--8云港城项目.xlsx")]  // 双连字符
        public void ParseFileName_InvalidFileNames_ShouldReturnNull(string fileName)
        {
            // Act
            var result = FileNameParser.ParseFileName(fileName);

            // Assert
            result.Should().BeNull();
        }

        [Theory]
        [InlineData("2025.4.18-8云港城项目4#地块.xls", true)]
        [InlineData("2025.4.20-16云港城项目4#地块.xlsx", true)]
        [InlineData("invalid_file_name.xlsx", false)]
        [InlineData("", false)]
        [InlineData("   ", false)]
        public void IsValidFileName_ShouldReturnCorrectValidation(string fileName, bool expected)
        {
            // Act
            var result = FileNameParser.IsValidFileName(fileName);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("2025/4/18", 8, "云港城项目4#地块.xls", "2025.4.18-08云港城项目4#地块.xls")]
        [InlineData("2025/4/20", 16, "测试项目.xlsx", "2025.4.20-16测试项目.xlsx")]
        [InlineData("2025/12/31", 0, "项目名称.xls", "2025.12.31-00项目名称.xls")]
        public void GenerateFileName_ShouldCreateCorrectFileName(
            string dateStr, int hour, string projectName, string expected)
        {
            // Arrange
            var date = DateTime.Parse(dateStr);

            // Act
            var result = FileNameParser.GenerateFileName(date, hour, projectName);

            // Assert
            result.Should().Be(expected);
        }

        [Fact]
        public void GetExpectedHours_ShouldReturnCorrectHours()
        {
            // Act
            var result = FileNameParser.GetExpectedHours();

            // Assert
            result.Should().HaveCount(3);
            result.Should().Contain(new[] { 0, 8, 16 });
        }

        [Theory]
        [InlineData(0, true)]
        [InlineData(8, true)]
        [InlineData(16, true)]
        [InlineData(1, false)]
        [InlineData(12, false)]
        [InlineData(24, false)]
        [InlineData(-1, false)]
        public void IsValidHour_ShouldValidateHourCorrectly(int hour, bool expected)
        {
            // Act
            var result = FileNameParser.IsValidHour(hour);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData(new int[] { 0, 8 }, new int[] { 16 })]
        [InlineData(new int[] { 16 }, new int[] { 0, 8 })]
        [InlineData(new int[] { 0, 8, 16 }, new int[] { })]
        [InlineData(new int[] { }, new int[] { 0, 8, 16 })]
        public void GetMissingHours_ShouldReturnCorrectMissingHours(int[] existingHours, int[] expectedMissing)
        {
            // Act
            var result = FileNameParser.GetMissingHours(existingHours);

            // Assert
            result.Should().BeEquivalentTo(expectedMissing);
        }

        [Fact]
        public void ParseFileName_WithFullPath_ShouldReturnCorrectResult()
        {
            // Arrange
            var fullPath = @"C:\Data\Excel\2025.4.18-8云港城项目4#地块.xls";

            // Act
            var result = FileNameParser.ParseFileName(fullPath);

            // Assert
            result.Should().NotBeNull();
            result!.IsValid.Should().BeTrue();
            result.OriginalFileName.Should().Be("2025.4.18-8云港城项目4#地块.xls");
            result.FilePath.Should().Be(fullPath);
            result.Date.Should().Be(new DateTime(2025, 4, 18));
            result.Hour.Should().Be(8);
            result.ProjectName.Should().Be("云港城项目4#地块.xls");
        }

        [Fact]
        public void FileNameParseResult_FormattedProperties_ShouldReturnCorrectFormats()
        {
            // Arrange
            var result = new FileNameParseResult
            {
                IsValid = true,
                Date = new DateTime(2025, 4, 8),
                Hour = 8,
                ProjectName = "测试项目.xlsx",
                OriginalFileName = "2025.4.8-8测试项目.xlsx"
            };

            // Act & Assert
            result.FormattedDate.Should().Be("2025.4.8");
            result.FormattedHour.Should().Be("08");
            result.FileIdentifier.Should().Be("2025.4.8-08测试项目.xlsx");
        }

        [Fact]
        public void FileNameParseResult_ToString_ShouldReturnCorrectFormat()
        {
            // Arrange
            var validResult = new FileNameParseResult
            {
                IsValid = true,
                Date = new DateTime(2025, 4, 8),
                Hour = 8,
                ProjectName = "测试项目.xlsx",
                OriginalFileName = "2025.4.8-8测试项目.xlsx"
            };

            var invalidResult = new FileNameParseResult
            {
                IsValid = false,
                OriginalFileName = "invalid_file.xlsx"
            };

            // Act & Assert
            validResult.ToString().Should().Be("2025.4.8-08测试项目.xlsx");
            invalidResult.ToString().Should().Be("Invalid: invalid_file.xlsx");
        }
    }
}