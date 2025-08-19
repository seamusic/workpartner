using System;
using System.IO;
using DataFixter.Models;
using DataFixter.Services;
using FluentAssertions;
using Microsoft.Extensions.Configuration;
using Serilog;
using Xunit;

namespace DataFixter.Tests
{
    /// <summary>
    /// 配置功能测试
    /// </summary>
    public class ConfigurationTests
    {
        private readonly ILogger _logger;

        public ConfigurationTests()
        {
            _logger = new LoggerConfiguration()
                .WriteTo.Console()
                .CreateLogger();
        }

        [Fact]
        public void CorrectionOptions_ShouldHaveCorrectDefaultValue()
        {
            // Arrange & Act
            var options = new CorrectionOptions();

            // Assert
            options.RandomChangeRange.Should().Be(0.3);
        }

        [Theory]
        [InlineData(0.1)]
        [InlineData(0.5)]
        [InlineData(1.0)]
        [InlineData(2.0)]
        public void CorrectionOptions_ShouldAcceptVariousRandomChangeRangeValues(double value)
        {
            // Arrange
            var options = new CorrectionOptions();

            // Act
            options.RandomChangeRange = value;

            // Assert
            options.RandomChangeRange.Should().Be(value);
        }

        [Fact]
        public void ConfigurationService_ShouldLoadRandomChangeRangeFromConfig()
        {
            // Arrange
            var configService = new ConfigurationService(_logger);

            // Act
            var options = configService.GetCorrectionOptions();

            // Assert
            // 由于ConfigurationService总是读取appsettings.json，我们验证它至少能正确读取默认值
            options.RandomChangeRange.Should().Be(0.3);
        }

        [Fact]
        public void ConfigurationService_ShouldLoadAllCorrectionOptions()
        {
            // Arrange
            var configService = new ConfigurationService(_logger);

            // Act
            var options = configService.GetCorrectionOptions();

            // Assert
            options.Should().NotBeNull();
            options.CumulativeTolerance.Should().Be(0.001); // 从appsettings.json读取
            options.MaxCurrentPeriodValue.Should().Be(1.0);
            options.MaxCumulativeValue.Should().Be(4.0);
            options.EnableMinimalModification.Should().BeTrue();
            options.RandomChangeRange.Should().Be(0.3);
        }
    }
}
