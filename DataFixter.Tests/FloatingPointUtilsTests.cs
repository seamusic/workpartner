using System;
using Xunit;
using DataFixter.Utils;

namespace DataFixter.Tests
{
    /// <summary>
    /// 浮点数工具类测试
    /// </summary>
    public class FloatingPointUtilsTests
    {
        [Fact]
        public void AreEqual_WithSameValues_ReturnsTrue()
        {
            // Arrange
            var a = 1.0;
            var b = 1.0;

            // Act
            var result = FloatingPointUtils.AreEqual(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void AreEqual_WithDifferentValues_ReturnsFalse()
        {
            // Arrange
            var a = 1.0;
            var b = 2.0;

            // Act
            var result = FloatingPointUtils.AreEqual(a, b);

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void AreEqual_WithVeryCloseValues_ReturnsTrue()
        {
            // Arrange
            var a = 1.0;
            var b = 1.0 + 1e-11; // 非常接近的值

            // Act
            var result = FloatingPointUtils.AreEqual(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void AreEqual_WithNaN_ReturnsFalse()
        {
            // Arrange
            var a = double.NaN;
            var b = 1.0;

            // Act
            var result = FloatingPointUtils.AreEqual(a, b);

            // Assert
            Assert.False(result);
        }

        [Fact]
        public void AreEqual_WithInfinity_ReturnsTrue()
        {
            // Arrange
            var a = double.PositiveInfinity;
            var b = double.PositiveInfinity;

            // Act
            var result = FloatingPointUtils.AreEqual(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void SafeAbs_WithPositiveValue_ReturnsSameValue()
        {
            // Arrange
            var value = 5.5;

            // Act
            var result = FloatingPointUtils.SafeAbs(value);

            // Assert
            Assert.Equal(5.5, result);
        }

        [Fact]
        public void SafeAbs_WithNegativeValue_ReturnsAbsoluteValue()
        {
            // Arrange
            var value = -5.5;

            // Act
            var result = FloatingPointUtils.SafeAbs(value);

            // Assert
            Assert.Equal(5.5, result);
        }

        [Fact]
        public void SafeAbs_WithNaN_ReturnsNaN()
        {
            // Arrange
            var value = double.NaN;

            // Act
            var result = FloatingPointUtils.SafeAbs(value);

            // Assert
            Assert.True(double.IsNaN(result));
        }

        [Fact]
        public void SafeSqrt_WithValidValue_ReturnsCorrectResult()
        {
            // Arrange
            var value = 4.0;

            // Act
            var result = FloatingPointUtils.SafeSqrt(value);

            // Assert
            Assert.Equal(2.0, result);
        }

        [Fact]
        public void SafeSqrt_WithNegativeValue_ReturnsNaN()
        {
            // Arrange
            var value = -4.0;

            // Act
            var result = FloatingPointUtils.SafeSqrt(value);

            // Assert
            Assert.True(double.IsNaN(result));
        }

        [Fact]
        public void SafePow_WithValidValues_ReturnsCorrectResult()
        {
            // Arrange
            var baseValue = 2.0;
            var exponent = 3.0;

            // Act
            var result = FloatingPointUtils.SafePow(baseValue, exponent);

            // Assert
            Assert.Equal(8.0, result);
        }

        [Fact]
        public void SafePow_WithZeroExponent_ReturnsOne()
        {
            // Arrange
            var baseValue = 5.0;
            var exponent = 0.0;

            // Act
            var result = FloatingPointUtils.SafePow(baseValue, exponent);

            // Assert
            Assert.Equal(1.0, result);
        }

        [Fact]
        public void SafeLog_WithValidValue_ReturnsCorrectResult()
        {
            // Arrange
            var value = Math.E; // e^1 = e

            // Act
            var result = FloatingPointUtils.SafeLog(value);

            // Assert
            Assert.Equal(1.0, result, 10);
        }

        [Fact]
        public void SafeLog_WithNegativeValue_ReturnsNaN()
        {
            // Arrange
            var value = -1.0;

            // Act
            var result = FloatingPointUtils.SafeLog(value);

            // Assert
            Assert.True(double.IsNaN(result));
        }

        [Fact]
        public void SafeCos_WithValidValue_ReturnsCorrectResult()
        {
            // Arrange
            var value = 0.0; // cos(0) = 1

            // Act
            var result = FloatingPointUtils.SafeCos(value);

            // Assert
            Assert.Equal(1.0, result);
        }

        [Fact]
        public void SafeSign_WithPositiveValue_ReturnsOne()
        {
            // Arrange
            var value = 5.5;

            // Act
            var result = FloatingPointUtils.SafeSign(value);

            // Assert
            Assert.Equal(1.0, result);
        }

        [Fact]
        public void SafeSign_WithNegativeValue_ReturnsNegativeOne()
        {
            // Arrange
            var value = -5.5;

            // Act
            var result = FloatingPointUtils.SafeSign(value);

            // Assert
            Assert.Equal(-1.0, result);
        }

        [Fact]
        public void SafeSign_WithZero_ReturnsZero()
        {
            // Arrange
            var value = 0.0;

            // Act
            var result = FloatingPointUtils.SafeSign(value);

            // Assert
            Assert.Equal(0.0, result);
        }

        [Fact]
        public void SafeMax_WithValidValues_ReturnsMaximum()
        {
            // Arrange
            var a = 5.5;
            var b = 3.2;

            // Act
            var result = FloatingPointUtils.SafeMax(a, b);

            // Assert
            Assert.Equal(5.5, result);
        }

        [Fact]
        public void SafeMin_WithValidValues_ReturnsMinimum()
        {
            // Arrange
            var a = 5.5;
            var b = 3.2;

            // Act
            var result = FloatingPointUtils.SafeMin(a, b);

            // Assert
            Assert.Equal(3.2, result);
        }

        [Fact]
        public void SafeRound_WithValidValue_ReturnsRoundedValue()
        {
            // Arrange
            var value = 3.14159;

            // Act
            var result = FloatingPointUtils.SafeRound(value, 2);

            // Assert
            Assert.Equal(3.14, result);
        }

        [Fact]
        public void SafeAverage_WithValidValues_ReturnsCorrectAverage()
        {
            // Arrange
            var values = new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 };

            // Act
            var result = FloatingPointUtils.SafeAverage(values);

            // Assert
            Assert.Equal(3.0, result);
        }

        [Fact]
        public void SafeStandardDeviation_WithValidValues_ReturnsCorrectResult()
        {
            // Arrange
            var values = new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 };

            // Act
            var result = FloatingPointUtils.SafeStandardDeviation(values);

            // Assert
            Assert.True(result > 0);
            Assert.True(result < 2.0);
        }

        [Fact]
        public void IsZero_WithZeroValue_ReturnsTrue()
        {
            // Arrange
            var value = 0.0;

            // Act
            var result = FloatingPointUtils.IsZero(value);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsZero_WithVerySmallValue_ReturnsTrue()
        {
            // Arrange
            var value = 1e-12; // 非常小的值

            // Act
            var result = FloatingPointUtils.IsZero(value);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsPositive_WithPositiveValue_ReturnsTrue()
        {
            // Arrange
            var value = 5.5;

            // Act
            var result = FloatingPointUtils.IsPositive(value);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsNegative_WithNegativeValue_ReturnsTrue()
        {
            // Arrange
            var value = -5.5;

            // Act
            var result = FloatingPointUtils.IsNegative(value);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void SafeClamp_WithValueInRange_ReturnsSameValue()
        {
            // Arrange
            var value = 5.0;
            var min = 0.0;
            var max = 10.0;

            // Act
            var result = FloatingPointUtils.SafeClamp(value, min, max);

            // Assert
            Assert.Equal(5.0, result);
        }

        [Fact]
        public void SafeClamp_WithValueAboveMax_ReturnsMax()
        {
            // Arrange
            var value = 15.0;
            var min = 0.0;
            var max = 10.0;

            // Act
            var result = FloatingPointUtils.SafeClamp(value, min, max);

            // Assert
            Assert.Equal(10.0, result);
        }

        [Fact]
        public void SafeClamp_WithValueBelowMin_ReturnsMin()
        {
            // Arrange
            var value = -5.0;
            var min = 0.0;
            var max = 10.0;

            // Act
            var result = FloatingPointUtils.SafeClamp(value, min, max);

            // Assert
            Assert.Equal(0.0, result);
        }

        [Fact]
        public void SafeAbsoluteDifference_WithValidValues_ReturnsCorrectResult()
        {
            // Arrange
            var a = 5.0;
            var b = 3.0;

            // Act
            var result = FloatingPointUtils.SafeAbsoluteDifference(a, b);

            // Assert
            Assert.Equal(2.0, result);
        }

        [Fact]
        public void SafeAbsoluteDifference_WithReversedOrder_ReturnsSameResult()
        {
            // Arrange
            var a = 3.0;
            var b = 5.0;

            // Act
            var result = FloatingPointUtils.SafeAbsoluteDifference(a, b);

            // Assert
            Assert.Equal(2.0, result);
        }

        [Fact]
        public void IsGreaterThan_WithValidValues_ReturnsCorrectResult()
        {
            // Arrange
            var a = 5.0;
            var b = 3.0;

            // Act
            var result = FloatingPointUtils.IsGreaterThan(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsLessThan_WithValidValues_ReturnsCorrectResult()
        {
            // Arrange
            var a = 3.0;
            var b = 5.0;

            // Act
            var result = FloatingPointUtils.IsLessThan(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsGreaterThanOrEqual_WithEqualValues_ReturnsTrue()
        {
            // Arrange
            var a = 5.0;
            var b = 5.0;

            // Act
            var result = FloatingPointUtils.IsGreaterThanOrEqual(a, b);

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void IsLessThanOrEqual_WithEqualValues_ReturnsTrue()
        {
            // Arrange
            var a = 5.0;
            var b = 5.0;

            // Act
            var result = FloatingPointUtils.IsLessThanOrEqual(a, b);

            // Assert
            Assert.True(result);
        }
    }
}
