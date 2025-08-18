using System;
using System.Linq;
using System.Numerics;

namespace DataFixter.Utils
{
    /// <summary>
    /// 高精度数值计算工具类
    /// 使用decimal类型进行内部计算，避免浮点数精度误差
    /// </summary>
    public static class FloatingPointUtils
    {
        /// <summary>
        /// 默认的浮点数比较容差
        /// </summary>
        public const double DefaultTolerance = 1e-10;

        /// <summary>
        /// 工程计算中常用的容差
        /// </summary>
        public const double EngineeringTolerance = 1e-6;

        /// <summary>
        /// 高精度计算的容差
        /// </summary>
        public const double HighPrecisionTolerance = 1e-12;

        /// <summary>
        /// decimal类型的默认精度
        /// </summary>
        public const int DefaultDecimalPrecision = 28;

        /// <summary>
        /// 将double转换为decimal，保持精度
        /// </summary>
        /// <param name="value">double值</param>
        /// <returns>decimal值</returns>
        public static decimal ToDecimal(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                throw new ArgumentException("无法将NaN或Infinity转换为decimal", nameof(value));
            
            return Convert.ToDecimal(value);
        }

        /// <summary>
        /// 将decimal转换为double，用于输出
        /// </summary>
        /// <param name="value">decimal值</param>
        /// <returns>double值</returns>
        public static double ToDouble(decimal value)
        {
            return Convert.ToDouble(value);
        }

        /// <summary>
        /// 使用decimal进行高精度加法运算
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>高精度结果</returns>
        public static decimal SafeAdd(double a, double b)
        {
            try
            {
                var decimalA = ToDecimal(a);
                var decimalB = ToDecimal(b);
                return decimalA + decimalB;
            }
            catch (ArgumentException)
            {
                return decimal.Zero;
            }
        }

        /// <summary>
        /// 使用decimal进行高精度减法运算
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>高精度结果</returns>
        public static decimal SafeSubtract(double a, double b)
        {
            try
            {
                var decimalA = ToDecimal(a);
                var decimalB = ToDecimal(b);
                return decimalA - decimalB;
            }
            catch (ArgumentException)
            {
                return decimal.Zero;
            }
        }

        /// <summary>
        /// 使用decimal进行高精度乘法运算
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>高精度结果</returns>
        public static decimal SafeMultiply(double a, double b)
        {
            try
            {
                var decimalA = ToDecimal(a);
                var decimalB = ToDecimal(b);
                return decimalA * decimalB;
            }
            catch (ArgumentException)
            {
                return decimal.Zero;
            }
        }

        /// <summary>
        /// 使用decimal进行高精度除法运算
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>高精度结果</returns>
        public static decimal SafeDivide(double a, double b)
        {
            try
            {
                var decimalA = ToDecimal(a);
                var decimalB = ToDecimal(b);
                
                if (decimalB == 0)
                    throw new DivideByZeroException("除数不能为零");
                
                return decimalA / decimalB;
            }
            catch (ArgumentException)
            {
                return decimal.Zero;
            }
        }

        /// <summary>
        /// 高精度四舍五入，避免浮点数精度问题
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="digits">小数位数</param>
        /// <returns>四舍五入后的值</returns>
        public static double SafeRound(double value, int digits = 0)
        {
            try
            {
                var decimalValue = ToDecimal(value);
                var rounded = Math.Round(decimalValue, digits, MidpointRounding.AwayFromZero);
                return ToDouble(rounded);
            }
            catch (ArgumentException)
            {
                return value;
            }
        }

        /// <summary>
        /// 高精度截断小数位
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="digits">保留的小数位数</param>
        /// <returns>截断后的值</returns>
        public static double SafeTruncate(double value, int digits = 0)
        {
            try
            {
                var decimalValue = ToDecimal(value);
                var multiplier = (decimal)Math.Pow(10, digits);
                var scaled = decimal.Truncate(decimalValue * multiplier);
                var result = scaled / multiplier;
                return ToDouble(result);
            }
            catch (ArgumentException)
            {
                return value;
            }
        }

        /// <summary>
        /// 高精度计算平均值
        /// </summary>
        /// <param name="values">值列表</param>
        /// <returns>平均值</returns>
        public static double SafeAverage(params double[] values)
        {
            if (values == null || values.Length == 0)
                return double.NaN;

            try
            {
                var decimalValues = values.Select(v => ToDecimal(v)).ToArray();
                var sum = decimalValues.Aggregate(decimal.Zero, (acc, val) => acc + val);
                var average = sum / decimalValues.Length;
                return ToDouble(average);
            }
            catch (ArgumentException)
            {
                return values.Average();
            }
        }

        /// <summary>
        /// 高精度计算总和
        /// </summary>
        /// <param name="values">值列表</param>
        /// <returns>总和</returns>
        public static double SafeSum(params double[] values)
        {
            if (values == null || values.Length == 0)
                return 0.0;

            try
            {
                var decimalValues = values.Select(v => ToDecimal(v)).ToArray();
                var sum = decimalValues.Aggregate(decimal.Zero, (acc, val) => acc + val);
                return ToDouble(sum);
            }
            catch (ArgumentException)
            {
                return values.Sum();
            }
        }

        /// <summary>
        /// 高精度计算差值
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>差值</returns>
        public static double SafeDifference(double a, double b)
        {
            try
            {
                var result = SafeSubtract(a, b);
                return ToDouble(result);
            }
            catch
            {
                return a - b;
            }
        }

        /// <summary>
        /// 高精度计算绝对差值
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>绝对差值</returns>
        public static double SafeAbsoluteDifference(double a, double b)
        {
            var diff = SafeDifference(a, b);
            return Math.Abs(diff);
        }

        /// <summary>
        /// 安全地比较两个浮点数是否相等
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否相等</returns>
        public static bool AreEqual(double a, double b, double tolerance = DefaultTolerance)
        {
            // 处理特殊值
            if (double.IsNaN(a) || double.IsNaN(b))
                return false;
            
            if (double.IsInfinity(a) || double.IsInfinity(b))
                return a == b;

            // 使用相对误差和绝对误差的组合
            var absA = Math.Abs(a);
            var absB = Math.Abs(b);
            var diff = Math.Abs(a - b);

            // 如果两个值都很小，使用绝对误差
            if (absA < tolerance && absB < tolerance)
                return diff < tolerance;

            // 否则使用相对误差
            var relativeError = diff / Math.Max(absA, absB);
            return relativeError < tolerance;
        }

        /// <summary>
        /// 安全地比较两个浮点数是否不相等
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否不相等</returns>
        public static bool AreNotEqual(double a, double b, double tolerance = DefaultTolerance)
        {
            return !AreEqual(a, b, tolerance);
        }

        /// <summary>
        /// 安全地比较浮点数是否大于
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否大于</returns>
        public static bool IsGreaterThan(double a, double b, double tolerance = DefaultTolerance)
        {
            return a > b && AreNotEqual(a, b, tolerance);
        }

        /// <summary>
        /// 安全地比较浮点数是否小于
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否小于</returns>
        public static bool IsLessThan(double a, double b, double tolerance = DefaultTolerance)
        {
            return a < b && AreNotEqual(a, b, tolerance);
        }

        /// <summary>
        /// 安全地比较浮点数是否大于等于
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否大于等于</returns>
        public static bool IsGreaterThanOrEqual(double a, double b, double tolerance = DefaultTolerance)
        {
            return a > b || AreEqual(a, b, tolerance);
        }

        /// <summary>
        /// 安全地比较浮点数是否小于等于
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否小于等于</returns>
        public static bool IsLessThanOrEqual(double a, double b, double tolerance = DefaultTolerance)
        {
            return a < b || AreEqual(a, b, tolerance);
        }

        /// <summary>
        /// 安全地计算浮点数的绝对值
        /// </summary>
        /// <param name="value">值</param>
        /// <returns>绝对值</returns>
        public static double SafeAbs(double value)
        {
            if (double.IsNaN(value))
                return double.NaN;

            if (double.IsInfinity(value))
                return double.PositiveInfinity;

            return Math.Abs(value);
        }

        /// <summary>
        /// 安全地计算浮点数的平方根
        /// </summary>
        /// <param name="value">值</param>
        /// <returns>平方根</returns>
        public static double SafeSqrt(double value)
        {
            if (double.IsNaN(value) || value < 0)
                return double.NaN;

            if (double.IsInfinity(value))
                return double.PositiveInfinity;

            if (AreEqual(value, 0.0))
                return 0.0;

            return Math.Sqrt(value);
        }

        /// <summary>
        /// 安全地计算浮点数的幂
        /// </summary>
        /// <param name="baseValue">底数</param>
        /// <param name="exponent">指数</param>
        /// <returns>幂值</returns>
        public static double SafePow(double baseValue, double exponent)
        {
            if (double.IsNaN(baseValue) || double.IsNaN(exponent))
                return double.NaN;

            if (double.IsInfinity(baseValue) || double.IsInfinity(exponent))
            {
                // 处理无穷大的幂运算
                if (baseValue == 0 && exponent > 0)
                    return 0.0;
                if (baseValue == 0 && exponent < 0)
                    return double.PositiveInfinity;
                if (baseValue == 1)
                    return 1.0;
                if (exponent == 0)
                    return 1.0;
                return double.PositiveInfinity;
            }

            return Math.Pow(baseValue, exponent);
        }

        /// <summary>
        /// 安全地计算浮点数的对数
        /// </summary>
        /// <param name="value">值</param>
        /// <returns>自然对数</returns>
        public static double SafeLog(double value)
        {
            if (double.IsNaN(value) || value <= 0)
                return double.NaN;

            if (double.IsInfinity(value))
                return double.PositiveInfinity;

            return Math.Log(value);
        }

        /// <summary>
        /// 安全地计算浮点数的余弦值
        /// </summary>
        /// <param name="value">值（弧度）</param>
        /// <returns>余弦值</returns>
        public static double SafeCos(double value)
        {
            if (double.IsNaN(value))
                return double.NaN;

            if (double.IsInfinity(value))
                return double.NaN;

            return Math.Cos(value);
        }

        /// <summary>
        /// 安全地计算浮点数的符号
        /// </summary>
        /// <param name="value">值</param>
        /// <returns>符号值</returns>
        public static double SafeSign(double value)
        {
            if (double.IsNaN(value))
                return double.NaN;

            if (double.IsInfinity(value))
                return value > 0 ? 1.0 : -1.0;

            if (AreEqual(value, 0.0))
                return 0.0;

            return Math.Sign(value);
        }

        /// <summary>
        /// 安全地计算浮点数的最大值
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>最大值</returns>
        public static double SafeMax(double a, double b)
        {
            if (double.IsNaN(a) || double.IsNaN(b))
                return double.NaN;

            if (double.IsInfinity(a))
                return a > 0 ? a : b;

            if (double.IsInfinity(b))
                return b > 0 ? b : a;

            return Math.Max(a, b);
        }

        /// <summary>
        /// 安全地计算浮点数的最小值
        /// </summary>
        /// <param name="a">第一个值</param>
        /// <param name="b">第二个值</param>
        /// <returns>最小值</returns>
        public static double SafeMin(double a, double b)
        {
            if (double.IsNaN(a) || double.IsNaN(b))
                return double.NaN;

            if (double.IsInfinity(a))
                return a < 0 ? a : b;

            if (double.IsInfinity(b))
                return b < 0 ? b : a;

            return Math.Min(a, b);
        }

        /// <summary>
        /// 安全地计算浮点数的标准差
        /// </summary>
        /// <param name="values">值列表</param>
        /// <returns>标准差</returns>
        public static double SafeStandardDeviation(params double[] values)
        {
            if (values == null || values.Length <= 1)
                return 0.0;

            var validValues = values.Where(v => !double.IsNaN(v) && !double.IsInfinity(v)).ToArray();
            
            if (validValues.Length <= 1)
                return 0.0;

            var mean = validValues.Average();
            var sumSquaredDiff = validValues.Sum(v => SafePow(v - mean, 2));
            var variance = sumSquaredDiff / (validValues.Length - 1);
            
            return SafeSqrt(variance);
        }

        /// <summary>
        /// 检查浮点数是否为零（考虑容差）
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否为零</returns>
        public static bool IsZero(double value, double tolerance = DefaultTolerance)
        {
            return AreEqual(value, 0.0, tolerance);
        }

        /// <summary>
        /// 检查浮点数是否为正数（考虑容差）
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否为正数</returns>
        public static bool IsPositive(double value, double tolerance = DefaultTolerance)
        {
            return IsGreaterThan(value, 0.0, tolerance);
        }

        /// <summary>
        /// 检查浮点数是否为负数（考虑容差）
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="tolerance">容差</param>
        /// <returns>是否为负数</returns>
        public static bool IsNegative(double value, double tolerance = DefaultTolerance)
        {
            return IsLessThan(value, 0.0, tolerance);
        }

        /// <summary>
        /// 安全地限制浮点数在指定范围内
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="min">最小值</param>
        /// <param name="max">最大值</param>
        /// <returns>限制后的值</returns>
        public static double SafeClamp(double value, double min, double max)
        {
            if (double.IsNaN(value))
                return double.NaN;

            if (double.IsInfinity(value))
            {
                if (value > 0) return max;
                return min;
            }

            return SafeMax(min, SafeMin(value, max));
        }

        /// <summary>
        /// 格式化数值，避免科学计数法
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="digits">小数位数</param>
        /// <returns>格式化后的字符串</returns>
        public static string FormatNumber(double value, int digits = 6)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return value.ToString();

            try
            {
                var decimalValue = ToDecimal(value);
                var rounded = Math.Round(decimalValue, digits, MidpointRounding.AwayFromZero);
                return rounded.ToString($"F{digits}");
            }
            catch
            {
                return value.ToString($"F{digits}");
            }
        }

        /// <summary>
        /// 检查数值是否在合理范围内
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="min">最小值</param>
        /// <param name="max">最大值</param>
        /// <returns>是否在范围内</returns>
        public static bool IsInRange(double value, double min, double max)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return false;

            return value >= min && value <= max;
        }
    }
}
