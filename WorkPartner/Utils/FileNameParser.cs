using System.Text.RegularExpressions;
using WorkPartner.Models;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 文件名解析工具类
    /// </summary>
    public static class FileNameParser
    {
        /// <summary>
        /// 文件名正则表达式
        /// 格式：日期-时项目名称.xls
        /// 示例：2025.4.20-16云港城项目4#地块.xls
        /// </summary>
        private static readonly Regex FileNameRegex = new(
            @"^(\d{4}\.\d{1,2}\.\d{1,2})-(\d{1,2})(.+)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// 解析文件名
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>解析结果</returns>
        public static FileNameParseResult? ParseFileName(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }

            var fileName = Path.GetFileName(filePath);
            var match = FileNameRegex.Match(fileName);

            if (!match.Success)
            {
                return null;
            }

            try
            {
                var dateStr = match.Groups[1].Value;
                var hourStr = match.Groups[2].Value;
                var projectName = match.Groups[3].Value;

                // 解析日期
                if (!DateTime.TryParse(dateStr.Replace('.', '/'), out var date))
                {
                    return null;
                }

                // 解析时间
                if (!int.TryParse(hourStr, out var hour) || hour < 0 || hour > 23)
                {
                    return null;
                }

                return new FileNameParseResult
                {
                    IsValid = true,
                    Date = date,
                    Hour = hour,
                    ProjectName = projectName,
                    OriginalFileName = fileName,
                    FilePath = filePath
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 验证文件名格式
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <returns>是否格式正确</returns>
        public static bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            return FileNameRegex.IsMatch(fileName);
        }

        /// <summary>
        /// 生成文件名
        /// </summary>
        /// <param name="date">日期</param>
        /// <param name="hour">时间</param>
        /// <param name="projectName">项目名称</param>
        /// <returns>生成的文件名</returns>
        public static string GenerateFileName(DateTime date, int hour, string projectName)
        {
            var dateStr = date.ToString("yyyy.M.d");
            var hourStr = hour.ToString("D2");
            return $"{dateStr}-{hourStr}{projectName}";
        }

        /// <summary>
        /// 获取期望的时间点列表
        /// </summary>
        /// <returns>时间点列表</returns>
        public static List<int> GetExpectedHours()
        {
            return new List<int> { 0, 8, 16 };
        }

        /// <summary>
        /// 检查时间点是否有效
        /// </summary>
        /// <param name="hour">时间点</param>
        /// <returns>是否有效</returns>
        public static bool IsValidHour(int hour)
        {
            return GetExpectedHours().Contains(hour);
        }

        /// <summary>
        /// 获取缺失的时间点
        /// </summary>
        /// <param name="existingHours">现有的时间点</param>
        /// <returns>缺失的时间点</returns>
        public static List<int> GetMissingHours(IEnumerable<int> existingHours)
        {
            var expectedHours = GetExpectedHours();
            return expectedHours.Except(existingHours).ToList();
        }
    }

    /// <summary>
    /// 文件名解析结果
    /// </summary>
    public class FileNameParseResult
    {
        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 时间点
        /// </summary>
        public int Hour { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 原始文件名
        /// </summary>
        public string OriginalFileName { get; set; } = string.Empty;

        /// <summary>
        /// 文件路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 获取格式化的日期字符串
        /// </summary>
        public string FormattedDate => Date.ToString("yyyy.M.d");

        /// <summary>
        /// 获取格式化的时间字符串
        /// </summary>
        public string FormattedHour => Hour.ToString("D2");

        /// <summary>
        /// 获取文件标识符
        /// </summary>
        public string FileIdentifier => $"{FormattedDate}-{FormattedHour}{ProjectName}";

        public override string ToString()
        {
            return IsValid ? FileIdentifier : $"Invalid: {OriginalFileName}";
        }
    }
} 