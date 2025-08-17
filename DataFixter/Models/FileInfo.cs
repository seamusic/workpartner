using System;
using System.Text.RegularExpressions;

namespace DataFixter.Models
{
    /// <summary>
    /// 文件信息模型类
    /// 解析文件名中的日期、时间、项目名称，支持时间排序
    /// </summary>
    public class FileInfo : IComparable<FileInfo>, IEquatable<FileInfo>
    {
        /// <summary>
        /// 原始文件名
        /// </summary>
        public string OriginalFileName { get; }

        /// <summary>
        /// 文件完整路径
        /// </summary>
        public string FullPath { get; }

        /// <summary>
        /// 解析出的日期
        /// </summary>
        public DateTime Date { get; private set; }

        /// <summary>
        /// 解析出的时间（小时）
        /// </summary>
        public int Hour { get; private set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; private set; } = string.Empty;

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize { get; }

        /// <summary>
        /// 最后修改时间
        /// </summary>
        public DateTime LastModified { get; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath">文件完整路径</param>
        /// <param name="fileSize">文件大小</param>
        /// <param name="lastModified">最后修改时间</param>
        public FileInfo(string filePath, long fileSize, DateTime lastModified)
        {
            FullPath = filePath;
            FileSize = fileSize;
            LastModified = lastModified;
            OriginalFileName = System.IO.Path.GetFileName(filePath);
            
            // 解析文件名
            ParseFileName(OriginalFileName);
        }

        /// <summary>
        /// 解析文件名
        /// 格式: {日期}-{时}{项目名称}.xls
        /// 示例: 2025.7.1-00云港城项目4#地块.xls
        /// </summary>
        /// <param name="fileName">文件名</param>
        private void ParseFileName(string fileName)
        {
            try
            {
                // 移除文件扩展名
                string nameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(fileName);
                
                // 使用正则表达式解析文件名
                // 匹配格式: 2025.7.1-00云港城项目4#地块
                var regex = new Regex(@"^(\d{4}\.\d{1,2}\.\d{1,2})-(\d{2})(.+)$");
                var match = regex.Match(nameWithoutExt);
                
                if (match.Success)
                {
                    // 解析日期
                    string dateStr = match.Groups[1].Value;
                    Date = ParseDate(dateStr);
                    
                    // 解析小时
                    Hour = int.Parse(match.Groups[2].Value);
                    
                    // 解析项目名称
                    ProjectName = match.Groups[3].Value.Trim();
                }
                else
                {
                    // 如果解析失败，使用文件修改时间
                    Date = LastModified.Date;
                    Hour = LastModified.Hour;
                    ProjectName = "未知项目";
                }
            }
            catch (Exception)
            {
                // 解析失败时使用默认值
                Date = LastModified.Date;
                Hour = LastModified.Hour;
                ProjectName = "未知项目";
            }
        }

        /// <summary>
        /// 解析日期字符串
        /// 支持格式: 2025.7.1, 2025.07.01
        /// </summary>
        /// <param name="dateStr">日期字符串</param>
        /// <returns>解析后的日期</returns>
        private DateTime ParseDate(string dateStr)
        {
            try
            {
                // 将 2025.7.1 转换为 2025-07-01 格式
                var parts = dateStr.Split('.');
                if (parts.Length == 3)
                {
                    int year = int.Parse(parts[0]);
                    int month = int.Parse(parts[1]);
                    int day = int.Parse(parts[2]);
                    return new DateTime(year, month, day);
                }
                
                // 尝试直接解析
                return DateTime.Parse(dateStr);
            }
            catch
            {
                // 解析失败时返回当前日期
                return DateTime.Today;
            }
        }

        /// <summary>
        /// 获取完整的时间（日期+小时）
        /// </summary>
        public DateTime FullDateTime => Date.AddHours(Hour);

        /// <summary>
        /// 获取格式化的时间字符串
        /// </summary>
        public string FormattedDateTime => $"{Date:yyyy-MM-dd} {Hour:D2}:00";

        /// <summary>
        /// 获取格式化的文件名（用于显示）
        /// </summary>
        public string DisplayName => $"{FormattedDateTime} - {ProjectName}";

        /// <summary>
        /// 比较两个文件信息（按时间排序）
        /// </summary>
        public int CompareTo(FileInfo? other)
        {
            if (other == null) return 1;
            
            // 首先按日期比较
            int dateComparison = Date.CompareTo(other.Date);
            if (dateComparison != 0) return dateComparison;
            
            // 日期相同时按小时比较
            return Hour.CompareTo(other.Hour);
        }

        /// <summary>
        /// 比较两个文件信息是否相等
        /// </summary>
        public bool Equals(FileInfo? other)
        {
            if (other == null) return false;
            return FullPath.Equals(other.FullPath, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 重写Equals方法
        /// </summary>
        public override bool Equals(object? obj)
        {
            return Equals(obj as FileInfo);
        }

        /// <summary>
        /// 重写GetHashCode方法
        /// </summary>
        public override int GetHashCode()
        {
            return FullPath.GetHashCode();
        }

        /// <summary>
        /// 重写ToString方法
        /// </summary>
        public override string ToString()
        {
            return DisplayName;
        }

        /// <summary>
        /// 重载比较操作符
        /// </summary>
        public static bool operator <(FileInfo? left, FileInfo? right)
        {
            return left?.CompareTo(right) < 0;
        }

        public static bool operator <=(FileInfo? left, FileInfo? right)
        {
            return left?.CompareTo(right) <= 0;
        }

        public static bool operator >(FileInfo? left, FileInfo? right)
        {
            return left?.CompareTo(right) > 0;
        }

        public static bool operator >=(FileInfo? left, FileInfo? right)
        {
            return left?.CompareTo(right) >= 0;
        }
    }
}
