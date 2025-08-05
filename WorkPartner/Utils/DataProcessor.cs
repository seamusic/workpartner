using WorkPartner.Models;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 数据处理工具类
    /// </summary>
    public static class DataProcessor
    {
        /// <summary>
        /// 处理缺失数据
        /// </summary>
        /// <param name="files">Excel文件列表</param>
        /// <returns>处理后的文件列表</returns>
        public static List<ExcelFile> ProcessMissingData(List<ExcelFile> files)
        {
            if (files == null || !files.Any())
            {
                return new List<ExcelFile>();
            }

            // 按时间顺序排序
            var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

            // 处理每个文件中的缺失数据
            for (int i = 0; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                ProcessFileMissingData(currentFile, sortedFiles, i);
            }

            return sortedFiles;
        }

        /// <summary>
        /// 处理单个文件的缺失数据
        /// </summary>
        /// <param name="currentFile">当前文件</param>
        /// <param name="allFiles">所有文件列表</param>
        /// <param name="currentIndex">当前文件索引</param>
        private static void ProcessFileMissingData(ExcelFile currentFile, List<ExcelFile> allFiles, int currentIndex)
        {
            foreach (var dataRow in currentFile.DataRows)
            {
                for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                {
                    if (!dataRow.Values[valueIndex].HasValue)
                    {
                        // 计算补充值
                        var supplementValue = CalculateSupplementValue(dataRow.Name, valueIndex, allFiles, currentIndex);
                        if (supplementValue.HasValue)
                        {
                            dataRow.Values[valueIndex] = supplementValue.Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 计算补充值
        /// </summary>
        /// <param name="dataName">数据名称</param>
        /// <param name="valueIndex">值索引</param>
        /// <param name="allFiles">所有文件列表</param>
        /// <param name="currentIndex">当前文件索引</param>
        /// <returns>补充值</returns>
        private static double? CalculateSupplementValue(string dataName, int valueIndex, List<ExcelFile> allFiles, int currentIndex)
        {
            var validValues = new List<double>();

            // 查找前后文件中相同名称的有效数据
            for (int i = 0; i < allFiles.Count; i++)
            {
                if (i == currentIndex) continue; // 跳过当前文件

                var file = allFiles[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataName);
                
                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    validValues.Add(dataRow.Values[valueIndex].Value);
                }
            }

            // 计算平均值
            return validValues.Any() ? validValues.Average() : null;
        }

        /// <summary>
        /// 检查数据完整性
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>完整性检查结果</returns>
        public static CompletenessCheckResult CheckCompleteness(List<ExcelFile> files)
        {
            var result = new CompletenessCheckResult();

            if (files == null || !files.Any())
            {
                return result;
            }

            // 按日期分组
            var fileGroups = files.GroupBy(f => f.Date.Date).ToList();

            foreach (var group in fileGroups)
            {
                var date = group.Key;
                var existingHours = group.Select(f => f.Hour).ToList();
                var missingHours = FileNameParser.GetMissingHours(existingHours);

                var dateCompleteness = new DateCompleteness
                {
                    Date = date,
                    ExistingHours = existingHours,
                    MissingHours = missingHours,
                    IsComplete = !missingHours.Any()
                };

                result.DateCompleteness.Add(dateCompleteness);

                if (!dateCompleteness.IsComplete)
                {
                    result.IncompleteDates.Add(date);
                }
            }

            result.IsAllComplete = !result.IncompleteDates.Any();
            return result;
        }

        /// <summary>
        /// 生成补充文件列表
        /// </summary>
        /// <param name="files">现有文件列表</param>
        /// <returns>需要补充的文件列表</returns>
        public static List<SupplementFileInfo> GenerateSupplementFiles(List<ExcelFile> files)
        {
            var supplementFiles = new List<SupplementFileInfo>();

            if (files == null || !files.Any())
            {
                return supplementFiles;
            }

            var completenessResult = CheckCompleteness(files);

            foreach (var dateCompleteness in completenessResult.DateCompleteness)
            {
                if (dateCompleteness.IsComplete) continue;

                var dateFiles = files.Where(f => f.Date.Date == dateCompleteness.Date).ToList();
                var sourceFile = dateFiles.FirstOrDefault(); // 使用当天第一个文件作为源文件

                if (sourceFile is null) continue;

                foreach (var missingHour in dateCompleteness.MissingHours)
                {
                    var supplementFile = new SupplementFileInfo
                    {
                        TargetDate = dateCompleteness.Date,
                        TargetHour = missingHour,
                        ProjectName = sourceFile.ProjectName,
                        SourceFile = sourceFile,
                        TargetFileName = FileNameParser.GenerateFileName(dateCompleteness.Date, missingHour, sourceFile.ProjectName)
                    };

                    supplementFiles.Add(supplementFile);
                }
            }

            return supplementFiles;
        }

        /// <summary>
        /// 验证数据质量
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>数据质量报告</returns>
        public static DataQualityReport ValidateDataQuality(List<ExcelFile> files)
        {
            var report = new DataQualityReport();

            if (files == null || !files.Any())
            {
                return report;
            }

            foreach (var file in files)
            {
                var fileQuality = new FileQualityInfo
                {
                    FileName = file.FileName,
                    TotalRows = file.DataRows.Count,
                    ValidRows = file.DataRows.Count(r => r.IsAllValid),
                    MissingRows = file.DataRows.Count(r => r.HasMissingData),
                    AllMissingRows = file.DataRows.Count(r => r.IsAllMissing),
                    AverageCompleteness = file.DataRows.Any() ? file.DataRows.Average(r => r.CompletenessPercentage) : 0
                };

                report.FileQuality.Add(fileQuality);
                report.TotalFiles++;
                report.TotalRows += fileQuality.TotalRows;
                report.ValidRows += fileQuality.ValidRows;
                report.MissingRows += fileQuality.MissingRows;
            }

            report.OverallCompleteness = report.TotalRows > 0 ? (double)report.ValidRows / report.TotalRows * 100 : 0;
            return report;
        }
    }

    /// <summary>
    /// 完整性检查结果
    /// </summary>
    public class CompletenessCheckResult
    {
        /// <summary>
        /// 是否所有日期都完整
        /// </summary>
        public bool IsAllComplete { get; set; }

        /// <summary>
        /// 不完整的日期列表
        /// </summary>
        public List<DateTime> IncompleteDates { get; set; } = new List<DateTime>();

        /// <summary>
        /// 每个日期的完整性信息
        /// </summary>
        public List<DateCompleteness> DateCompleteness { get; set; } = new List<DateCompleteness>();
    }

    /// <summary>
    /// 日期完整性信息
    /// </summary>
    public class DateCompleteness
    {
        /// <summary>
        /// 日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 现有的时间点
        /// </summary>
        public List<int> ExistingHours { get; set; } = new List<int>();

        /// <summary>
        /// 缺失的时间点
        /// </summary>
        public List<int> MissingHours { get; set; } = new List<int>();

        /// <summary>
        /// 是否完整
        /// </summary>
        public bool IsComplete { get; set; }
    }

    /// <summary>
    /// 补充文件信息
    /// </summary>
    public class SupplementFileInfo
    {
        /// <summary>
        /// 目标日期
        /// </summary>
        public DateTime TargetDate { get; set; }

        /// <summary>
        /// 目标时间点
        /// </summary>
        public int TargetHour { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 源文件
        /// </summary>
        public ExcelFile SourceFile { get; set; } = null!;

        /// <summary>
        /// 目标文件名
        /// </summary>
        public string TargetFileName { get; set; } = string.Empty;
    }

    /// <summary>
    /// 数据质量报告
    /// </summary>
    public class DataQualityReport
    {
        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 总行数
        /// </summary>
        public int TotalRows { get; set; }

        /// <summary>
        /// 有效行数
        /// </summary>
        public int ValidRows { get; set; }

        /// <summary>
        /// 缺失行数
        /// </summary>
        public int MissingRows { get; set; }

        /// <summary>
        /// 整体完整性百分比
        /// </summary>
        public double OverallCompleteness { get; set; }

        /// <summary>
        /// 每个文件的质量信息
        /// </summary>
        public List<FileQualityInfo> FileQuality { get; set; } = new List<FileQualityInfo>();
    }

    /// <summary>
    /// 文件质量信息
    /// </summary>
    public class FileQualityInfo
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 总行数
        /// </summary>
        public int TotalRows { get; set; }

        /// <summary>
        /// 有效行数
        /// </summary>
        public int ValidRows { get; set; }

        /// <summary>
        /// 缺失行数
        /// </summary>
        public int MissingRows { get; set; }

        /// <summary>
        /// 全部缺失的行数
        /// </summary>
        public int AllMissingRows { get; set; }

        /// <summary>
        /// 平均完整性百分比
        /// </summary>
        public double AverageCompleteness { get; set; }
    }
} 