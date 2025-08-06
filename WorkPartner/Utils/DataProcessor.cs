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
            // 优先级策略：
            // 1. 前后相邻文件的平均值
            // 2. 同一天其他时间点的数据
            // 3. 最近有效数据的值

            var currentFile = allFiles[currentIndex];
            
            // 策略1：查找前后相邻文件
            var beforeValue = GetNearestValidValue(dataName, valueIndex, allFiles, currentIndex, searchBackward: true);
            var afterValue = GetNearestValidValue(dataName, valueIndex, allFiles, currentIndex, searchBackward: false);
            
            if (beforeValue.HasValue && afterValue.HasValue)
            {
                return (beforeValue.Value + afterValue.Value) / 2.0;
            }
            
            // 策略2：同一天其他时间点
            var sameDayFiles = allFiles.Where(f => f.Date.Date == currentFile.Date.Date && f != currentFile).ToList();
            var sameDayValues = new List<double>();
            
            foreach (var file in sameDayFiles)
            {
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataName);
                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    sameDayValues.Add(dataRow.Values[valueIndex].Value);
                }
            }
            
            if (sameDayValues.Any())
            {
                return sameDayValues.Average();
            }
            
            // 策略3：使用单个最近有效值
            if (beforeValue.HasValue)
            {
                return beforeValue.Value;
            }
            
            if (afterValue.HasValue)
            {
                return afterValue.Value;
            }
            
            // 策略4：作为最后手段，使用所有有效数据的平均值
            var allValidValues = new List<double>();
            foreach (var file in allFiles)
            {
                if (file == currentFile) continue;
                
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataName);
                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    allValidValues.Add(dataRow.Values[valueIndex].Value);
                }
            }
            
            return allValidValues.Any() ? allValidValues.Average() : null;
        }

        /// <summary>
        /// 获取最近的有效值
        /// </summary>
        /// <param name="dataName">数据名称</param>
        /// <param name="valueIndex">值索引</param>
        /// <param name="allFiles">所有文件列表</param>
        /// <param name="currentIndex">当前文件索引</param>
        /// <param name="searchBackward">是否向前搜索</param>
        /// <returns>最近的有效值</returns>
        private static double? GetNearestValidValue(string dataName, int valueIndex, List<ExcelFile> allFiles, int currentIndex, bool searchBackward)
        {
            var step = searchBackward ? -1 : 1;
            var startIndex = currentIndex + step;
            var endIndex = searchBackward ? 0 : allFiles.Count - 1;
            
            for (int i = startIndex; searchBackward ? i >= endIndex : i <= endIndex; i += step)
            {
                var file = allFiles[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataName);
                
                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    return dataRow.Values[valueIndex].Value;
                }
            }
            
            return null;
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

                foreach (var missingHour in dateCompleteness.MissingHours)
                {
                    // 选择最合适的源文件策略：
                    // 1. 同一天相同时间点的文件（如果有的话）
                    // 2. 同一天的其他时间点文件
                    // 3. 最近日期的相同时间点文件
                    // 4. 最近日期的任意文件
                    var sourceFile = SelectBestSourceFile(files, dateCompleteness.Date, missingHour, dateFiles);

                    if (sourceFile is null) continue;

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
        /// 选择最合适的源文件
        /// </summary>
        /// <param name="allFiles">所有文件列表</param>
        /// <param name="targetDate">目标日期</param>
        /// <param name="targetHour">目标时间</param>
        /// <param name="sameDayFiles">同一天的文件列表</param>
        /// <returns>最合适的源文件</returns>
        private static ExcelFile? SelectBestSourceFile(List<ExcelFile> allFiles, DateTime targetDate, int targetHour, List<ExcelFile> sameDayFiles)
        {
            // 策略1：同一天的其他时间点文件（优先选择）
            if (sameDayFiles.Any())
            {
                // 优先选择与目标时间最接近的时间点
                var bestSameDayFile = sameDayFiles
                    .OrderBy(f => Math.Abs(f.Hour - targetHour))
                    .FirstOrDefault();
                    
                if (bestSameDayFile != null)
                {
                    return bestSameDayFile;
                }
            }

            // 策略2：最近日期的相同时间点文件
            var sameHourFiles = allFiles.Where(f => f.Hour == targetHour).ToList();
            if (sameHourFiles.Any())
            {
                // 选择时间上最接近的日期
                var bestSameHourFile = sameHourFiles
                    .OrderBy(f => Math.Abs((f.Date.Date - targetDate).TotalDays))
                    .FirstOrDefault();
                    
                if (bestSameHourFile != null)
                {
                    return bestSameHourFile;
                }
            }

            // 策略3：最近日期的任意文件
            var nearestFile = allFiles
                .OrderBy(f => Math.Abs((f.Date.Date - targetDate).TotalDays))
                .ThenBy(f => Math.Abs(f.Hour - targetHour))
                .FirstOrDefault();

            return nearestFile;
        }

        /// <summary>
        /// 创建补充文件
        /// </summary>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>创建成功的文件数量</returns>
        public static int CreateSupplementFiles(List<SupplementFileInfo> supplementFiles, string outputDirectory)
        {
            if (supplementFiles == null || !supplementFiles.Any())
            {
                return 0;
            }

            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            int createdCount = 0;

            foreach (var supplementFile in supplementFiles)
            {
                try
                {
                    // 优先从输出目录中查找已处理的源文件
                    var processedSourcePath = Path.Combine(outputDirectory, supplementFile.SourceFile.FileName);
                    string sourceFilePath;
                    
                    if (File.Exists(processedSourcePath))
                    {
                        // 使用已处理的文件作为源文件
                        sourceFilePath = processedSourcePath;
                        Console.WriteLine($"✅ 使用已处理的源文件: {supplementFile.SourceFile.FileName}");
                    }
                    else
                    {
                        // 回退到原始文件
                        sourceFilePath = supplementFile.SourceFile.FilePath;
                        Console.WriteLine($"⚠️  使用原始源文件: {Path.GetFileName(sourceFilePath)}");
                    }
                    
                    var targetFilePath = Path.Combine(outputDirectory, supplementFile.TargetFileName);

                    // 复制源文件到目标位置
                    File.Copy(sourceFilePath, targetFilePath, true);
                    
                    createdCount++;
                    
                    Console.WriteLine($"✅ 已创建补充文件: {supplementFile.TargetFileName}");
                    Console.WriteLine($"   源文件: {Path.GetFileName(sourceFilePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 创建补充文件失败: {supplementFile.TargetFileName}");
                    Console.WriteLine($"   错误: {ex.Message}");
                }
            }

            return createdCount;
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