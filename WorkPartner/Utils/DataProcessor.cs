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
        /// <param name="files">文件列表</param>
        /// <returns>处理后的文件列表</returns>
        public static List<ExcelFile> ProcessMissingData(List<ExcelFile> files)
        {
            if (files == null || !files.Any())
            {
                return new List<ExcelFile>();
            }

            Console.WriteLine($"🔄 开始处理缺失数据，共 {files.Count} 个文件...");
            
            // 按时间顺序排序
            var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

            // 创建缓存以提高性能
            var valueCache = new Dictionary<string, Dictionary<int, List<double>>>();
            
            // 预处理：为每个数据名称和值索引创建有效值缓存
            Console.WriteLine("📊 预处理数据缓存...");
            PreprocessValueCache(sortedFiles, valueCache);

            // 处理每个文件中的缺失数据
            var totalFiles = sortedFiles.Count;
            var processedCount = 0;
            var lastProgressTime = DateTime.Now;
            
            for (int i = 0; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                ProcessFileMissingDataOptimized(currentFile, sortedFiles, i, valueCache);
                
                processedCount++;
                
                // 每处理10个文件或每30秒显示一次进度
                if (processedCount % 10 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                {
                    var progress = (double)processedCount / totalFiles * 100;
                    Console.WriteLine($"📈 处理进度: {processedCount}/{totalFiles} ({progress:F1}%) - 当前文件: {currentFile.FileName}");
                    lastProgressTime = DateTime.Now;
                }
            }

            Console.WriteLine($"✅ 缺失数据处理完成，共处理 {totalFiles} 个文件");
            
            // 处理所有文件都为空的数据行
            Console.WriteLine("🔄 处理所有文件都为空的数据行...");
            ProcessAllEmptyDataRows(sortedFiles);
            
            return sortedFiles;
        }

        /// <summary>
        /// 预处理值缓存以提高性能
        /// </summary>
        private static void PreprocessValueCache(List<ExcelFile> files, Dictionary<string, Dictionary<int, List<double>>> valueCache)
        {
            foreach (var file in files)
            {
                foreach (var dataRow in file.DataRows)
                {
                    if (!valueCache.ContainsKey(dataRow.Name))
                    {
                        valueCache[dataRow.Name] = new Dictionary<int, List<double>>();
                    }
                    
                    for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                    {
                        if (!valueCache[dataRow.Name].ContainsKey(valueIndex))
                        {
                            valueCache[dataRow.Name][valueIndex] = new List<double>();
                        }
                        
                        if (dataRow.Values[valueIndex].HasValue)
                        {
                            valueCache[dataRow.Name][valueIndex].Add(dataRow.Values[valueIndex].Value);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 优化后的单个文件缺失数据处理
        /// </summary>
        private static void ProcessFileMissingDataOptimized(ExcelFile currentFile, List<ExcelFile> allFiles, int currentIndex, Dictionary<string, Dictionary<int, List<double>>> valueCache)
        {
            foreach (var dataRow in currentFile.DataRows)
            {
                for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                {
                    if (!dataRow.Values[valueIndex].HasValue)
                    {
                        // 使用缓存的优化计算补充值
                        var supplementValue = CalculateSupplementValueOptimized(dataRow.Name, valueIndex, allFiles, currentIndex, valueCache);
                        if (supplementValue.HasValue)
                        {
                            dataRow.Values[valueIndex] = supplementValue.Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 优化后的补充值计算
        /// </summary>
        private static double? CalculateSupplementValueOptimized(string dataName, int valueIndex, List<ExcelFile> allFiles, int currentIndex, Dictionary<string, Dictionary<int, List<double>>> valueCache)
        {
            var currentFile = allFiles[currentIndex];
            
            // 策略1：前后相邻文件的平均值
            var beforeValue = GetNearestValidValueOptimized(dataName, valueIndex, allFiles, currentIndex, searchBackward: true);
            var afterValue = GetNearestValidValueOptimized(dataName, valueIndex, allFiles, currentIndex, searchBackward: false);
            
            if (beforeValue.HasValue && afterValue.HasValue)
            {
                return (beforeValue.Value + afterValue.Value) / 2.0;
            }
            
            // 策略2：同一天其他时间点（使用缓存）
            if (valueCache.ContainsKey(dataName) && valueCache[dataName].ContainsKey(valueIndex))
            {
                var sameDayValues = new List<double>();
                var currentDate = currentFile.Date.Date;
                
                foreach (var file in allFiles)
                {
                    if (file.Date.Date == currentDate && file != currentFile)
                    {
                        var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataName);
                        if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                        {
                            sameDayValues.Add(dataRow.Values[valueIndex].Value);
                        }
                    }
                }
                
                if (sameDayValues.Any())
                {
                    return sameDayValues.Average();
                }
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
            
            // 策略4：使用全局平均值（从缓存中获取）
            if (valueCache.ContainsKey(dataName) && valueCache[dataName].ContainsKey(valueIndex) && valueCache[dataName][valueIndex].Any())
            {
                return valueCache[dataName][valueIndex].Average();
            }
            
            return null;
        }

        /// <summary>
        /// 优化后的最近有效值获取
        /// </summary>
        private static double? GetNearestValidValueOptimized(string dataName, int valueIndex, List<ExcelFile> allFiles, int currentIndex, bool searchBackward)
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
        /// 处理所有文件都为空的数据行，使用前一行和后一行的平均值
        /// </summary>
        /// <param name="files">文件列表</param>
        private static void ProcessAllEmptyDataRows(List<ExcelFile> files)
        {
            if (!files.Any()) return;

            // 获取所有唯一的数据行名称
            var allDataRowNames = files
                .SelectMany(f => f.DataRows)
                .Select(r => r.Name)
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            var processedCount = 0;
            var totalRows = allDataRowNames.Count;
            var lastProgressTime = DateTime.Now;

            foreach (var dataRowName in allDataRowNames)
            {
                // 检查该数据行在所有文件中的值
                var allValuesForThisRow = new List<double?>();
                var maxValueCount = 0;

                // 收集所有文件中该数据行的所有值
                foreach (var file in files)
                {
                    var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                    if (dataRow != null)
                    {
                        allValuesForThisRow.AddRange(dataRow.Values);
                        maxValueCount = Math.Max(maxValueCount, dataRow.Values.Count);
                    }
                }

                // 检查每个值索引位置是否所有文件都为空
                for (int valueIndex = 0; valueIndex < maxValueCount; valueIndex++)
                {
                    var valuesAtThisIndex = new List<double?>();
                    
                    foreach (var file in files)
                    {
                        var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                        if (dataRow != null && valueIndex < dataRow.Values.Count)
                        {
                            valuesAtThisIndex.Add(dataRow.Values[valueIndex]);
                        }
                    }

                    // 如果该索引位置的所有值都为空，则使用前一行和后一行的平均值
                    if (valuesAtThisIndex.Any() && valuesAtThisIndex.All(v => !v.HasValue))
                    {
                        var supplementValue = CalculateAverageFromAdjacentRows(files, dataRowName, valueIndex);
                        if (supplementValue.HasValue)
                        {
                            // 为所有文件中的该数据行补充值
                            foreach (var file in files)
                            {
                                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                                if (dataRow != null && valueIndex < dataRow.Values.Count)
                                {
                                    dataRow.Values[valueIndex] = supplementValue.Value;
                                }
                            }
                        }
                    }
                }

                processedCount++;
                
                // 每处理10个数据行或每30秒显示一次进度
                if (processedCount % 10 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                {
                    var progress = (double)processedCount / totalRows * 100;
                    Console.WriteLine($"📈 空行处理进度: {processedCount}/{totalRows} ({progress:F1}%) - 当前数据行: {dataRowName}");
                    lastProgressTime = DateTime.Now;
                }
            }

            Console.WriteLine($"✅ 空行数据处理完成，共处理 {totalRows} 个数据行");
        }

        /// <summary>
        /// 计算前一行和后一行数据的平均值
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="currentDataRowName">当前数据行名称</param>
        /// <param name="valueIndex">值索引</param>
        /// <returns>平均值</returns>
        private static double? CalculateAverageFromAdjacentRows(List<ExcelFile> files, string currentDataRowName, int valueIndex)
        {
            // 获取所有数据行名称，按名称排序
            var allDataRowNames = files
                .SelectMany(f => f.DataRows)
                .Select(r => r.Name)
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            var currentIndex = allDataRowNames.IndexOf(currentDataRowName);
            if (currentIndex == -1) return null;

            var beforeValue = GetValueFromAdjacentRow(files, allDataRowNames, currentIndex - 1, valueIndex);
            var afterValue = GetValueFromAdjacentRow(files, allDataRowNames, currentIndex + 1, valueIndex);

            if (beforeValue.HasValue && afterValue.HasValue)
            {
                return (beforeValue.Value + afterValue.Value) / 2.0;
            }
            else if (beforeValue.HasValue)
            {
                return beforeValue.Value;
            }
            else if (afterValue.HasValue)
            {
                return afterValue.Value;
            }

            return null;
        }

        /// <summary>
        /// 从相邻行获取值
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="allDataRowNames">所有数据行名称</param>
        /// <param name="targetIndex">目标索引</param>
        /// <param name="valueIndex">值索引</param>
        /// <returns>值</returns>
        private static double? GetValueFromAdjacentRow(List<ExcelFile> files, List<string> allDataRowNames, int targetIndex, int valueIndex)
        {
            if (targetIndex < 0 || targetIndex >= allDataRowNames.Count)
                return null;

            var targetDataRowName = allDataRowNames[targetIndex];
            var validValues = new List<double>();

            foreach (var file in files)
            {
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == targetDataRowName);
                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    validValues.Add(dataRow.Values[valueIndex].Value);
                }
            }

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
                // 空列表时认为所有数据都是完整的
                result.IsAllComplete = true;
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