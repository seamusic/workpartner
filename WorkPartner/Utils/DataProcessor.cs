using WorkPartner.Models;
using WorkPartner.Services;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace WorkPartner.Utils
{

    /// <summary>
    /// 数据处理工具类
    /// </summary>
    public static class DataProcessor
    {
        /// <summary>
        /// 判断是否为累计变化量列（G列）
        /// </summary>
        private static bool IsCumulativeColumn(string columnName, DataProcessorConfig config)
        {
            return columnName.StartsWith(config.CumulativeColumnPrefix);
        }

        /// <summary>
        /// 判断是否为本期变化量列（D列）
        /// </summary>
        private static bool IsChangeColumn(string columnName, DataProcessorConfig config)
        {
            return columnName.StartsWith(config.ChangeColumnPrefix);
        }

        /// <summary>
        /// 根据累计列名称获取对应的变化列名称
        /// </summary>
        private static string GetBaseColumnForCumulative(string cumulativeColumnName, DataProcessorConfig config)
        {
            if (cumulativeColumnName.StartsWith(config.CumulativeColumnPrefix))
            {
                return cumulativeColumnName.Replace(config.CumulativeColumnPrefix, config.ChangeColumnPrefix);
            }
            return cumulativeColumnName;
        }

        /// <summary>
        /// 处理缺失数据（支持配置）
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="config">配置参数</param>
        /// <returns>处理后的文件列表</returns>
        public static List<ExcelFile> ProcessMissingData(List<ExcelFile> files, DataProcessorConfig? config = null)
        {
            config ??= DataProcessorConfig.Default;

            // 初始化性能监控
            var metrics = new PerformanceMetrics
            {
                StartTime = DateTime.Now
            };

            // 初始化缓存
            var cache = config.EnableCaching ? new DataCache(config.MaxCacheSize, config.CacheExpirationMinutes) : null;

            Console.WriteLine("🚀 开始处理缺失数据...");
            Console.WriteLine($"⚙️ 配置: 缓存={config.EnableCaching}, 批量处理={config.EnableBatchProcessing}, 批次大小={config.BatchSize}");

            if (files == null || !files.Any())
            {
                Console.WriteLine("⚠️ 文件列表为空，无需处理");
                return new List<ExcelFile>();
            }

            // 验证数据有效性
            ValidateDataIntegrity(files);

            // 按时间顺序排序文件
            var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();
            Console.WriteLine($"📊 共 {sortedFiles.Count} 个文件需要处理");

            //// 第一步：处理累计变化量
            //Console.WriteLine("📊 处理累计变化量...");
            //var cumulativeMetrics = ProcessCumulativeChangesOptimized(sortedFiles, config, cache, metrics);

            // 第二步：处理连续缺失数据的差异化
            Console.WriteLine("🔄 处理连续缺失数据的差异化...");
            var missingDataMetrics = ProcessConsecutiveMissingDataOptimized(sortedFiles, config, cache, metrics);

            // 第三步：处理补充文件数据微调
            Console.WriteLine("🔄 处理补充文件数据微调...");
            var supplementMetrics = ProcessSupplementFilesAdjustmentOptimized(sortedFiles, config, cache, metrics);

            // 创建缓存以提高性能
            var valueCache = new Dictionary<string, Dictionary<int, List<double>>>();

            // 预处理：为每个数据名称和值索引创建有效值缓存
            Console.WriteLine("📊 预处理数据缓存...");
            PreprocessValueCacheOptimized(sortedFiles, valueCache, cache, config);

            // 批量处理文件中的缺失数据
            var totalFiles = sortedFiles.Count;
            var processedCount = 0;
            var lastProgressTime = DateTime.Now;

            if (config.EnableBatchProcessing)
            {
                // 批量处理
                var batches = sortedFiles.Select((file, index) => new { file, index })
                                       .GroupBy(x => x.index / config.BatchSize)
                                       .Select(g => g.Select(x => x.file).ToList())
                                       .ToList();

                Console.WriteLine($"📦 将 {totalFiles} 个文件分为 {batches.Count} 个批次处理");

                foreach (var batch in batches)
                {
                    ProcessBatchOptimized(batch, sortedFiles, valueCache, cache, config, metrics);
                    processedCount += batch.Count;

                    // 显示进度
                    if (config.EnableDetailedLogging || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                    {
                        var progress = (double)processedCount / totalFiles * 100;
                        Console.WriteLine($"📈 处理进度: {processedCount}/{totalFiles} ({progress:F1}%) - 当前批次: {batch.Count} 个文件");
                        lastProgressTime = DateTime.Now;
                    }
                }
            }
            else
            {
                // 逐个处理
                for (int i = 0; i < sortedFiles.Count; i++)
                {
                    var currentFile = sortedFiles[i];
                    ProcessFileMissingDataOptimized(currentFile, sortedFiles, i, valueCache, cache, config, metrics);

                    processedCount++;

                    // 每处理10个文件或每30秒显示一次进度
                    if (processedCount % 10 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                    {
                        var progress = (double)processedCount / totalFiles * 100;
                        Console.WriteLine($"📈 处理进度: {processedCount}/{totalFiles} ({progress:F1}%) - 当前文件: {currentFile.FileName}");
                        lastProgressTime = DateTime.Now;
                    }
                }
            }

            // 处理所有文件都为空的数据行
            Console.WriteLine("🔄 处理所有文件都为空的数据行...");
            ProcessAllEmptyDataRows(sortedFiles);

            // 最终数据完整性检查：确保所有缺失值都被补充
            Console.WriteLine("🔍 执行最终数据完整性检查...");
            var finalCheckResult = PerformFinalDataIntegrityCheck(sortedFiles);
            Console.WriteLine($"✅ 最终检查完成: 补充了 {finalCheckResult} 个缺失值");

            // 显示性能指标
            metrics.EndTime = DateTime.Now;
            metrics.TotalFilesProcessed = sortedFiles.Count;
            metrics.TotalDataRowsProcessed = sortedFiles.Sum(f => f.DataRows.Count);
            metrics.TotalValuesProcessed = sortedFiles.Sum(f => f.DataRows.Sum(r => r.Values.Count));

            if (config.EnablePerformanceMonitoring)
            {
                DisplayPerformanceMetrics(metrics);
            }

            Console.WriteLine("🎉 缺失数据处理完成！");
            return sortedFiles;
        }

        /// <summary>
        /// 处理缺失数据（保持向后兼容）
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>处理后的文件列表</returns>
        public static List<ExcelFile> ProcessMissingData(List<ExcelFile> files)
        {
            return ProcessMissingData(files, DataProcessorConfig.Default);
        }

        /// <summary>
        /// 验证并重新计算第4、5、6列的值，确保符合"1. 基本逻辑重构"的要求
        /// 逻辑：本期4列=本期6列值-上期6列值，本期5列=本期6列值-上期6列值，本期6列=本期6列值-上期6列值
        /// 如果变化量超过阈值，则重新计算累计值
        /// </summary>
        /// <param name="files">已处理的文件列表</param>
        /// <param name="config">配置参数</param>
        /// <returns>处理后的文件列表</returns>
        public static List<ExcelFile> ValidateAndRecalculateColumns456(List<ExcelFile> files, DataProcessorConfig? config = null)
        {
            config ??= DataProcessorConfig.Default;

            Console.WriteLine("🔍 开始验证并修正第4、5、6列数据，确保符合基本逻辑重构要求...");
            Console.WriteLine($"⚙️ 验证配置: 误差容忍度={config.ColumnValidationTolerance:P0}, 累计值调整阈值={config.CumulativeAdjustmentThreshold:F2}");

            if (files == null || !files.Any())
            {
                Console.WriteLine("⚠️ 文件列表为空，无需验证");
                return new List<ExcelFile>();
            }

            // 按时间顺序排序文件
            var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();
            var totalColumnCorrections = 0;
            var totalCumulativeAdjustments = 0;

            Console.WriteLine($"📊 开始验证并修正 {sortedFiles.Count} 个文件的数据逻辑...");

            // 从第二个文件开始处理（需要上一期的数据）
            for (int i = 1; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                var previousFile = sortedFiles[i - 1];
                var fileColumnCorrections = 0;
                var fileCumulativeAdjustments = 0;

                Console.WriteLine($"\n📅 处理文件: {currentFile.Date:yyyy-MM-dd} {currentFile.Hour:D2}:00 (对比上一期: {previousFile.Date:yyyy-MM-dd} {previousFile.Hour:D2}:00)");

                foreach (var dataRow in currentFile.DataRows)
                {
                    // 验证并修正第4、5、6列的值
                    var (columnCorrections, cumulativeAdjustments) = ValidateAndCorrectRowColumns456(dataRow, previousFile, config);
                    fileColumnCorrections += columnCorrections;
                    fileCumulativeAdjustments += cumulativeAdjustments;
                }

                if (fileColumnCorrections > 0 || fileCumulativeAdjustments > 0)
                {
                    Console.WriteLine($"📊 文件 {currentFile.Date:yyyy-MM-dd} {currentFile.Hour:D2}:00 修正完成:");
                    Console.WriteLine($"   - 修正第4、5、6列值: {fileColumnCorrections} 个");
                    Console.WriteLine($"   - 调整累计值: {fileCumulativeAdjustments} 个");
                }

                totalColumnCorrections += fileColumnCorrections;
                totalCumulativeAdjustments += fileCumulativeAdjustments;
            }

            Console.WriteLine($"\n✅ 第4、5、6列验证和修正完成:");
            Console.WriteLine($"   - 总修正第4、5、6列值: {totalColumnCorrections} 个");
            Console.WriteLine($"   - 总调整累计值: {totalCumulativeAdjustments} 个");
            Console.WriteLine($"   - 总修正操作: {totalColumnCorrections + totalCumulativeAdjustments} 个");

            return sortedFiles;
        }

        /// <summary>
        /// 验证并修正单个数据行的第4、5、6列值
        /// </summary>
        /// <param name="currentRow">当前数据行</param>
        /// <param name="previousFile">上一期文件</param>
        /// <param name="config">配置参数</param>
        /// <returns>(修正的列值数量, 调整的累计值数量)</returns>
        private static (int ColumnCorrections, int CumulativeAdjustments) ValidateAndCorrectRowColumns456(DataRow currentRow, ExcelFile previousFile, DataProcessorConfig config)
        {
            var columnCorrections = 0;
            var cumulativeAdjustments = 0;

            // 获取上一期对应的数据行
            var previousRow = previousFile.DataRows.FirstOrDefault(r => r.Name == currentRow.Name);
            if (previousRow == null) return (0, 0);

            // 检查第4、5、6列（索引为3、4、5）
            var columnsToCheck = new[] { 3, 4, 5 }; // 对应第4、5、6列
            var baseColumnIndex = 5; // 第6列作为基准列（累计值）

            // 确保基准列（第6列）有值
            if (!currentRow.Values[baseColumnIndex].HasValue || !previousRow.Values[baseColumnIndex].HasValue)
            {
                return (0, 0);
            }

            foreach (var columnIndex in columnsToCheck)
            {
                // 确保列索引在有效范围内
                if (columnIndex >= currentRow.Values.Count)
                    continue;

                var currentCumulativeValue = currentRow.Values[columnIndex].Value;
                var previousCumulativeValue = previousRow.Values[columnIndex].Value;
                // 计算期望的变化量：本期变化量 = 本期累计值 - 上期累计值
                var expectedChangeAmount = currentCumulativeValue - previousCumulativeValue;


                // 检查是否需要调整累计值
                if (Math.Abs(expectedChangeAmount) > config.CumulativeAdjustmentThreshold)
                {
                    // 变化量超过阈值，需要重新计算累计值
                    Console.WriteLine($"    ⚠️ 变化量 {Math.Abs(expectedChangeAmount):F2} 超过阈值 {config.CumulativeAdjustmentThreshold:F2}，需要调整累计值");

                    var currentValue1 = currentRow.Values[columnIndex - 3];
                    if (currentValue1.HasValue)
                    {

                        // 重新计算累计值：新累计值 = 上期累计值 + 期望变化量
                        var newCumulativeValue = previousCumulativeValue + currentValue1.Value;

                        // 如果调整幅度过大，采用保守策略
                        var adjustmentAmount = Math.Abs(newCumulativeValue - currentCumulativeValue);
                        if (adjustmentAmount > config.CumulativeAdjustmentThreshold * 2)
                        {
                            // 调整幅度过大，采用保守策略：使用当前变化量的平均值
                            var conservativeChangeAmount = expectedChangeAmount * 0.5; // 使用50%的变化量
                            newCumulativeValue = previousCumulativeValue + conservativeChangeAmount;
                            Console.WriteLine($"    🔧 采用保守策略: 变化量从 {expectedChangeAmount:F2} 调整为 {conservativeChangeAmount:F2}");
                        }

                        // 应用新的累计值
                        currentRow.Values[columnIndex] = newCumulativeValue;
                        cumulativeAdjustments++;

                        Console.WriteLine($"    🔧 调整累计值: {currentCumulativeValue:F2} → {newCumulativeValue:F2}");
                    }
                }

                var currentValue = currentRow.Values[columnIndex - 3];
                var isCurrentColumnHasValue = currentValue.HasValue;

                if (isCurrentColumnHasValue)
                {
                    // 如果当前列有值，检查是否符合逻辑
                    var actualChangeAmount = currentValue.Value;
                    var difference = Math.Abs(actualChangeAmount - expectedChangeAmount);

                    if (difference > config.ColumnValidationTolerance)
                    {
                        // 变化量不符合期望，需要修正
                        Console.WriteLine($"    🔄 修正第{columnIndex + 1}列: 当前值={actualChangeAmount:F2}, 期望值={expectedChangeAmount:F2}, 差异={difference:F2}");
                        currentRow.Values[columnIndex-3] = expectedChangeAmount;
                        columnCorrections++;
                    }
                }
                else
                {
                    // 如果当前列为空，直接填入期望的变化量
                    Console.WriteLine($"    ➕ 填充第{columnIndex + 1}列: 期望变化量={expectedChangeAmount:F2}");
                    currentRow.Values[columnIndex-3] = expectedChangeAmount;
                    columnCorrections++;
                }
            }

            return (columnCorrections, cumulativeAdjustments);
        }

        /// <summary>
        /// 调整累计值以修正变化量过大的问题
        /// </summary>
        /// <param name="currentRow">当前数据行</param>
        /// <param name="previousRow">上一期数据行</param>
        /// <param name="changeColumnIndex">变化列索引（第4、5、6列）</param>
        /// <param name="cumulativeColumnIndex">累计列索引（第6列）</param>
        /// <param name="currentValue">当前值</param>
        /// <param name="expectedValue">期望值</param>
        /// <param name="config">配置参数</param>
        /// <returns>是否进行了调整</returns>
        private static bool AdjustCumulativeValue(DataRow currentRow, DataRow previousRow,
            int changeColumnIndex, int cumulativeColumnIndex,
            double currentValue, double expectedValue, DataProcessorConfig config)
        {
            try
            {
                // 获取上一期的累计值
                var previousCumulativeValue = previousRow.Values[cumulativeColumnIndex];
                if (!previousCumulativeValue.HasValue)
                    return false;

                // 计算新的累计值：新累计值 = 上期累计值 + 当前变化值
                var newCumulativeValue = previousCumulativeValue.Value + currentValue;

                // 检查调整后的累计值是否合理
                var adjustmentAmount = Math.Abs(newCumulativeValue - currentRow.Values[cumulativeColumnIndex].Value);

                // 如果调整幅度过大，可能需要进一步处理
                if (adjustmentAmount > config.CumulativeAdjustmentThreshold * 2)
                {
                    // 调整幅度过大，采用更保守的策略
                    // 将累计值调整为：上期累计值 + 期望变化值
                    newCumulativeValue = previousCumulativeValue.Value + expectedValue;
                    Console.WriteLine($"⚠️ 调整幅度过大，采用保守策略: {currentRow.Name} 第{cumulativeColumnIndex + 1}列");
                }

                // 应用新的累计值
                currentRow.Values[cumulativeColumnIndex] = newCumulativeValue;

                Console.WriteLine($"🔧 累计值调整: {currentRow.Name} 第{cumulativeColumnIndex + 1}列: {currentRow.Values[cumulativeColumnIndex]:F2} → {newCumulativeValue:F2}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 累计值调整失败 {currentRow.Name}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 处理累计变化量计算
        /// </summary>
        private static void ProcessCumulativeChanges(List<ExcelFile> sortedFiles, DataProcessorConfig config)
        {
            var random = new Random(config.RandomSeed);

            for (int i = 1; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                var previousFile = sortedFiles[i - 1];

                foreach (var dataRow in currentFile.DataRows)
                {
                    if (IsCumulativeColumn(dataRow.Name, config))
                    {
                        ProcessCumulativeRow(dataRow, previousFile, currentFile, config, random);
                    }
                }
            }
        }

        /// <summary>
        /// 处理累计变化量计算（优化版本）
        /// </summary>
        private static (int Calculations, int CacheHits, int CacheMisses) ProcessCumulativeChangesOptimized(
            List<ExcelFile> sortedFiles, DataProcessorConfig config, DataCache? cache, PerformanceMetrics metrics)
        {
            var random = new Random(config.RandomSeed);
            var calculations = 0;
            var cacheHits = 0;
            var cacheMisses = 0;

            for (int i = 1; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                var previousFile = sortedFiles[i - 1];

                foreach (var dataRow in currentFile.DataRows)
                {
                    if (IsCumulativeColumn(dataRow.Name, config))
                    {
                        var (calc, hits, misses) = ProcessCumulativeRowOptimized(dataRow, previousFile, currentFile, config, random, cache);
                        calculations += calc;
                        cacheHits += hits;
                        cacheMisses += misses;
                    }
                }
            }

            // 更新性能指标
            metrics.CumulativeCalculations = calculations;
            metrics.CacheHits += cacheHits;
            metrics.CacheMisses += cacheMisses;

            return (calculations, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 处理单个累计变化量行
        /// </summary>
        private static void ProcessCumulativeRow(DataRow cumulativeRow, ExcelFile previousFile, ExcelFile currentFile, DataProcessorConfig config, Random random)
        {
            var baseColumnName = GetBaseColumnForCumulative(cumulativeRow.Name, config);

            for (int valueIndex = 0; valueIndex < cumulativeRow.Values.Count; valueIndex++)
            {
                if (!cumulativeRow.Values[valueIndex].HasValue)
                {
                    var previousCumulative = GetPreviousCumulativeValue(previousFile, cumulativeRow.Name, valueIndex);
                    var currentChange = GetCurrentChangeValue(currentFile, baseColumnName, valueIndex);

                    if (previousCumulative.HasValue && currentChange.HasValue)
                    {
                        // 计算累计量：本期累计 = 上期累计 + 本期变化
                        var newCumulative = previousCumulative.Value + currentChange.Value;
                        cumulativeRow.Values[valueIndex] = newCumulative;
                    }
                    else if (previousCumulative.HasValue)
                    {
                        // 如果只有上期累计值，使用上期值
                        cumulativeRow.Values[valueIndex] = previousCumulative.Value;
                    }
                    else if (currentChange.HasValue)
                    {
                        // 如果只有本期变化值，使用变化值作为累计值
                        cumulativeRow.Values[valueIndex] = currentChange.Value;
                    }
                }
            }
        }

        /// <summary>
        /// 获取上期累计值
        /// </summary>
        private static double? GetPreviousCumulativeValue(ExcelFile previousFile, string columnName, int valueIndex)
        {
            var dataRow = previousFile.DataRows.FirstOrDefault(r => r.Name == columnName);
            return dataRow?.Values.ElementAtOrDefault(valueIndex);
        }

        /// <summary>
        /// 获取本期变化值
        /// </summary>
        private static double? GetCurrentChangeValue(ExcelFile currentFile, string columnName, int valueIndex)
        {
            var dataRow = currentFile.DataRows.FirstOrDefault(r => r.Name == columnName);
            return dataRow?.Values.ElementAtOrDefault(valueIndex);
        }

        /// <summary>
        /// 处理连续缺失数据的差异化（优化版本）
        /// </summary>
        private static (int Processings, int CacheHits, int CacheMisses) ProcessConsecutiveMissingDataOptimized(
            List<ExcelFile> files, DataProcessorConfig config, DataCache? cache, PerformanceMetrics metrics)
        {
            Console.WriteLine("🔍 开始识别连续缺失时间段...");
            var missingPeriods = IdentifyMissingPeriods(files);
            Console.WriteLine($"📊 识别到 {missingPeriods.Count} 个连续缺失时间段");

            var processings = 0;
            var cacheHits = 0;
            var cacheMisses = 0;
            var totalPeriods = missingPeriods.Count;
            var lastProgressTime = DateTime.Now;

            for (int i = 0; i < missingPeriods.Count; i++)
            {
                var period = missingPeriods[i];
                var currentTime = DateTime.Now;

                // 显示当前处理的缺失时间段信息
                if (config.EnableDetailedLogging || (currentTime - lastProgressTime).TotalSeconds >= 10)
                {
                    var progress = (double)(i + 1) / totalPeriods * 100;
                    Console.WriteLine($"🔄 处理进度: {i + 1}/{totalPeriods} ({progress:F1}%) - 当前处理: {period.StartTime:yyyy-MM-dd HH:mm} 到 {period.EndTime:yyyy-MM-dd HH:mm}, 缺失 {period.MissingHours} 小时");
                    lastProgressTime = currentTime;
                }

                var (proc, hits, misses) = ProcessMissingPeriodOptimized(period, files, config, cache);
                processings += proc;
                cacheHits += hits;
                cacheMisses += misses;

                // 每处理5个时间段显示一次详细进度
                if ((i + 1) % 5 == 0)
                {
                    Console.WriteLine($"📈 已处理 {i + 1}/{totalPeriods} 个时间段，累计处理 {processings} 个缺失数据点");
                }
            }

            // 更新性能指标
            metrics.MissingDataProcessings = processings;
            metrics.CacheHits += cacheHits;
            metrics.CacheMisses += cacheMisses;

            Console.WriteLine($"✅ 连续缺失数据处理完成，共处理 {processings} 个缺失数据点");
            return (processings, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 识别连续缺失时间段
        /// </summary>
        private static List<MissingPeriod> IdentifyMissingPeriods(List<ExcelFile> sortedFiles)
        {
            var missingPeriods = new List<MissingPeriod>();

            // 获取所有数据行名称
            var allDataRowNames = sortedFiles
                .SelectMany(f => f.DataRows)
                .Select(r => r.Name)
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            Console.WriteLine($"🔍 开始分析 {allDataRowNames.Count} 个数据行的缺失情况...");
            var lastProgressTime = DateTime.Now;

            // 为每个数据行识别连续缺失时间段
            for (int i = 0; i < allDataRowNames.Count; i++)
            {
                var dataRowName = allDataRowNames[i];
                var currentTime = DateTime.Now;

                // 每处理10个数据行或每15秒显示一次进度
                if ((i + 1) % 10 == 0 || (currentTime - lastProgressTime).TotalSeconds >= 15)
                {
                    var progress = (double)(i + 1) / allDataRowNames.Count * 100;
                    Console.WriteLine($"🔍 识别进度: {i + 1}/{allDataRowNames.Count} ({progress:F1}%) - 当前分析: {dataRowName}");
                    lastProgressTime = currentTime;
                }

                var periodsForRow = IdentifyMissingPeriodsForDataRow(dataRowName, sortedFiles);
                missingPeriods.AddRange(periodsForRow);
            }

            Console.WriteLine($"📊 识别完成，共发现 {missingPeriods.Count} 个缺失时间段");

            // 合并重叠的时间段
            Console.WriteLine("🔄 开始合并重叠的时间段...");
            var mergedPeriods = MergeOverlappingPeriods(missingPeriods);
            Console.WriteLine($"✅ 合并完成，最终有 {mergedPeriods.Count} 个时间段");

            return mergedPeriods;
        }

        /// <summary>
        /// 为单个数据行识别连续缺失时间段
        /// </summary>
        private static List<MissingPeriod> IdentifyMissingPeriodsForDataRow(string dataRowName, List<ExcelFile> sortedFiles)
        {
            var periods = new List<MissingPeriod>();
            var missingTimes = new List<DateTime>();
            var currentPeriodStart = (DateTime?)null;

            for (int i = 0; i < sortedFiles.Count; i++)
            {
                var file = sortedFiles[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);

                if (dataRow == null) continue;

                // 检查该数据行是否有缺失值
                var hasMissingValues = dataRow.Values.Any(v => !v.HasValue);

                if (hasMissingValues)
                {
                    if (currentPeriodStart == null)
                    {
                        currentPeriodStart = file.Date.AddHours(file.Hour);
                    }
                    missingTimes.Add(file.Date.AddHours(file.Hour));
                }
                else
                {
                    // 如果当前有缺失时间段，则结束该时间段
                    if (currentPeriodStart.HasValue && missingTimes.Any())
                    {
                        var period = new MissingPeriod
                        {
                            StartTime = currentPeriodStart.Value,
                            EndTime = missingTimes.Last(),
                            MissingHours = missingTimes.Count,
                            MissingTimes = new List<DateTime>(missingTimes),
                            MissingDataRows = new List<string> { dataRowName },
                            PreviousValidTime = GetPreviousValidTime(sortedFiles, i),
                            NextValidTime = GetNextValidTime(sortedFiles, i)
                        };

                        periods.Add(period);

                        // 重置
                        currentPeriodStart = null;
                        missingTimes.Clear();
                    }
                }
            }

            // 处理最后一个时间段（如果文件末尾有缺失）
            if (currentPeriodStart.HasValue && missingTimes.Any())
            {
                var period = new MissingPeriod
                {
                    StartTime = currentPeriodStart.Value,
                    EndTime = missingTimes.Last(),
                    MissingHours = missingTimes.Count,
                    MissingTimes = new List<DateTime>(missingTimes),
                    MissingDataRows = new List<string> { dataRowName },
                    PreviousValidTime = GetPreviousValidTime(sortedFiles, sortedFiles.Count - 1),
                    NextValidTime = null
                };

                periods.Add(period);
            }

            return periods;
        }

        /// <summary>
        /// 获取前一个有效时间点
        /// </summary>
        private static DateTime? GetPreviousValidTime(List<ExcelFile> sortedFiles, int currentIndex)
        {
            for (int i = currentIndex - 1; i >= 0; i--)
            {
                var file = sortedFiles[i];
                if (file.DataRows.Any(r => r.Values.Any(v => v.HasValue)))
                {
                    return file.Date.AddHours(file.Hour);
                }
            }
            return null;
        }

        /// <summary>
        /// 获取后一个有效时间点
        /// </summary>
        private static DateTime? GetNextValidTime(List<ExcelFile> sortedFiles, int currentIndex)
        {
            for (int i = currentIndex + 1; i < sortedFiles.Count; i++)
            {
                var file = sortedFiles[i];
                if (file.DataRows.Any(r => r.Values.Any(v => v.HasValue)))
                {
                    return file.Date.AddHours(file.Hour);
                }
            }
            return null;
        }

        /// <summary>
        /// 合并重叠的时间段
        /// </summary>
        private static List<MissingPeriod> MergeOverlappingPeriods(List<MissingPeriod> periods)
        {
            if (!periods.Any()) return periods;

            var sortedPeriods = periods.OrderBy(p => p.StartTime).ToList();
            var mergedPeriods = new List<MissingPeriod>();
            var currentPeriod = sortedPeriods[0];

            for (int i = 1; i < sortedPeriods.Count; i++)
            {
                var nextPeriod = sortedPeriods[i];

                // 检查是否重叠或相邻
                if (nextPeriod.StartTime <= currentPeriod.EndTime.AddHours(1))
                {
                    // 合并时间段
                    currentPeriod.EndTime = nextPeriod.EndTime;
                    currentPeriod.MissingHours += nextPeriod.MissingHours;
                    currentPeriod.MissingTimes.AddRange(nextPeriod.MissingTimes);
                    currentPeriod.MissingDataRows.AddRange(nextPeriod.MissingDataRows);
                    currentPeriod.MissingDataRows = currentPeriod.MissingDataRows.Distinct().ToList();

                    // 更新前后有效时间点
                    if (nextPeriod.PreviousValidTime.HasValue &&
                        (!currentPeriod.PreviousValidTime.HasValue ||
                         nextPeriod.PreviousValidTime.Value < currentPeriod.PreviousValidTime.Value))
                    {
                        currentPeriod.PreviousValidTime = nextPeriod.PreviousValidTime;
                    }

                    if (nextPeriod.NextValidTime.HasValue &&
                        (!currentPeriod.NextValidTime.HasValue ||
                         nextPeriod.NextValidTime.Value > currentPeriod.NextValidTime.Value))
                    {
                        currentPeriod.NextValidTime = nextPeriod.NextValidTime;
                    }
                }
                else
                {
                    // 不重叠，添加当前时间段并开始新的时间段
                    mergedPeriods.Add(currentPeriod);
                    currentPeriod = nextPeriod;
                }
            }

            // 添加最后一个时间段
            mergedPeriods.Add(currentPeriod);

            return mergedPeriods;
        }

        /// <summary>
        /// 处理单个缺失时间段
        /// </summary>
        private static void ProcessMissingPeriod(MissingPeriod period, List<ExcelFile> sortedFiles, DataProcessorConfig config)
        {
            var random = new Random(config.RandomSeed + period.StartTime.GetHashCode());

            foreach (var dataRowName in period.MissingDataRows)
            {
                foreach (var missingTime in period.MissingTimes)
                {
                    var file = sortedFiles.FirstOrDefault(f =>
                        f.Date.Date == missingTime.Date && f.Hour == missingTime.Hour);

                    if (file == null) continue;

                    var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                    if (dataRow == null) continue;

                    // 处理该数据行的所有缺失值
                    for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                    {
                        if (!dataRow.Values[valueIndex].HasValue)
                        {
                            var adjustedValue = CalculateAdjustedValueForMissingPoint(
                                dataRowName, valueIndex, missingTime, period, sortedFiles, config, random);

                            if (adjustedValue.HasValue)
                            {
                                dataRow.Values[valueIndex] = adjustedValue.Value;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 为缺失数据点计算调整后的值（重载版本，用于向后兼容）
        /// </summary>
        private static double? CalculateAdjustedValueForMissingPoint(
            string dataRowName, int valueIndex, DateTime missingTime,
            MissingPeriod period, List<ExcelFile> sortedFiles, DataProcessorConfig config, Random random)
        {
            // 获取前后有效值
            var previousValue = GetNearestValidValueForDataRow(dataRowName, valueIndex, sortedFiles, missingTime, true);
            var nextValue = GetNearestValidValueForDataRow(dataRowName, valueIndex, sortedFiles, missingTime, false);

            if (!previousValue.HasValue && !nextValue.HasValue)
            {
                return null; // 无法计算
            }

            // 计算基础值（前后值的平均值或单个值）
            double baseValue;
            if (previousValue.HasValue && nextValue.HasValue)
            {
                baseValue = (previousValue.Value + nextValue.Value) / 2.0;
            }
            else
            {
                baseValue = previousValue ?? nextValue.Value;
            }

            // 计算时间因子（基于在缺失时间段中的位置）
            var timeFactor = CalculateTimeFactor(missingTime, period, config);

            // 计算调整值
            var adjustment = CalculateAdjustment(baseValue, timeFactor, period, config, random);

            return baseValue + adjustment;
        }

        /// <summary>
        /// 为缺失数据点计算调整后的值（用于缺失时间段处理）
        /// </summary>
        private static double? CalculateAdjustedValueForMissingPoint(
            MissingDataPoint missingPoint, MissingPeriod period, DataProcessorConfig config)
        {
            if (!missingPoint.PreviousValue.HasValue && !missingPoint.NextValue.HasValue)
            {
                return null; // 无法计算
            }

            // 计算基础值（前后值的平均值或单个值）
            double baseValue;
            if (missingPoint.PreviousValue.HasValue && missingPoint.NextValue.HasValue)
            {
                baseValue = (missingPoint.PreviousValue.Value + missingPoint.NextValue.Value) / 2.0;
            }
            else
            {
                baseValue = missingPoint.PreviousValue ?? missingPoint.NextValue.Value;
            }

            // 计算时间因子（基于在缺失时间段中的位置）
            var timeFactor = CalculateTimeFactor(missingPoint.TimePoint, period, config);

            // 计算调整值
            var random = new Random(config.RandomSeed + missingPoint.TimePoint.GetHashCode());
            var adjustment = CalculateAdjustment(baseValue, timeFactor, period, config, random);

            return baseValue + adjustment;
        }

        /// <summary>
        /// 获取数据行的最近有效值（用于缺失时间段处理）
        /// </summary>
        private static double? GetNearestValidValueForDataRow(
            List<ExcelFile> files, string dataRowName, DateTime targetTime, bool searchBackward)
        {
            var targetDateTime = targetTime;
            var step = searchBackward ? -1 : 1;

            // 找到目标时间在文件列表中的位置
            var targetIndex = files.FindIndex(f =>
                f.Date.Date == targetDateTime.Date && f.Hour == targetDateTime.Hour);

            if (targetIndex == -1) return null;

            // 向前或向后搜索有效值
            for (int i = targetIndex + step; searchBackward ? i >= 0 : i < files.Count; i += step)
            {
                var file = files[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                // TODO:此处取所有的数值取平均，不正确。应该是对应列的平均值。
                if (dataRow != null && dataRow.Values.Any(v => v.HasValue))
                {
                    // 返回第一个有效值的平均值
                    var validValues = dataRow.Values.Where(v => v.HasValue).Select(v => v.Value).ToList();
                    return validValues.Any() ? validValues.Average() : null;
                }
            }

            return null;
        }

        /// <summary>
        /// 获取数据行的最近有效值（重载版本，用于向后兼容）
        /// </summary>
        private static double? GetNearestValidValueForDataRow(
            string dataRowName, int valueIndex, List<ExcelFile> sortedFiles, DateTime targetTime, bool searchBackward)
        {
            var targetDateTime = targetTime;
            var step = searchBackward ? -1 : 1;

            // 找到目标时间在文件列表中的位置
            var targetIndex = sortedFiles.FindIndex(f =>
                f.Date.Date == targetDateTime.Date && f.Hour == targetDateTime.Hour);

            if (targetIndex == -1) return null;

            // 向前或向后搜索有效值
            for (int i = targetIndex + step; searchBackward ? i >= 0 : i < sortedFiles.Count; i += step)
            {
                var file = sortedFiles[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);

                if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                {
                    return dataRow.Values[valueIndex].Value;
                }
            }

            return null;
        }

        /// <summary>
        /// 计算时间因子
        /// </summary>
        private static double CalculateTimeFactor(DateTime missingTime, MissingPeriod period, DataProcessorConfig config)
        {
            if (period.MissingTimes.Count <= 1) return 0.5;

            // 计算缺失时间在时间段中的相对位置（0-1之间）
            var totalDuration = (period.EndTime - period.StartTime).TotalHours;
            var timeFromStart = (missingTime - period.StartTime).TotalHours;
            var relativePosition = totalDuration > 0 ? timeFromStart / totalDuration : 0.5;

            // 应用时间因子权重
            return relativePosition * config.TimeFactorWeight;
        }

        /// <summary>
        /// 计算调整值
        /// </summary>
        private static double CalculateAdjustment(
            double baseValue, double timeFactor, MissingPeriod period,
            DataProcessorConfig config, Random random)
        {
            // 基础调整范围
            var baseAdjustmentRange = baseValue * config.AdjustmentRange;

            // 基于时间因子的调整
            var timeBasedAdjustment = baseAdjustmentRange * timeFactor;

            // 添加随机性
            var randomFactor = (random.NextDouble() - 0.5) * 2; // -1 到 1 之间
            var randomAdjustment = timeBasedAdjustment * randomFactor;

            // 确保调整值不小于最小调整值
            var finalAdjustment = Math.Abs(randomAdjustment) < config.MinimumAdjustment
                ? (randomAdjustment >= 0 ? config.MinimumAdjustment : -config.MinimumAdjustment)
                : randomAdjustment;

            return finalAdjustment;
        }

        /// <summary>
        /// 处理补充文件数据微调
        /// </summary>
        private static void ProcessSupplementFilesAdjustment(List<ExcelFile> sortedFiles, DataProcessorConfig config)
        {
            var supplementFiles = sortedFiles.Where(f => f.IsSupplementFile).ToList();

            if (!supplementFiles.Any())
            {
                Console.WriteLine("✅ 未发现需要微调的补充文件");
                return;
            }

            Console.WriteLine($"📊 发现 {supplementFiles.Count} 个需要微调的补充文件");

            foreach (var supplementFile in supplementFiles)
            {
                Console.WriteLine($"🔄 微调补充文件: {supplementFile.FileName}");
                AdjustSupplementFileData(supplementFile, config);
            }
        }

        /// <summary>
        /// 微调单个补充文件的数据
        /// </summary>
        private static void AdjustSupplementFileData(ExcelFile supplementFile, DataProcessorConfig config)
        {
            if (supplementFile.SupplementSource == null)
            {
                Console.WriteLine($"⚠️ 补充文件 {supplementFile.FileName} 缺少源文件信息，跳过微调");
                return;
            }

            var adjustmentParams = supplementFile.SupplementSource.AdjustmentParams;
            var random = new Random(adjustmentParams.RandomSeed);

            Console.WriteLine($"🔧 使用调整参数: 范围={adjustmentParams.AdjustmentRange:P0}, 种子={adjustmentParams.RandomSeed}");

            foreach (var dataRow in supplementFile.DataRows)
            {
                AdjustDataRowValues(dataRow, adjustmentParams, random, config);
            }
        }

        /// <summary>
        /// 微调数据行的值
        /// </summary>
        private static void AdjustDataRowValues(DataRow dataRow, AdjustmentParameters adjustmentParams, Random random, DataProcessorConfig config)
        {
            for (int i = 0; i < dataRow.Values.Count; i++)
            {
                if (dataRow.Values[i].HasValue)
                {
                    var originalValue = dataRow.Values[i].Value;
                    var adjustedValue = CalculateSupplementAdjustment(originalValue, adjustmentParams, random, config);
                    dataRow.Values[i] = adjustedValue;
                }
            }
        }

        /// <summary>
        /// 计算补充文件的调整值
        /// </summary>
        private static double CalculateSupplementAdjustment(double originalValue, AdjustmentParameters adjustmentParams, Random random, DataProcessorConfig config)
        {
            // 基础调整范围
            var baseAdjustmentRange = originalValue * adjustmentParams.AdjustmentRange;

            // 生成随机调整值
            var randomFactor = (random.NextDouble() - 0.5) * 2; // -1 到 1 之间
            var randomAdjustment = baseAdjustmentRange * randomFactor;

            // 应用相关性权重（如果启用）
            if (adjustmentParams.MaintainDataCorrelation)
            {
                var correlationFactor = 1.0 + (random.NextDouble() - 0.5) * 0.2; // ±10%的相关性变化
                correlationFactor = Math.Max(0.5, Math.Min(1.5, correlationFactor)); // 限制在0.5-1.5范围内
                randomAdjustment *= correlationFactor * adjustmentParams.CorrelationWeight;
            }

            // 确保调整值不小于最小调整值
            var finalAdjustment = Math.Abs(randomAdjustment) < adjustmentParams.MinimumAdjustment
                ? (randomAdjustment >= 0 ? adjustmentParams.MinimumAdjustment : -adjustmentParams.MinimumAdjustment)
                : randomAdjustment;

            return originalValue + finalAdjustment;
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
        /// 增强的数据补充值计算（确保所有缺失值都能被补充）
        /// </summary>
        private static double? CalculateSupplementValueOptimized(string dataName, int valueIndex, List<ExcelFile> allFiles, int currentIndex, Dictionary<string, Dictionary<int, List<double>>> valueCache)
        {
            var currentFile = allFiles[currentIndex];

            // 策略1：前后相邻文件的平均值（优先策略）
            var beforeValue = GetNearestValidValueOptimized(dataName, valueIndex, allFiles, currentIndex, searchBackward: true);
            var afterValue = GetNearestValidValueOptimized(dataName, valueIndex, allFiles, currentIndex, searchBackward: false);

            if (beforeValue.HasValue && afterValue.HasValue)
            {
                return (beforeValue.Value + afterValue.Value) / 2.0;
            }

            // 策略2：同一天其他时间点的平均值
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

            // 策略5：使用相邻数据行的平均值（作为最后的备选方案）
            var adjacentValue = CalculateAverageFromAdjacentRows(allFiles, dataName, valueIndex);
            if (adjacentValue.HasValue)
            {
                return adjacentValue.Value;
            }

            // 策略6：如果所有策略都失败，使用默认值（基于数据类型）
            return GetDefaultValueForDataType(dataName);
        }

        /// <summary>
        /// 根据数据类型获取默认值
        /// </summary>
        private static double GetDefaultValueForDataType(string dataName)
        {
            // 根据数据名称的特征判断数据类型，返回合理的默认值
            if (dataName.Contains("温度") || dataName.Contains("温度"))
            {
                return 20.0; // 室温
            }
            else if (dataName.Contains("湿度"))
            {
                return 60.0; // 标准湿度
            }
            else if (dataName.Contains("压力"))
            {
                return 101.325; // 标准大气压
            }
            else if (dataName.Contains("流量"))
            {
                return 0.0; // 流量默认为0
            }
            else if (dataName.StartsWith("G")) // 累计值
            {
                return 0.0; // 累计值默认为0
            }
            else if (dataName.StartsWith("D")) // 变化值
            {
                return 0.0; // 变化值默认为0
            }
            else
            {
                return 0.0; // 通用默认值
            }
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

            // 通用处理：检查所有行的缺失数据问题
            Console.WriteLine("🔍 通用处理：检查所有行的缺失数据问题...");
            ProcessAllRowsMissingData(files);

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
                        else
                        {
                            // 如果无法计算相邻行平均值，使用默认值
                            var defaultValue = GetDefaultValueForDataType(dataRowName);
                            foreach (var file in files)
                            {
                                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                                if (dataRow != null && valueIndex < dataRow.Values.Count)
                                {
                                    dataRow.Values[valueIndex] = defaultValue;
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
        /// 通用处理所有行的缺失数据问题
        /// 检查每一行是否存在D列到I列的空值，如果存在则使用相邻行的平均值补充
        /// 同时确保累计变化量的逻辑正确：G行 = G(行-1) + D行
        /// </summary>
        /// <param name="files">文件列表</param>
        private static void ProcessAllRowsMissingData(List<ExcelFile> files)
        {
            if (!files.Any()) return;

            // 获取配置
            var config = DataProcessorConfig.Default;

            // 获取所有可能的行索引
            var allRowIndices = files
                .SelectMany(f => f.DataRows)
                .Select(r => r.RowIndex)
                .Distinct()
                .OrderBy(rowIndex => rowIndex)
                .ToList();

            Console.WriteLine($"🔍 开始检查 {allRowIndices.Count} 个数据行的缺失数据问题...");
            Console.WriteLine($"📊 行索引范围: {allRowIndices.Min()} - {allRowIndices.Max()}");

            // 特殊处理：强制检查第200行
            var specialRows = new[] { 185, 200 };
            foreach (var specialRow in specialRows)
            {
                if (!allRowIndices.Contains(specialRow))
                {
                    Console.WriteLine($"⚠️ 警告：第{specialRow}行不在allRowIndices中，可能存在问题");
                }
                else
                {
                    Console.WriteLine($"✅ 第{specialRow}行在allRowIndices中，索引位置: {allRowIndices.IndexOf(specialRow)}");
                }
            }

            foreach (var file in files)
            {
                Console.WriteLine($"📁 检查文件: {file.FileName}");

                foreach (var rowIndex in allRowIndices)
                {
                    // 查找指定行号的DataRow
                    var dataRow = file.DataRows.FirstOrDefault(r => r.RowIndex == rowIndex);
                    if (dataRow == null)
                    {
                        Console.WriteLine($"⚠️ 文件 {file.FileName} 中未找到第{rowIndex}行对应的DataRow");
                        continue;
                    }

                    // 检查该行是否存在缺失数据
                    var hasMissingData = false;
                    var missingColumns = new List<int>();

                    for (int colIndex = 0; colIndex < dataRow.Values.Count && colIndex < 6; colIndex++)
                    {
                        if (!dataRow.Values[colIndex].HasValue)
                        {
                            hasMissingData = true;
                            missingColumns.Add(colIndex);
                        }
                    }

                    // 特殊处理：强制检查第200行
                    if (rowIndex == 200)
                    {
                        Console.WriteLine($"🔍 特殊检查第200行: 数据行名称={dataRow.Name}, 列数={dataRow.Values.Count}");
                        for (int colIndex = 0; colIndex < Math.Min(dataRow.Values.Count, 6); colIndex++)
                        {
                            var value = dataRow.Values[colIndex];
                            var colName = GetColumnName(colIndex);
                            Console.WriteLine($"  第200行{colName}列: {(value.HasValue ? value.Value.ToString("F2") : "空值")}");
                        }
                    }

                    if (hasMissingData)
                    {
                        Console.WriteLine($"⚠️ 发现第{rowIndex}行存在缺失数据，缺失列: [{string.Join(", ", missingColumns.Select(c => GetColumnName(c)))}]，开始补充...");

                        // 第一步：补充D列到I列的缺失值（使用相邻行的平均值）
                        ProcessRowMissingDataByAverage(file, rowIndex, files);

                        // 第二步：处理累计变化量逻辑
                        ProcessRowCumulativeChanges(file, rowIndex, config);
                    }
                    else if (rowIndex == 200)
                    {
                        Console.WriteLine($"ℹ️ 第200行没有检测到缺失数据，但强制处理...");

                        // 强制处理第200行
                        ProcessRowMissingDataByAverage(file, rowIndex, files);
                        ProcessRowCumulativeChanges(file, rowIndex, config);
                    }
                }
            }
        }

        /// <summary>
        /// 使用相邻行平均值补充指定行的缺失数据
        /// </summary>
        /// <param name="file">当前文件</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="files">文件列表（用于获取相邻行数据）</param>
        private static void ProcessRowMissingDataByAverage(ExcelFile file, int rowIndex, List<ExcelFile> files)
        {
            var dataRow = file.DataRows.FirstOrDefault(r => r.RowIndex == rowIndex);
            if (dataRow == null)
            {
                Console.WriteLine($"⚠️ 第{rowIndex}行未找到对应的DataRow");
                return;
            }

            Console.WriteLine($"🔍 处理第{rowIndex}行缺失数据，数据行名称: {dataRow.Name}");

            // 查找相邻行（前一行和后一行）
            var previousRowIndex = rowIndex - 1;
            var nextRowIndex = rowIndex + 1;

            for (int colIndex = 0; colIndex < dataRow.Values.Count && colIndex < 6; colIndex++)
            {
                // 如果该列为空，使用相邻行的平均值
                if (!dataRow.Values[colIndex].HasValue)
                {
                    Console.WriteLine($"🔍 第{rowIndex}行{GetColumnName(colIndex)}列为空，尝试获取相邻行数据...");

                    var valuePrevious = GetValueFromRowAndColumn(files, file, previousRowIndex, colIndex);
                    var valueNext = GetValueFromRowAndColumn(files, file, nextRowIndex, colIndex);

                    Console.WriteLine($"  前一行({previousRowIndex}){GetColumnName(colIndex)}列: {(valuePrevious.HasValue ? valuePrevious.Value.ToString("F2") : "无数据")}");
                    Console.WriteLine($"  后一行({nextRowIndex}){GetColumnName(colIndex)}列: {(valueNext.HasValue ? valueNext.Value.ToString("F2") : "无数据")}");

                    if (valuePrevious.HasValue && valueNext.HasValue)
                    {
                        var averageValue = (valuePrevious.Value + valueNext.Value) / 2.0;
                        dataRow.Values[colIndex] = averageValue;
                        Console.WriteLine($"✅ 补充第{rowIndex}行{GetColumnName(colIndex)}列: {averageValue:F2} (前一行:{valuePrevious:F2} + 后一行:{valueNext:F2})");
                    }
                    else if (valuePrevious.HasValue)
                    {
                        dataRow.Values[colIndex] = valuePrevious.Value;
                        Console.WriteLine($"✅ 补充第{rowIndex}行{GetColumnName(colIndex)}列: {valuePrevious.Value:F2} (使用前一行值)");
                    }
                    else if (valueNext.HasValue)
                    {
                        dataRow.Values[colIndex] = valueNext.Value;
                        Console.WriteLine($"✅ 补充第{rowIndex}行{GetColumnName(colIndex)}列: {valueNext.Value:F2} (使用后一行值)");
                    }
                    else
                    {
                        // 如果相邻行都没有值，尝试使用其他策略
                        Console.WriteLine($"⚠️ 第{rowIndex}行{GetColumnName(colIndex)}列无法获取相邻行值，尝试其他策略...");

                        // 策略1：尝试从其他文件获取相同行的数据
                        var otherFileValue = GetValueFromOtherFiles(files, file, rowIndex, colIndex);
                        if (otherFileValue.HasValue)
                        {
                            dataRow.Values[colIndex] = otherFileValue.Value;
                            Console.WriteLine($"✅ 从其他文件获取第{rowIndex}行{GetColumnName(colIndex)}列值: {otherFileValue.Value:F2}");
                        }
                        else
                        {
                            // 策略2：使用默认值
                            var defaultValue = GetDefaultValueForDataType(dataRow.Name);
                            dataRow.Values[colIndex] = defaultValue;
                            Console.WriteLine($"⚠️ 第{rowIndex}行{GetColumnName(colIndex)}列使用默认值: {defaultValue:F2}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 从其他文件获取指定行和列的值
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="currentFile">当前文件</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="colIndex">列索引</param>
        /// <returns>值</returns>
        private static double? GetValueFromOtherFiles(List<ExcelFile> files, ExcelFile currentFile, int rowIndex, int colIndex)
        {
            var validValues = new List<double>();

            foreach (var file in files)
            {
                if (file == currentFile) continue; // 跳过当前文件

                var dataRow = file.DataRows.FirstOrDefault(r => r.RowIndex == rowIndex);
                if (dataRow != null && colIndex < dataRow.Values.Count && dataRow.Values[colIndex].HasValue)
                {
                    validValues.Add(dataRow.Values[colIndex].Value);
                }
            }

            if (validValues.Any())
            {
                // 返回平均值
                var average = validValues.Average();
                Console.WriteLine($"  从其他文件获取到 {validValues.Count} 个有效值，平均值: {average:F2}");
                return average;
            }

            return null;
        }

        /// <summary>
        /// 处理指定行的累计变化量逻辑
        /// 确保G行 = G(行-1) + D行
        /// </summary>
        /// <param name="file">当前文件</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="config">配置</param>
        private static void ProcessRowCumulativeChanges(ExcelFile file, int rowIndex, DataProcessorConfig config)
        {
            var dataRow = file.DataRows.FirstOrDefault(r => r.RowIndex == rowIndex);
            if (dataRow == null) return;

            // 检查G列（累计变化量）
            for (int colIndex = 0; colIndex < dataRow.Values.Count && colIndex < 6; colIndex++)
            {
                var columnName = GetColumnName(colIndex);

                // 如果是G列（累计变化量），需要计算：G行 = G(行-1) + D行
                if (IsCumulativeColumn(columnName, config))
                {
                    var currentGValue = dataRow.Values[colIndex];
                    var previousGValue = GetValueFromRowAndColumn(new List<ExcelFile> { file }, file, rowIndex - 1, colIndex);

                    // 获取对应的D列索引（变化量）
                    var dColumnIndex = GetDColumnIndexForGColumn(colIndex);
                    var currentDValue = dataRow.Values[dColumnIndex];

                    if (previousGValue.HasValue && currentDValue.HasValue)
                    {
                        // 计算正确的G行值：G行 = G(行-1) + D行
                        var calculatedGValue = previousGValue.Value + currentDValue.Value;

                        // 如果计算值与当前值不同，更新为计算值
                        if (!currentGValue.HasValue || Math.Abs(currentGValue.Value - calculatedGValue) > 0.001)
                        {
                            dataRow.Values[colIndex] = calculatedGValue;
                            Console.WriteLine($"✅ 修正第{rowIndex}行{columnName}列累计变化量: {calculatedGValue:F2} (G{rowIndex - 1}:{previousGValue:F2} + D{rowIndex}:{currentDValue:F2})");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 根据G列索引获取对应的D列索引
        /// </summary>
        /// <param name="gColumnIndex">G列索引</param>
        /// <returns>对应的D列索引</returns>
        private static int GetDColumnIndexForGColumn(int gColumnIndex)
        {
            // 假设G列和D列是一一对应的
            // 如果G列是索引0，对应的D列也是索引0
            return gColumnIndex;
        }

        /// <summary>
        /// 从指定行和列获取值
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="currentFile">当前文件</param>
        /// <param name="rowIndex">行索引（1基）</param>
        /// <param name="colIndex">列索引（0基，对应D-I列）</param>
        /// <returns>值</returns>
        private static double? GetValueFromRowAndColumn(List<ExcelFile> files, ExcelFile currentFile, int rowIndex, int colIndex)
        {
            // 查找指定行号的DataRow
            var dataRow = currentFile.DataRows.FirstOrDefault(r => r.RowIndex == rowIndex);

            if (dataRow == null)
            {
                Console.WriteLine($"    ⚠️ 在文件 {currentFile.FileName} 中未找到第{rowIndex}行对应的DataRow");
                return null;
            }

            if (colIndex >= dataRow.Values.Count)
            {
                Console.WriteLine($"    ⚠️ 第{rowIndex}行{GetColumnName(colIndex)}列索引超出范围 (列数: {dataRow.Values.Count})");
                return null;
            }

            var value = dataRow.Values[colIndex];
            if (value.HasValue)
            {
                Console.WriteLine($"    ✅ 成功获取第{rowIndex}行{GetColumnName(colIndex)}列值: {value.Value:F2}");
                return value.Value;
            }
            else
            {
                Console.WriteLine($"    ⚠️ 第{rowIndex}行{GetColumnName(colIndex)}列值为空");
                return null;
            }
        }

        /// <summary>
        /// 获取列名
        /// </summary>
        /// <param name="colIndex">列索引</param>
        /// <returns>列名</returns>
        private static string GetColumnName(int colIndex)
        {
            return ((char)('A' + colIndex)).ToString();
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
                        TargetFileName = FileNameParser.GenerateFileName(dateCompleteness.Date, missingHour, sourceFile.ProjectName),
                        NeedsAdjustment = true,
                        AdjustmentParams = new AdjustmentParameters
                        {
                            AdjustmentRange = 0.05, // 5%调整范围
                            RandomSeed = sourceFile.GetHashCode() + missingHour, // 基于源文件和目标时间的种子
                            MinimumAdjustment = 0.001,
                            MaintainDataCorrelation = true,
                            CorrelationWeight = 0.7
                        }
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
        /// 创建补充文件并修改A2列数据内容
        /// </summary>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="allFiles">所有文件列表（用于确定上期观测时间）</param>
        /// <returns>创建的文件数量</returns>
        public static int CreateSupplementFilesWithA2Update(List<SupplementFileInfo> supplementFiles, string outputDirectory, List<ExcelFile> allFiles)
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

                    // 修改A2列数据内容
                    UpdateA2CellContent(targetFilePath, supplementFile, allFiles);

                    createdCount++;

                    Console.WriteLine($"✅ 已创建补充文件: {supplementFile.TargetFileName}");
                    Console.WriteLine($"   源文件: {Path.GetFileName(sourceFilePath)}");
                    Console.WriteLine($"   A2列已更新: 本期观测 {supplementFile.TargetDate:yyyy-M-d} {supplementFile.TargetHour:00}:00");
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
        /// 更新Excel文件的A2列内容
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="supplementFile">补充文件信息</param>
        /// <param name="allFiles">所有文件列表</param>
        private static void UpdateA2CellContent(string filePath, SupplementFileInfo supplementFile, List<ExcelFile> allFiles)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".xlsx")
                {
                    UpdateA2CellContentXlsx(filePath, supplementFile, allFiles);
                }
                else if (extension == ".xls")
                {
                    UpdateA2CellContentXls(filePath, supplementFile, allFiles);
                }
                else
                {
                    Console.WriteLine($"⚠️  不支持的文件格式: {extension}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 更新A2列内容失败: {Path.GetFileName(filePath)}");
                Console.WriteLine($"   错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新XLSX文件的A2列内容
        /// </summary>
        private static void UpdateA2CellContentXlsx(string filePath, SupplementFileInfo supplementFile, List<ExcelFile> allFiles)
        {
            using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                throw new InvalidOperationException("Excel文件中没有找到工作表");
            }

            // 确定本期观测时间
            var currentObservationTime = $"{supplementFile.TargetDate:yyyy-M-d} {supplementFile.TargetHour:00}:00";

            // 确定上期观测时间
            var previousObservationTime = GetPreviousObservationTime(supplementFile, allFiles);

            // 更新A2列内容
            var a2Content = $"本期观测：{currentObservationTime} 上期观测：{previousObservationTime}";
            worksheet.Cells["A2"].Value = a2Content;

            package.Save();
        }

        /// <summary>
        /// 更新XLS文件的A2列内容
        /// </summary>
        private static void UpdateA2CellContentXls(string filePath, SupplementFileInfo supplementFile, List<ExcelFile> allFiles)
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);
            var workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
            var worksheet = workbook.GetSheetAt(0);

            // 确定本期观测时间
            var currentObservationTime = $"{supplementFile.TargetDate:yyyy-M-d} {supplementFile.TargetHour:00}:00";

            // 确定上期观测时间
            var previousObservationTime = GetPreviousObservationTime(supplementFile, allFiles);

            // 更新A2列内容
            var a2Content = $"本期观测：{currentObservationTime} 上期观测：{previousObservationTime}";
            var cell = worksheet.GetRow(1)?.GetCell(0) ?? worksheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue(a2Content);

            stream.Position = 0;
            workbook.Write(stream);
        }

        /// <summary>
        /// 获取上期观测时间
        /// </summary>
        /// <param name="supplementFile">补充文件信息</param>
        /// <param name="allFiles">所有文件列表</param>
        /// <returns>上期观测时间字符串</returns>
        public static string GetPreviousObservationTime(SupplementFileInfo supplementFile, List<ExcelFile> allFiles)
        {
            // 按时间顺序排序所有文件
            var sortedFiles = allFiles.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

            // 找到当前文件在排序列表中的位置
            var currentFileIndex = sortedFiles.FindIndex(f =>
                f.Date.Date == supplementFile.TargetDate.Date && f.Hour == supplementFile.TargetHour);

            // 如果找不到当前文件，说明这是一个新创建的文件
            if (currentFileIndex == -1)
            {
                // 找到目标时间点之前的最后一个文件
                var previousFile = sortedFiles
                    .Where(f => f.Date.Date < supplementFile.TargetDate.Date ||
                               (f.Date.Date == supplementFile.TargetDate.Date && f.Hour < supplementFile.TargetHour))
                    .OrderBy(f => f.Date).ThenBy(f => f.Hour)
                    .LastOrDefault();

                if (previousFile != null)
                {
                    return $"{previousFile.Date:yyyy-M-d} {previousFile.Hour:00}:00";
                }
                else
                {
                    // 如果没有找到前一个文件，使用当前时间作为上期观测时间
                    return $"{supplementFile.TargetDate:yyyy-M-d} {supplementFile.TargetHour:00}:00";
                }
            }
            else
            {
                // 如果找到了当前文件，获取前一个文件
                if (currentFileIndex > 0)
                {
                    var previousFile = sortedFiles[currentFileIndex - 1];
                    return $"{previousFile.Date:yyyy-M-d} {previousFile.Hour:00}:00";
                }
                else
                {
                    // 如果是第一个文件，使用当前时间作为上期观测时间
                    return $"{supplementFile.TargetDate:yyyy-M-d} {supplementFile.TargetHour:00}:00";
                }
            }
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

        /// <summary>
        /// 获取所有需要处理的文件（包括原始文件和补充文件）
        /// </summary>
        /// <param name="originalFiles">原始文件列表</param>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>所有需要处理的文件列表</returns>
        public static List<ExcelFile> GetAllFilesForProcessing(List<ExcelFile> originalFiles, List<SupplementFileInfo> supplementFiles, string outputDirectory)
        {
            var allFiles = new List<ExcelFile>(originalFiles);

            // 为补充文件创建ExcelFile对象
            foreach (var supplementFile in supplementFiles)
            {
                var supplementFilePath = Path.Combine(outputDirectory, supplementFile.TargetFileName);

                if (File.Exists(supplementFilePath))
                {
                    // 创建补充文件的ExcelFile对象
                    var supplementExcelFile = new ExcelFile
                    {
                        FilePath = supplementFilePath,
                        FileName = supplementFile.TargetFileName,
                        Date = supplementFile.TargetDate,
                        Hour = supplementFile.TargetHour,
                        ProjectName = supplementFile.ProjectName,
                        FileSize = new FileInfo(supplementFilePath).Length,
                        LastModified = new FileInfo(supplementFilePath).LastWriteTime,
                        IsValid = true,
                        IsSupplementFile = true,
                        SupplementSource = supplementFile
                    };

                    // 读取补充文件的数据
                    try
                    {
                        var excelService = new ExcelService();
                        var supplementFileWithData = excelService.ReadExcelFile(supplementFilePath);
                        supplementExcelFile.DataRows = supplementFileWithData.DataRows;
                        supplementExcelFile.IsValid = supplementFileWithData.IsValid;
                        supplementExcelFile.IsLocked = supplementFileWithData.IsLocked;

                        allFiles.Add(supplementExcelFile);
                        Console.WriteLine($"✅ 已加载补充文件数据: {supplementFile.TargetFileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ 读取补充文件失败: {supplementFile.TargetFileName} - {ex.Message}");
                    }
                }
            }

            // 按时间顺序排序
            allFiles.Sort((a, b) =>
            {
                var dateComparison = a.Date.CompareTo(b.Date);
                if (dateComparison != 0)
                    return dateComparison;
                return a.Hour.CompareTo(b.Hour);
            });

            Console.WriteLine($"📊 准备处理 {allFiles.Count} 个文件（原始文件: {originalFiles.Count}, 补充文件: {supplementFiles.Count}）");
            return allFiles;
        }

        /// <summary>
        /// 为所有文件更新A2列数据内容
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>更新的文件数量</returns>
        public static int UpdateA2ColumnForAllFiles(List<ExcelFile> files, string outputDirectory)
        {
            if (files == null || !files.Any())
            {
                return 0;
            }

            int updatedCount = 0;
            var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

            Console.WriteLine($"📝 开始更新A2列内容，共 {files.Count} 个文件...");

            for (int i = 0; i < sortedFiles.Count; i++)
            {
                var currentFile = sortedFiles[i];
                var filePath = Path.Combine(outputDirectory, currentFile.FileName);

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"⚠️  文件不存在，跳过A2列更新: {currentFile.FileName}");
                    continue;
                }

                try
                {
                    // 确定本期观测时间
                    var currentObservationTime = $"{currentFile.Date:yyyy-M-d} {currentFile.Hour:00}:00";

                    // 确定上期观测时间
                    string previousObservationTime;
                    if (i > 0)
                    {
                        var previousFile = sortedFiles[i - 1];
                        previousObservationTime = $"{previousFile.Date:yyyy-M-d} {previousFile.Hour:00}:00";
                    }
                    else
                    {
                        // 如果是第一个文件，使用当前时间作为上期观测时间
                        previousObservationTime = currentObservationTime;
                    }

                    // 更新A2列内容
                    UpdateA2CellContentForFile(filePath, currentObservationTime, previousObservationTime);

                    updatedCount++;

                    if (updatedCount % 10 == 0)
                    {
                        Console.WriteLine($"📈 A2列更新进度: {updatedCount}/{files.Count}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 更新A2列失败: {currentFile.FileName} - {ex.Message}");
                }
            }

            return updatedCount;
        }

        /// <summary>
        /// 为单个文件更新A2列内容
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="currentObservationTime">本期观测时间</param>
        /// <param name="previousObservationTime">上期观测时间</param>
        private static void UpdateA2CellContentForFile(string filePath, string currentObservationTime, string previousObservationTime)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                var a2Content = $"本期观测：{currentObservationTime} 上期观测：{previousObservationTime}";

                if (extension == ".xlsx")
                {
                    UpdateA2CellContentXlsxForFile(filePath, a2Content);
                }
                else if (extension == ".xls")
                {
                    UpdateA2CellContentXlsForFile(filePath, a2Content);
                }
                else
                {
                    Console.WriteLine($"⚠️  不支持的文件格式: {extension}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 更新A2列内容失败: {Path.GetFileName(filePath)}");
                Console.WriteLine($"   错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新XLSX文件的A2列内容
        /// </summary>
        private static void UpdateA2CellContentXlsxForFile(string filePath, string a2Content)
        {
            using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                throw new InvalidOperationException("Excel文件中没有找到工作表");
            }

            worksheet.Cells["A2"].Value = a2Content;
            package.Save();
        }

        /// <summary>
        /// 更新XLS文件的A2列内容
        /// </summary>
        private static void UpdateA2CellContentXlsForFile(string filePath, string a2Content)
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);
            var workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
            var worksheet = workbook.GetSheetAt(0);

            var cell = worksheet.GetRow(1)?.GetCell(0) ?? worksheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue(a2Content);

            stream.Position = 0;
            workbook.Write(stream);
        }

        /// <summary>
        /// 处理补充文件数据微调（优化版本）
        /// </summary>
        private static (int Adjustments, int CacheHits, int CacheMisses) ProcessSupplementFilesAdjustmentOptimized(
            List<ExcelFile> sortedFiles, DataProcessorConfig config, DataCache? cache, PerformanceMetrics metrics)
        {
            var supplementFiles = sortedFiles.Where(f => f.IsSupplementFile).ToList();
            var adjustments = 0;
            var cacheHits = 0;
            var cacheMisses = 0;

            if (!supplementFiles.Any())
            {
                Console.WriteLine("✅ 未发现需要微调的补充文件");
                return (0, 0, 0);
            }

            Console.WriteLine($"📊 发现 {supplementFiles.Count} 个需要微调的补充文件");

            foreach (var supplementFile in supplementFiles)
            {
                Console.WriteLine($"🔄 微调补充文件: {supplementFile.FileName}");
                var (adj, hits, misses) = AdjustSupplementFileDataOptimized(supplementFile, config, cache);
                adjustments += adj;
                cacheHits += hits;
                cacheMisses += misses;
            }

            // 更新性能指标
            metrics.SupplementFileAdjustments = adjustments;
            metrics.CacheHits += cacheHits;
            metrics.CacheMisses += cacheMisses;

            return (adjustments, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 批量处理文件（优化版本）
        /// </summary>
        private static void ProcessBatchOptimized(List<ExcelFile> batch, List<ExcelFile> allFiles,
            Dictionary<string, Dictionary<int, List<double>>> valueCache, DataCache? cache,
            DataProcessorConfig config, PerformanceMetrics metrics)
        {
            foreach (var file in batch)
            {
                var fileIndex = allFiles.IndexOf(file);
                ProcessFileMissingDataOptimized(file, allFiles, fileIndex, valueCache, cache, config, metrics);
            }
        }

        /// <summary>
        /// 预处理值缓存（优化版本）
        /// </summary>
        private static void PreprocessValueCacheOptimized(List<ExcelFile> files,
            Dictionary<string, Dictionary<int, List<double>>> valueCache, DataCache? cache,
            DataProcessorConfig config)
        {
            if (cache == null)
            {
                // 如果没有缓存，使用原来的方法
                PreprocessValueCache(files, valueCache);
                return;
            }

            foreach (var file in files)
            {
                foreach (var dataRow in file.DataRows)
                {
                    var cacheKey = $"datarow_{dataRow.Name}_{file.FileName}";

                    // 尝试从缓存获取
                    var cachedValues = cache.Get<List<double>>(cacheKey);
                    if (cachedValues != null)
                    {
                        // 缓存命中
                        if (!valueCache.ContainsKey(dataRow.Name))
                        {
                            valueCache[dataRow.Name] = new Dictionary<int, List<double>>();
                        }

                        for (int valueIndex = 0; valueIndex < cachedValues.Count; valueIndex++)
                        {
                            if (!valueCache[dataRow.Name].ContainsKey(valueIndex))
                            {
                                valueCache[dataRow.Name][valueIndex] = new List<double>();
                            }
                            valueCache[dataRow.Name][valueIndex].Add(cachedValues[valueIndex]);
                        }
                        continue;
                    }

                    // 缓存未命中，计算并缓存
                    var values = new List<double>();
                    for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                    {
                        if (dataRow.Values[valueIndex].HasValue)
                        {
                            values.Add(dataRow.Values[valueIndex].Value);

                            if (!valueCache.ContainsKey(dataRow.Name))
                            {
                                valueCache[dataRow.Name] = new Dictionary<int, List<double>>();
                            }

                            if (!valueCache[dataRow.Name].ContainsKey(valueIndex))
                            {
                                valueCache[dataRow.Name][valueIndex] = new List<double>();
                            }

                            valueCache[dataRow.Name][valueIndex].Add(dataRow.Values[valueIndex].Value);
                        }
                    }

                    // 缓存计算结果
                    cache.Set(cacheKey, values);
                }
            }
        }

        /// <summary>
        /// 处理单个文件缺失数据（优化版本）
        /// </summary>
        private static void ProcessFileMissingDataOptimized(ExcelFile currentFile, List<ExcelFile> allFiles, int fileIndex,
            Dictionary<string, Dictionary<int, List<double>>> valueCache, DataCache? cache,
            DataProcessorConfig config, PerformanceMetrics metrics)
        {
            // 使用缓存优化文件处理
            var cacheKey = $"file_processing_{currentFile.FileName}";

            if (cache != null)
            {
                var cachedResult = cache.Get<bool>(cacheKey);
                if (cachedResult)
                {
                    // 文件已处理过，跳过
                    return;
                }
            }

            // 处理文件缺失数据
            foreach (var dataRow in currentFile.DataRows)
            {
                for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
                {
                    if (!dataRow.Values[valueIndex].HasValue)
                    {
                        // 使用缓存的优化计算补充值
                        var supplementValue = CalculateSupplementValueOptimized(dataRow.Name, valueIndex, allFiles, fileIndex, valueCache);
                        if (supplementValue.HasValue)
                        {
                            dataRow.Values[valueIndex] = supplementValue.Value;
                        }
                    }
                }
            }

            // 缓存处理结果
            cache?.Set(cacheKey, true);
        }

        /// <summary>
        /// 处理累计变化量行（优化版本）
        /// </summary>
        private static (int Calculations, int CacheHits, int CacheMisses) ProcessCumulativeRowOptimized(
            DataRow cumulativeRow, ExcelFile previousFile, ExcelFile currentFile,
            DataProcessorConfig config, Random random, DataCache? cache)
        {
            var calculations = 0;
            var cacheHits = 0;
            var cacheMisses = 0;

            var baseColumnName = GetBaseColumnForCumulative(cumulativeRow.Name, config);

            for (int valueIndex = 0; valueIndex < cumulativeRow.Values.Count; valueIndex++)
            {
                if (!cumulativeRow.Values[valueIndex].HasValue)
                {
                    var cacheKey = $"cumulative_{cumulativeRow.Name}_{valueIndex}_{previousFile.FileName}_{currentFile.FileName}";

                    if (cache != null)
                    {
                        var cachedValue = cache.Get<double?>(cacheKey);
                        if (cachedValue.HasValue)
                        {
                            cumulativeRow.Values[valueIndex] = cachedValue.Value;
                            cacheHits++;
                            continue;
                        }
                    }

                    var previousCumulative = GetPreviousCumulativeValue(previousFile, cumulativeRow.Name, valueIndex);
                    var currentChange = GetCurrentChangeValue(currentFile, baseColumnName, valueIndex);

                    if (previousCumulative.HasValue && currentChange.HasValue)
                    {
                        // 计算累计量：本期累计 = 上期累计 + 本期变化
                        var newCumulative = previousCumulative.Value + currentChange.Value;
                        cumulativeRow.Values[valueIndex] = newCumulative;

                        // 缓存结果
                        cache?.Set(cacheKey, newCumulative);
                        calculations++;
                    }
                    else
                    {
                        cacheMisses++;
                    }
                }
            }

            return (calculations, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 处理缺失时间段（优化版本）
        /// </summary>
        private static (int Processings, int CacheHits, int CacheMisses) ProcessMissingPeriodOptimized(
            MissingPeriod period, List<ExcelFile> files, DataProcessorConfig config, DataCache? cache)
        {
            var processings = 0;
            var cacheHits = 0;
            var cacheMisses = 0;
            var totalTimePoints = period.MissingTimes.Count;
            var totalDataRows = period.MissingDataRows.Count;
            var lastProgressTime = DateTime.Now;

            Console.WriteLine($"  📅 处理时间段: {period.StartTime:yyyy-MM-dd HH:mm} 到 {period.EndTime:yyyy-MM-dd HH:mm}, 共 {totalTimePoints} 个时间点, {totalDataRows} 个数据行");

            // 按时间顺序处理，确保前面的数据补充能影响后面的计算
            for (int timeIndex = 0; timeIndex < period.MissingTimes.Count; timeIndex++)
            {
                var missingTime = period.MissingTimes[timeIndex];
                var currentTime = DateTime.Now;

                // 每处理5个时间点或每20秒显示一次进度
                if ((timeIndex + 1) % 5 == 0 || (currentTime - lastProgressTime).TotalSeconds >= 20)
                {
                    var timeProgress = (double)(timeIndex + 1) / totalTimePoints * 100;
                    Console.WriteLine($"    ⏰ 时间点进度: {timeIndex + 1}/{totalTimePoints} ({timeProgress:F1}%) - 当前处理: {missingTime:yyyy-MM-dd HH:mm}");
                    lastProgressTime = currentTime;
                }

                // 找到当前时间点的文件
                var targetFile = files.FirstOrDefault(f =>
                    f.Date.Date == missingTime.Date && f.Hour == missingTime.Hour);

                if (targetFile == null) continue;

                // 处理当前时间点的所有数据行
                foreach (var dataRowName in period.MissingDataRows)
                {
                    var dataRow = targetFile.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                    if (dataRow == null) continue;

                    // 处理数据行中的每个缺失值（只处理前一半的列）
                    for (int valueIndex = 0; valueIndex < dataRow.Values.Count / 2; valueIndex++)
                    {
                        if (dataRow.Values[valueIndex].HasValue) continue; // 跳过已有值

                        var cacheKey = $"missing_period_{dataRowName}_{valueIndex}_{missingTime:yyyyMMddHH}";

                        if (cache != null)
                        {
                            var cachedValue = cache.Get<double?>(cacheKey);
                            if (cachedValue.HasValue)
                            {
                                dataRow.Values[valueIndex] = cachedValue.Value;
                                cacheHits++;
                                continue;
                            }
                        }

                        // 获取对应列的前后有效值
                        var (previousValue, nextValue) = GetNearestValuesForTimePoint(
                            files, dataRowName, missingTime, valueIndex);

                        // 获取上一期的数据（使用最新的补充数据）
                        var cumulativeColumnIndex = valueIndex + (dataRow.Values.Count / 2);
                        var previousPeriodValue = GetPreviousPeriodData(files, dataRowName, missingTime, cumulativeColumnIndex);

                        if (previousValue.HasValue && nextValue.HasValue)
                        {
                            var missingPoint = new MissingDataPoint
                            {
                                DataRowName = dataRowName,
                                ValueIndex = valueIndex,
                                TimePoint = missingTime,
                                PreviousValue = previousValue,
                                NextValue = nextValue,
                                BaseValue = (previousValue.Value + nextValue.Value) / 2
                            };

                            var adjustedValue = CalculateAdjustedValueForMissingPoint(missingPoint, period, config);

                            // 应用调整后的值
                            dataRow.Values[valueIndex] = adjustedValue;

                            // 计算累计变化量（如果有上一期数据）
                            if (previousPeriodValue.HasValue)
                            {
                                if (cumulativeColumnIndex < dataRow.Values.Count)
                                {
                                    var newCumulativeValue = previousPeriodValue.Value + adjustedValue;
                                    dataRow.Values[cumulativeColumnIndex] = newCumulativeValue;
                                    //Console.WriteLine($"    📊 更新累计值: {dataRowName} 第{cumulativeColumnIndex + 1}列 = {previousPeriodValue:F2} + {adjustedValue:F2} = {newCumulativeValue:F2}");
                                }
                            }

                            // 缓存结果
                            cache?.Set(cacheKey, adjustedValue);
                            processings++;
                        }
                        else
                        {
                            cacheMisses++;
                        }
                    }
                }
            }

            Console.WriteLine($"    ✅ 时间段处理完成: 补充了 {processings} 个缺失值, 缓存命中 {cacheHits} 次, 缓存未命中 {cacheMisses} 次");
            return (processings, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 微调补充文件数据（优化版本）
        /// </summary>
        private static (int Adjustments, int CacheHits, int CacheMisses) AdjustSupplementFileDataOptimized(
            ExcelFile supplementFile, DataProcessorConfig config, DataCache? cache)
        {
            var adjustments = 0;
            var cacheHits = 0;
            var cacheMisses = 0;

            if (supplementFile.SupplementSource == null)
            {
                Console.WriteLine($"⚠️ 补充文件 {supplementFile.FileName} 缺少源文件信息，跳过微调");
                return (0, 0, 0);
            }

            var adjustmentParams = supplementFile.SupplementSource.AdjustmentParams;
            var random = new Random(adjustmentParams.RandomSeed);

            Console.WriteLine($"🔧 使用调整参数: 范围={adjustmentParams.AdjustmentRange:P0}, 种子={adjustmentParams.RandomSeed}");

            foreach (var dataRow in supplementFile.DataRows)
            {
                var (adj, hits, misses) = AdjustDataRowValuesOptimized(dataRow, adjustmentParams, random, config, cache);
                adjustments += adj;
                cacheHits += hits;
                cacheMisses += misses;
            }

            return (adjustments, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 微调数据行的值（优化版本）
        /// </summary>
        private static (int Adjustments, int CacheHits, int CacheMisses) AdjustDataRowValuesOptimized(
            DataRow dataRow, AdjustmentParameters adjustmentParams, Random random,
            DataProcessorConfig config, DataCache? cache)
        {
            var adjustments = 0;
            var cacheHits = 0;
            var cacheMisses = 0;

            for (int i = 0; i < dataRow.Values.Count; i++)
            {
                if (dataRow.Values[i].HasValue)
                {
                    var cacheKey = $"adjustment_{dataRow.Name}_{i}_{adjustmentParams.RandomSeed}";

                    if (cache != null)
                    {
                        var cachedValue = cache.Get<double?>(cacheKey);
                        if (cachedValue.HasValue)
                        {
                            dataRow.Values[i] = cachedValue.Value;
                            cacheHits++;
                            continue;
                        }
                    }

                    var originalValue = dataRow.Values[i].Value;
                    var adjustedValue = CalculateSupplementAdjustment(originalValue, adjustmentParams, random, config);
                    dataRow.Values[i] = adjustedValue;

                    // 缓存结果
                    cache?.Set(cacheKey, adjustedValue);
                    adjustments++;
                }
            }

            return (adjustments, cacheHits, cacheMisses);
        }

        /// <summary>
        /// 显示性能指标
        /// </summary>
        private static void DisplayPerformanceMetrics(PerformanceMetrics metrics)
        {
            Console.WriteLine("\n📊 ========== 性能统计 ==========");
            Console.WriteLine($"⏱️ 总处理时间: {metrics.TotalProcessingTime.TotalMilliseconds:F2}ms");
            Console.WriteLine($"📁 处理文件数: {metrics.TotalFilesProcessed}");
            Console.WriteLine($"📊 处理数据行数: {metrics.TotalDataRowsProcessed}");
            Console.WriteLine($"🔄 累计变化量计算: {metrics.CumulativeCalculations}");
            Console.WriteLine($"🔍 连续缺失数据处理: {metrics.MissingDataProcessings}");
            Console.WriteLine($"🔧 补充文件微调: {metrics.SupplementFileAdjustments}");
            Console.WriteLine($"⚡ 平均每文件处理时间: {metrics.AverageTimePerFile.TotalMilliseconds:F2}ms");

            Console.WriteLine("================================\n");
        }

        /// <summary>
        /// 执行最终数据完整性检查，确保所有缺失值都被补充
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>补充的缺失值数量</returns>
        private static int PerformFinalDataIntegrityCheck(List<ExcelFile> files)
        {
            var supplementedCount = 0;

            foreach (var file in files)
            {
                foreach (var dataRow in file.DataRows)
                {
                    for (int i = 0; i < dataRow.Values.Count; i++)
                    {
                        if (!dataRow.Values[i].HasValue)
                        {
                            // 使用默认值补充
                            var defaultValue = GetDefaultValueForDataType(dataRow.Name);
                            dataRow.Values[i] = defaultValue;
                            supplementedCount++;
                        }
                    }
                }
            }

            return supplementedCount;
        }

        /// <summary>
        /// 验证数据完整性，检查是否存在无效值（NaN、Infinity等）
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <exception cref="InvalidOperationException">当发现无效数据时抛出</exception>
        private static void ValidateDataIntegrity(List<ExcelFile> files)
        {
            foreach (var file in files)
            {
                foreach (var dataRow in file.DataRows)
                {
                    for (int i = 0; i < dataRow.Values.Count; i++)
                    {
                        if (dataRow.Values[i].HasValue)
                        {
                            var value = dataRow.Values[i].Value;
                            if (double.IsNaN(value) || double.IsInfinity(value))
                            {
                                throw new InvalidOperationException($"Invalid data detected in file {file.FileName}, row {dataRow.Name}, column {i}: {value}");
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获取指定时间点对应列的前后有效值
        /// </summary>
        private static (double? PreviousValue, double? NextValue) GetNearestValuesForTimePoint(
            List<ExcelFile> files,
            string dataRowName,
            DateTime targetTime,
            int valueIndex)
        {
            // 找到目标时间在文件列表中的位置
            var targetIndex = files.FindIndex(f =>
                f.Date.Date == targetTime.Date && f.Hour == targetTime.Hour);

            if (targetIndex == -1) return (null, null);

            // 向前搜索有效值（对应列）
            double? previousValue = null;
            for (int i = targetIndex - 1; i >= 0; i--)
            {
                var file = files[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                if (dataRow != null && valueIndex < dataRow.Values.Count &&
                    dataRow.Values[valueIndex].HasValue)
                {
                    previousValue = dataRow.Values[valueIndex].Value;
                    break;
                }
            }

            // 向后搜索有效值（对应列）
            double? nextValue = null;
            for (int i = targetIndex + 1; i < files.Count; i++)
            {
                var file = files[i];
                var dataRow = file.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                if (dataRow != null && valueIndex < dataRow.Values.Count &&
                    dataRow.Values[valueIndex].HasValue)
                {
                    nextValue = dataRow.Values[valueIndex].Value;
                    break;
                }
            }

            return (previousValue, nextValue);
        }

        /// <summary>
        /// 比较原始目录和已处理目录中文件的数值差异
        /// 只比较原始文件每行有值的数据
        /// </summary>
        /// <param name="originalDirectory">原始文件目录路径</param>
        /// <param name="processedDirectory">已处理文件目录路径</param>
        /// <param name="config">配置参数</param>
        /// <returns>比较结果统计</returns>
        public static ComparisonResult CompareOriginalAndProcessedFiles(string originalDirectory, string processedDirectory, DataProcessorConfig? config = null)
        {
            config ??= DataProcessorConfig.Default;

            Console.WriteLine("🔍 开始比较原始文件和已处理文件的数值差异...");
            Console.WriteLine($"📁 原始目录: {originalDirectory}");
            Console.WriteLine($"📁 已处理目录: {processedDirectory}");

            var result = new ComparisonResult();
            var excelService = new ExcelService();

            try
            {
                // 获取原始目录中的文件
                var originalFiles = Directory.GetFiles(originalDirectory, "*.xls")
                    .Where(f => !f.Contains("processed")) // 排除processed子目录
                    .OrderBy(f => Path.GetFileName(f))
                    .ToList();

                Console.WriteLine($"📊 原始目录找到 {originalFiles.Count} 个.xls文件");

                foreach (var originalFilePath in originalFiles)
                {
                    var fileName = Path.GetFileName(originalFilePath);
                    var processedFilePath = Path.Combine(processedDirectory, fileName);

                    // 检查对应的已处理文件是否存在
                    if (!File.Exists(processedFilePath))
                    {
                        Console.WriteLine($"⚠️ 未找到对应的已处理文件: {fileName}");
                        result.MissingProcessedFiles.Add(fileName);
                        continue;
                    }

                    try
                    {
                        // 加载原始文件和已处理文件
                        var originalFile = excelService.ReadExcelFile(originalFilePath);
                        var processedFile = excelService.ReadExcelFile(processedFilePath);

                        if (originalFile == null || processedFile == null)
                        {
                            Console.WriteLine($"❌ 文件加载失败: {fileName}");
                            result.FailedComparisons.Add(fileName);
                            continue;
                        }

                        // 比较文件内容
                        var fileComparison = CompareFileContent(originalFile, processedFile, fileName, config);
                        result.FileComparisons.Add(fileComparison);

                        // 累计统计
                        result.TotalOriginalValues += fileComparison.OriginalValuesCount;
                        result.TotalProcessedValues += fileComparison.ProcessedValuesCount;
                        result.TotalDifferences += fileComparison.DifferencesCount;
                        result.TotalSignificantDifferences += fileComparison.SignificantDifferencesCount;

                        Console.WriteLine($"✅ 完成比较: {fileName} - 差异: {fileComparison.DifferencesCount}/{fileComparison.OriginalValuesCount}");

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ 比较文件失败 {fileName}: {ex.Message}");
                        result.FailedComparisons.Add(fileName);
                    }
                }

                // 输出总结
                Console.WriteLine($"\\n📊 比较完成总结:");
                Console.WriteLine($"   - 原始文件总数: {originalFiles.Count}");
                Console.WriteLine($"   - 成功比较文件数: {result.FileComparisons.Count}");
                Console.WriteLine($"   - 缺失已处理文件数: {result.MissingProcessedFiles.Count}");
                Console.WriteLine($"   - 比较失败文件数: {result.FailedComparisons.Count}");
                Console.WriteLine($"   - 原始数据值总数: {result.TotalOriginalValues}");
                Console.WriteLine($"   - 已处理数据值总数: {result.TotalProcessedValues}");
                Console.WriteLine($"   - 数值差异总数: {result.TotalDifferences}");
                Console.WriteLine($"   - 显著差异总数: {result.TotalSignificantDifferences}");

                if (result.TotalOriginalValues > 0)
                {
                    var differencePercentage = (double)result.TotalDifferences / result.TotalOriginalValues * 100;
                    var significantDifferencePercentage = (double)result.TotalSignificantDifferences / result.TotalOriginalValues * 100;
                    Console.WriteLine($"   - 差异率: {differencePercentage:F2}%");
                    Console.WriteLine($"   - 显著差异率: {significantDifferencePercentage:F2}%");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 比较过程发生错误: {ex.Message}");
                result.HasError = true;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 比较单个文件的内容
        /// </summary>
        private static FileComparisonResult CompareFileContent(ExcelFile originalFile, ExcelFile processedFile, string fileName, DataProcessorConfig config)
        {
            var result = new FileComparisonResult
            {
                FileName = fileName,
                OriginalDate = originalFile.Date,
                ProcessedDate = processedFile.Date
            };

            // 按行名匹配数据行
            var originalRows = originalFile.DataRows.ToDictionary(r => r.Name, r => r);
            var processedRows = processedFile.DataRows.ToDictionary(r => r.Name, r => r);

            foreach (var originalRow in originalFile.DataRows)
            {
                if (!processedRows.ContainsKey(originalRow.Name))
                {
                    result.MissingProcessedRows.Add(originalRow.Name);
                    continue;
                }

                var processedRow = processedRows[originalRow.Name];
                var rowComparison = CompareRowContent(originalRow, processedRow, config);
                result.RowComparisons.Add(rowComparison);

                // 累计统计
                result.OriginalValuesCount += rowComparison.OriginalValuesCount;
                result.ProcessedValuesCount += rowComparison.ProcessedValuesCount;
                result.DifferencesCount += rowComparison.DifferencesCount;
                result.SignificantDifferencesCount += rowComparison.SignificantDifferencesCount;
            }

            return result;
        }

        /// <summary>
        /// 比较单个数据行的内容
        /// </summary>
        private static RowComparisonResult CompareRowContent(DataRow originalRow, DataRow processedRow, DataProcessorConfig config)
        {
            var result = new RowComparisonResult
            {
                RowName = originalRow.Name
            };

            // 比较每一列的值
            var maxColumns = Math.Max(originalRow.Values.Count, processedRow.Values.Count);

            for (int i = 0; i < maxColumns; i++)
            {
                var originalValue = i < originalRow.Values.Count ? originalRow.Values[i] : null;
                var processedValue = i < processedRow.Values.Count ? processedRow.Values[i] : null;

                // 只比较原始文件有值的数据
                if (originalValue.HasValue)
                {
                    result.OriginalValuesCount++;

                    if (processedValue.HasValue)
                    {
                        result.ProcessedValuesCount++;

                        // 计算差异
                        var difference = Math.Abs(processedValue.Value - originalValue.Value);
                        var isSignificant = difference > config.ColumnValidationTolerance;

                        if (difference > 0)
                        {
                            result.DifferencesCount++;

                            if (isSignificant)
                            {
                                result.SignificantDifferencesCount++;
                            }

                            // 记录列差异详情
                            result.ColumnDifferences.Add(new ColumnDifference
                            {
                                ColumnIndex = i,
                                OriginalValue = originalValue.Value,
                                ProcessedValue = processedValue.Value,
                                Difference = difference,
                                IsSignificant = isSignificant
                            });
                        }
                    }
                    else
                    {
                        result.MissingProcessedValues++;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 获取指定数据行名称的上一期数据
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="dataRowName">数据行名称</param>
        /// <param name="currentTime">当前时间</param>
        /// <param name="valueIndex">列索引</param>
        /// <returns>上一期数据，如果没有找到则返回null</returns>
        private static double? GetPreviousPeriodData(List<ExcelFile> files, string dataRowName, DateTime currentTime, int valueIndex)
        {
            try
            {
                // 按时间顺序排序文件
                var sortedFiles = files.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

                // 找到当前时间在文件列表中的位置
                var currentIndex = sortedFiles.FindIndex(f =>
                    f.Date.Date == currentTime.Date && f.Hour == currentTime.Hour);

                if (currentIndex <= 0) return null; // 第一个文件或未找到当前时间

                // 向前搜索上一期的数据
                for (int i = currentIndex - 1; i >= 0; i--)
                {
                    var previousFile = sortedFiles[i];
                    var dataRow = previousFile.DataRows.FirstOrDefault(r => r.Name == dataRowName);
                    
                    if (dataRow != null && valueIndex < dataRow.Values.Count && dataRow.Values[valueIndex].HasValue)
                    {
                        var previousValue = dataRow.Values[valueIndex].Value;
                        //Console.WriteLine($"    📊 找到上一期数据: {dataRowName} 在 {previousFile.Date:yyyy-MM-dd} {previousFile.Hour:D2}:00, 第{valueIndex + 1}列值: {previousValue:F2}");
                        return previousValue;
                    }
                }

                //Console.WriteLine($"    ⚠️ 未找到 {dataRowName} 的上一期数据");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    ❌ 获取上一期数据失败: {ex.Message}");
                return null;
            }
        }
    }
}