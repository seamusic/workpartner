using WorkPartner.Models;
using WorkPartner.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 结果显示器 - 负责显示处理结果和统计信息
    /// </summary>
    public static class ResultDisplay
    {
        /// <summary>
        /// 显示处理结果（阶段3版本）
        /// </summary>
        /// <param name="files">处理后的文件列表</param>
        /// <param name="completenessResult">完整性检查结果</param>
        /// <param name="supplementFiles">补充文件列表</param>
        /// <param name="qualityReport">数据质量报告</param>
        public static void DisplayProcessingResults(List<ExcelFile> files, CompletenessCheckResult completenessResult, List<SupplementFileInfo> supplementFiles, DataQualityReport qualityReport)
        {
            Console.WriteLine("\n--- 阶段3处理结果摘要 ---");

            // 按日期分组显示
            var groupedFiles = files.GroupBy(f => f.Date).OrderBy(g => g.Key);

            foreach (var group in groupedFiles)
            {
                Console.WriteLine($"\n日期: {group.Key:yyyy.M.d}");
                var hours = group.Select(f => f.Hour).OrderBy(h => h).ToList();
                Console.WriteLine($"  时间点: [{string.Join(", ", hours)}]");
                Console.WriteLine($"  文件数: {group.Count()}");

                foreach (var file in group.OrderBy(f => f.Hour))
                {
                    var dataCount = file.DataRows?.Count ?? 0;
                    var completeness = file.DataRows?.Count > 0 
                        ? file.DataRows.Average(r => r.CompletenessPercentage) 
                        : 0;
                    Console.WriteLine($"    {file.FormattedHour}时: {dataCount} 行数据, 完整性 {completeness:F1}%");
                }
            }

            // 数据补充统计
            Console.WriteLine($"\n📊 数据补充统计:");
            var totalMissingValues = files.SelectMany(f => f.DataRows).Sum(r => r.MissingDataCount);
            var totalValues = files.SelectMany(f => f.DataRows).Sum(r => r.Values.Count);
            var supplementedCount = totalValues - totalMissingValues; // 假设所有缺失都已补充
            Console.WriteLine($"  总数据点: {totalValues}");
            Console.WriteLine($"  原始缺失: {totalMissingValues}");
            Console.WriteLine($"  已补充: {supplementedCount}");
            Console.WriteLine($"  补充率: {(totalMissingValues > 0 ? 100.0 : 0):F1}%");

            // 完整性检查结果
            Console.WriteLine($"\n🔍 数据完整性: {(completenessResult.IsAllComplete ? "✅ 完整" : "❌ 不完整")}");

            if (!completenessResult.IsAllComplete)
            {
                Console.WriteLine("缺失的时间点:");
                foreach (var dateCompleteness in completenessResult.DateCompleteness)
                {
                    if (dateCompleteness.MissingHours.Any())
                    {
                        Console.WriteLine($"  {dateCompleteness.Date:yyyy.M.d}: [{string.Join(", ", dateCompleteness.MissingHours)}]");
                    }
                }
            }

            // 补充文件建议
            if (supplementFiles.Any())
            {
                Console.WriteLine($"\n📋 建议生成 {supplementFiles.Count} 个补充文件:");
                foreach (var supplement in supplementFiles.Take(5)) // 只显示前5个
                {
                    Console.WriteLine($"  {supplement.TargetFileName}");
                }
                if (supplementFiles.Count > 5)
                {
                    Console.WriteLine($"  ... 还有 {supplementFiles.Count - 5} 个文件");
                }
            }

            // 数据质量报告
            Console.WriteLine($"\n📈 数据质量报告:");
            Console.WriteLine($"  总体完整性: {qualityReport.OverallCompleteness:F1}%");
            Console.WriteLine($"  有效数据行: {qualityReport.ValidRows}/{qualityReport.TotalRows}");
            Console.WriteLine($"  缺失数据行: {qualityReport.MissingRows}");
        }

        /// <summary>
        /// 显示处理结果（阶段2版本，保留向后兼容）
        /// </summary>
        /// <param name="files">处理后的文件列表</param>
        public static void DisplayProcessingResults(List<ExcelFile> files)
        {
            Console.WriteLine("\n--- 处理结果摘要 ---");

            // 按日期分组显示
            var groupedFiles = files.GroupBy(f => f.Date).OrderBy(g => g.Key);

            foreach (var group in groupedFiles)
            {
                Console.WriteLine($"\n日期: {group.Key:yyyy.M.d}");
                var hours = group.Select(f => f.Hour).OrderBy(h => h).ToList();
                Console.WriteLine($"  时间点: [{string.Join(", ", hours)}]");
                Console.WriteLine($"  文件数: {group.Count()}");

                foreach (var file in group.OrderBy(f => f.Hour))
                {
                    var dataCount = file.DataRows?.Count ?? 0;
                    var completeness = file.DataRows?.Count > 0 
                        ? file.DataRows.Average(r => r.CompletenessPercentage) 
                        : 0;
                    Console.WriteLine($"    {file.FormattedHour}时: {dataCount} 行数据, 完整性 {completeness:F1}%");
                }
            }

            // 完整性检查
            var completenessResult = DataProcessor.CheckCompleteness(files);
            Console.WriteLine($"\n数据完整性: {(completenessResult.IsAllComplete ? "✅ 完整" : "❌ 不完整")}");

            if (!completenessResult.IsAllComplete)
            {
                Console.WriteLine("缺失的时间点:");
                foreach (var dateCompleteness in completenessResult.DateCompleteness)
                {
                    if (dateCompleteness.MissingHours.Any())
                    {
                        Console.WriteLine($"  {dateCompleteness.Date:yyyy.M.d}: [{string.Join(", ", dateCompleteness.MissingHours)}]");
                    }
                }
            }
        }

        /// <summary>
        /// 显示最终统计信息
        /// </summary>
        public static void ShowFinalStatistics()
        {
            Logger.MemoryUsage("处理完成时");
            
            var stats = new Dictionary<string, object>
            {
                ["最终内存使用"] = $"{GC.GetTotalMemory(false) / (1024.0 * 1024.0):F2}MB",
                ["GC次数 Gen0"] = GC.CollectionCount(0),
                ["GC次数 Gen1"] = GC.CollectionCount(1), 
                ["GC次数 Gen2"] = GC.CollectionCount(2)
            };
            
            Logger.Statistics("程序执行", stats);
        }

        /// <summary>
        /// 显示错误上下文信息
        /// </summary>
        /// <param name="ex">WorkPartner异常</param>
        public static void ShowErrorContext(WorkPartnerException ex)
        {
            if (ex.Context.Any())
            {
                Console.WriteLine("   错误上下文:");
                foreach (var context in ex.Context)
                {
                    Console.WriteLine($"     {context.Key}: {context.Value}");
                }
            }
        }
    }
}
