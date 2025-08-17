using WorkPartner.Models;
using WorkPartner.Utils;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Linq; // Added for .Any()

namespace WorkPartner.Utils
{
    /// <summary>
    /// 模式运行器 - 负责运行不同的程序模式
    /// </summary>
    public static class ModeRunner
    {
        /// <summary>
        /// 运行已处理结果的累计逻辑校验模式
        /// </summary>
        public static async Task RunValidateProcessedMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner 已处理结果累计逻辑校验");
            Console.WriteLine("================================");

            var dir = !string.IsNullOrEmpty(arguments.ValidateProcessedDirectory)
                ? arguments.ValidateProcessedDirectory
                : (!string.IsNullOrEmpty(arguments.OutputPath) ? arguments.OutputPath : arguments.InputPath);

            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                Console.WriteLine("❌ 请提供有效的已处理目录路径");
                Console.WriteLine("使用方法: WorkPartner.exe --validate-processed <处理后目录>");
                return;
            }

            Console.WriteLine($"📁 校验目录: {dir}");

            try
            {
                var result = DataProcessor.ValidateProcessedCumulativeLogic(dir, arguments.Tolerance);

                if (!string.IsNullOrEmpty(result.ErrorMessage))
                {
                    Console.WriteLine($"⚠️ 校验发生错误: {result.ErrorMessage}");
                    return;
                }

                Console.WriteLine($"✅ 已读取 {result.TotalFiles} 个文件，涉及 {result.TotalRows} 条数据行");

                if (result.InvalidGroups.Count == 0)
                {
                    Console.WriteLine("🎉 所有数据均符合累计逻辑: 本期累计 = 上期累计 + 本期变化");
                }
                else
                {
                    Console.WriteLine($"❗ 发现 {result.InvalidGroups.Count} 个数据名称存在不符合累计逻辑的记录");
                    if (arguments.Verbose)
                    {
                        foreach (var group in result.InvalidGroups)
                        {
                            Console.WriteLine($"\n🔸 数据名称: {group.Name}");
                            foreach (var item in group.Items)
                            {
                                Console.WriteLine($"  - 时间: {item.Timestamp:yyyy-MM-dd HH}: {item.Detail}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("不合规的数据名称列表:");
                        foreach (var group in result.InvalidGroups)
                        {
                            Console.WriteLine($"  - {group.Name}");
                        }
                        Console.WriteLine("(使用 -v 查看详细不合规项)");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 校验执行失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 运行大值检查模式
        /// </summary>
        /// <param name="arguments">命令行参数</param>
        public static async Task RunLargeValueCheckMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner Excel大值数据检查工具");
            Console.WriteLine("================================");

            // 确定要检查的目录
            string checkDirectory;
            if (!string.IsNullOrEmpty(arguments.InputPath))
            {
                checkDirectory = arguments.InputPath;
            }
            else if (!string.IsNullOrEmpty(arguments.OutputPath))
            {
                checkDirectory = arguments.OutputPath;
            }
            else
            {
                Console.WriteLine("❌ 请指定要检查的目录路径");
                Console.WriteLine("使用方法: WorkPartner.exe --check-large-values <目录路径> [--large-value-threshold <阈值>]");
                return;
            }

            Console.WriteLine($"📁 检查目录: {checkDirectory}");
            Console.WriteLine($"⚙️ 阈值: {arguments.LargeValueThreshold}");

            try
            {
                // 执行大值检查
                var checkResult = DataProcessor.CheckLargeValuesInOutputDirectory(checkDirectory, arguments.LargeValueThreshold);

                if (!string.IsNullOrEmpty(checkResult.ErrorMessage))
                {
                    Console.WriteLine($"⚠️ 检查过程发生错误: {checkResult.ErrorMessage}");
                }
                else
                {
                    Console.WriteLine($"✅ 大值数据检查完成");
                    
                    // 显示详细结果
                    if (arguments.Verbose && checkResult.FileResults.Any())
                    {
                        Console.WriteLine($"\n📊 详细检查结果:");
                        foreach (var fileResult in checkResult.FileResults)
                        {
                            Console.WriteLine($"\n📄 文件: {fileResult.FileName}");
                            Console.WriteLine($"   发现 {fileResult.LargeValues.Count} 个大值数据:");
                            
                            foreach (var largeValue in fileResult.LargeValues)
                            {
                                Console.WriteLine($"   - {largeValue.RowName} (第{largeValue.RowIndex}行, {largeValue.ColumnName}列): {largeValue.OriginalValue:F3} (绝对值: {largeValue.AbsoluteValue:F3})");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 大值检查功能执行失败: {ex.Message}");
                Logger.Error($"大值检查功能执行失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 运行文件比较模式
        /// </summary>
        /// <param name="arguments">命令行参数</param>
        public static async Task RunCompareMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner Excel文件比较工具");
            Console.WriteLine("============================");

            // 验证比较路径
            if (string.IsNullOrEmpty(arguments.CompareOriginalPath))
            {
                Console.WriteLine("❌ 原始文件目录路径不能为空");
                return;
            }

            if (string.IsNullOrEmpty(arguments.CompareProcessedPath))
            {
                Console.WriteLine("❌ 对比文件目录路径不能为空");
                return;
            }

            if (!Directory.Exists(arguments.CompareOriginalPath))
            {
                Console.WriteLine($"❌ 原始文件目录不存在: {arguments.CompareOriginalPath}");
                return;
            }

            if (!Directory.Exists(arguments.CompareProcessedPath))
            {
                Console.WriteLine($"❌ 对比文件目录不存在: {arguments.CompareProcessedPath}");
                return;
            }

            Console.WriteLine($"📁 原始文件目录: {arguments.CompareOriginalPath}");
            Console.WriteLine($"📁 对比文件目录: {arguments.CompareProcessedPath}");
            Console.WriteLine($"⚙️ 比较容差: {arguments.Tolerance}");
            Console.WriteLine($"📊 详细差异显示: {(arguments.ShowDetailedDifferences ? "启用" : "禁用")}");
            Console.WriteLine($"🔢 最大差异显示数量: {arguments.MaxDifferencesToShow}");

            try
            {
                // 执行文件比较
                var comparisonResult = DataProcessor.CompareOriginalAndProcessedFiles(
                    arguments.CompareOriginalPath,
                    arguments.CompareProcessedPath,
                    showDetailedDifferences: arguments.ShowDetailedDifferences,
                    tolerance: arguments.Tolerance,
                    maxDifferencesToShow: arguments.MaxDifferencesToShow
                );

                if (comparisonResult.HasError)
                {
                    Console.WriteLine($"⚠️ 文件比较过程发生错误: {comparisonResult.ErrorMessage}");
                }
                else
                {
                    Console.WriteLine($"✅ 文件比较分析完成");
                    
                    // 显示简要总结
                    if (arguments.Verbose)
                    {
                        Console.WriteLine($"\n📊 比较结果总结:");
                        Console.WriteLine($"   - 原始文件总数: {comparisonResult.FileComparisons.Count + comparisonResult.MissingProcessedFiles.Count}");
                        Console.WriteLine($"   - 成功比较文件数: {comparisonResult.FileComparisons.Count}");
                        Console.WriteLine($"   - 缺失对比文件数: {comparisonResult.MissingProcessedFiles.Count}");
                        Console.WriteLine($"   - 比较失败文件数: {comparisonResult.FailedComparisons.Count}");
                        
                        if (comparisonResult.TotalOriginalValues > 0)
                        {
                            var modificationPercentage = (double)comparisonResult.TotalDifferences / comparisonResult.TotalOriginalValues * 100;
                            Console.WriteLine($"   - 修改比例: {modificationPercentage:F2}% ({comparisonResult.TotalDifferences}/{comparisonResult.TotalOriginalValues})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 文件比较功能执行失败: {ex.Message}");
                Logger.Error($"文件比较功能执行失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 运行数据修正模式
        /// </summary>
        /// <param name="arguments">命令行参数</param>
        public static async Task RunDataCorrectionMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner Excel数据修正工具");
            Console.WriteLine("============================");

            // 确定要修正的目录
            string originalDirectory;
            string processedDirectory;

            if (!string.IsNullOrEmpty(arguments.CorrectionOriginalPath) && !string.IsNullOrEmpty(arguments.CorrectionProcessedPath))
            {
                originalDirectory = arguments.CorrectionOriginalPath;
                processedDirectory = arguments.CorrectionProcessedPath;
            }
            else
            {
                Console.WriteLine("❌ 请指定原目录和处理后目录路径");
                Console.WriteLine("使用方法: WorkPartner.exe --data-correction <原目录> <处理后目录>");
                return;
            }

            Console.WriteLine($"📁 原目录: {originalDirectory}");
            Console.WriteLine($"📁 处理后目录: {processedDirectory}");

            try
            {
                // 执行数据修正
                var correctionResult = DataProcessor.ProcessDataCorrection(originalDirectory, processedDirectory);

                if (correctionResult.IsSuccess)
                {
                    Console.WriteLine("\n✅ 数据修正完成");
                    Console.WriteLine($"📊 原目录文件数: {correctionResult.OriginalFilesCount}");
                    Console.WriteLine($"📊 处理后目录文件数: {correctionResult.ProcessedFilesCount}");
                    Console.WriteLine($"📊 新补充文件数: {correctionResult.SupplementFilesCount}");
                    Console.WriteLine($"📊 发现异常数据的文件数: {correctionResult.FilesWithAbnormalData}");
                    Console.WriteLine($"📊 总修正次数: {correctionResult.TotalCorrections}");
                    Console.WriteLine($"⏱️ 处理时间: {correctionResult.ProcessingTime.TotalSeconds:F2}秒");

                    if (correctionResult.FilesWithAbnormalData > 0)
                    {
                        Console.WriteLine("\n📋 修正详情:");
                        foreach (var fileCorrection in correctionResult.FileCorrections.Where(f => f.HasAbnormalData))
                        {
                            Console.WriteLine($"  📄 {fileCorrection.FileName}: 修正了 {fileCorrection.CorrectionsCount} 个异常数据");
                            foreach (var correction in fileCorrection.Corrections)
                            {
                                Console.WriteLine($"    - {correction.DataRowName} 第{correction.ColumnIndex + 1}列: {correction.OriginalValue:F2} → {correction.CorrectedValue:F2}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"❌ 数据修正失败: {correctionResult.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 数据修正过程中发生错误: {ex.Message}");
            }
        }
    }
}
