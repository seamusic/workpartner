using WorkPartner.Models;
using WorkPartner.Utils;
using WorkPartner.Services;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorkPartner
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("WorkPartner Excel数据处理工具 - 阶段5增强版");
            Console.WriteLine("==========================================");

            // 初始化日志
            Logger.Initialize("logs/workpartner.log", LogLevel.Info);
            Logger.Info("WorkPartner启动");
            Logger.MemoryUsage("启动时");

            try
            {
                using var mainOperation = Logger.StartOperation("主程序执行");
                ExceptionHandler.ClearErrorStatistics();

                // 解析命令行参数
                var arguments = CommandLineParser.ParseCommandLineArguments(args);
                if (arguments == null)
                {
                    CommandLineParser.ShowUsage();
                    return;
                }

                // 检查是否为比较模式
                if (arguments.CompareMode)
                {
                    await ModeRunner.RunCompareMode(arguments);
                    return;
                }

                // 检查是否为大值检查模式
                if (arguments.CheckLargeValues)
                {
                    await ModeRunner.RunLargeValueCheckMode(arguments);
                    return;
                }

                // 检查是否为数据修正模式
                if (arguments.DataCorrectionMode)
                {
                    await ModeRunner.RunDataCorrectionMode(arguments);
                    return;
                }

                // 验证输入路径
                if (!FileProcessor.ValidateInputPath(arguments.InputPath))
                {
                    Console.WriteLine("❌ 输入路径无效或不存在");
                    return;
                }

                // 创建输出目录
                FileProcessor.CreateOutputDirectory(arguments.OutputPath);

                // 扫描Excel文件
                var excelFiles = FileProcessor.ScanExcelFiles(arguments.InputPath);
                if (excelFiles.Count == 0)
                {
                    Console.WriteLine("❌ 未找到任何Excel文件");
                    return;
                }

                Console.WriteLine($"✅ 找到 {excelFiles.Count} 个Excel文件");

                // 解析文件名并排序
                var parsedFiles = FileProcessor.ParseAndSortFiles(excelFiles);
                if (parsedFiles.Count == 0)
                {
                    Console.WriteLine("❌ 没有找到符合格式的Excel文件");
                    return;
                }

                Console.WriteLine($"✅ 成功解析 {parsedFiles.Count} 个文件");

                // 读取Excel数据
                var filesWithData = FileProcessor.ReadExcelData(parsedFiles);
                Console.WriteLine($"✅ 成功读取 {filesWithData.Count} 个文件的数据");

                // 阶段3：数据处理逻辑
                Console.WriteLine("\n🔄 开始阶段3数据处理...");
                
                // 3.1 数据完整性检查
                Console.WriteLine("🔍 检查数据完整性...");
                var completenessResult = DataProcessor.CheckCompleteness(filesWithData);
                
                // 生成补充文件列表
                var supplementFiles = DataProcessor.GenerateSupplementFiles(filesWithData);
                
                // 创建补充文件（不包含A2列数据修改）
                if (supplementFiles.Any())
                {
                    Console.WriteLine($"📁 创建 {supplementFiles.Count} 个补充文件...");
                    var createdCount = DataProcessor.CreateSupplementFiles(supplementFiles, arguments.OutputPath);
                    Console.WriteLine($"✅ 成功创建 {createdCount} 个补充文件");
                }
                else
                {
                    Console.WriteLine("ℹ️ 无需创建补充文件，所有时间点数据都完整");
                }
                
                // 3.2 数据补充算法 - 处理所有文件（包括新创建的补充文件）
                Console.WriteLine("📊 处理缺失数据...");
                var allFilesForProcessing = DataProcessor.GetAllFilesForProcessing(filesWithData, supplementFiles, arguments.OutputPath);
                var processedFiles = DataProcessor.ProcessMissingData(allFilesForProcessing);

                // 3.3 第4、5、6列验证和重新计算 - 确保数据符合"1. 基本逻辑重构"要求
                Console.WriteLine("🔍 验证第4、5、6列数据是否符合基本逻辑重构要求...");
                var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles);
                //var validatedFiles = processedFiles;
                Console.WriteLine($"✅ 第4、5、6列验证和重新计算完成");
                                
                // 保存处理后的数据到Excel文件（包含A2列更新）
                Console.WriteLine("💾 保存处理后的数据并更新A2列...");
                await FileProcessor.SaveProcessedFiles(validatedFiles, arguments.OutputPath);
                
                // 3.4 原始文件与已处理文件比较 - 检查数据处理前后的差异（在保存之后进行比较）
                Console.WriteLine("🔍 比较原始文件与修正后文件的数值差异...");
                var originalDirectory = arguments.InputPath;
                var processedDirectory = arguments.OutputPath;
                
                try
                {
                    // 使用增强的比较功能，支持详细差异显示和自定义容差
                    var comparisonResult = DataProcessor.CompareOriginalAndProcessedFiles(
                        originalDirectory, 
                        processedDirectory,
                        showDetailedDifferences: true,  // 启用详细差异显示
                        tolerance: 0.001,               // 设置比较容差为0.001
                        maxDifferencesToShow: 10        // 每个文件最多显示10个差异
                    );
                    
                    if (comparisonResult.HasError)
                    {
                        Console.WriteLine($"⚠️ 文件比较过程发生错误: {comparisonResult.ErrorMessage}");
                    }
                    else
                    {
                        // 比较结果已在方法内部显示，这里只显示简要总结
                        Console.WriteLine($"✅ 文件比较分析完成");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ 文件比较功能执行失败: {ex.Message}");
                    Logger.Warning($"文件比较功能执行失败: {ex.Message}");
                }
                
                // 数据质量验证
                var qualityReport = DataProcessor.ValidateDataQuality(validatedFiles);

                // 显示处理结果
                ResultDisplay.DisplayProcessingResults(validatedFiles, completenessResult, supplementFiles, qualityReport);

                Console.WriteLine("\n✅ 阶段5数据处理逻辑完成！");
                
                // 显示最终统计和错误报告
                ResultDisplay.ShowFinalStatistics();
            }
            catch (WorkPartnerException ex)
            {
                Logger.Error($"WorkPartner专用错误 - {ex.Category}", ex);
                Console.WriteLine($"\n❌ 程序执行失败 ({ex.Category}): {ex.Message}");
                if (ex.FilePath != null)
                {
                    Console.WriteLine($"   相关文件: {ex.FilePath}");
                }
                ResultDisplay.ShowErrorContext(ex);
            }
            catch (Exception ex)
            {
                Logger.Error("程序执行过程中发生未知错误", ex);
                Console.WriteLine($"\n❌ 程序执行失败: {ex.Message}");
                Console.WriteLine($"   异常类型: {ex.GetType().Name}");
            }
            finally
            {
                // 最终清理工作
                Logger.MemoryUsage("程序结束时");
                Logger.Info("WorkPartner执行完成");
                
                // 显示错误报告
                var errorReport = ExceptionHandler.GenerateErrorReport();
                if (!errorReport.Contains("未发现错误"))
                {
                    Console.WriteLine("\n📊 错误统计报告:");
                    Console.WriteLine(errorReport);
                    Logger.Info("错误统计报告:");
                    Logger.Info(errorReport);
                }
                
                // 清理日志文件
                Logger.CleanupLogFile();
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }
    }
}
