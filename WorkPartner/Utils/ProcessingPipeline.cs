using System;
using System.Linq;
using System.Collections.Generic;
using WorkPartner.Models;
using WorkPartner.Services;
using System.Threading.Tasks;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 将 Program.cs 中的主编排流程抽离为可测试的管线。
    /// 仅做结构拆分，不改变任何业务逻辑。
    /// </summary>
    public class ProcessingPipeline
    {
        public ProcessingPipeline()
        {
        }

        public async Task RunAsync(ProcessingOptions options)
        {
            using var mainOperation = Logger.StartOperation("主程序执行");
            ExceptionHandler.ClearErrorStatistics();

            // 验证输入路径
            if (!FileProcessor.ValidateInputPath(options.InputPath))
            {
                Console.WriteLine("❌ 输入路径无效或不存在");
                return;
            }

            // 创建输出目录
            FileProcessor.CreateOutputDirectory(options.OutputPath);

            // 扫描Excel文件
            var excelFiles = FileProcessor.ScanExcelFiles(options.InputPath);
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
                var createdCount = DataProcessor.CreateSupplementFiles(supplementFiles, options.OutputPath);
                Console.WriteLine($"✅ 成功创建 {createdCount} 个补充文件");
            }
            else
            {
                Console.WriteLine("ℹ️ 无需创建补充文件，所有时间点数据都完整");
            }

            // 3.2 数据补充算法 - 处理所有文件（包括新创建的补充文件）
            Console.WriteLine("📊 处理缺失数据...");
            var allFilesForProcessing = DataProcessor.GetAllFilesForProcessing(filesWithData, supplementFiles, options.OutputPath);
            var processors = DependencyInjection.ServiceCollectionExtensions.CreateDefaultProcessors();
            var processedFiles = processors.dataProcessor.ProcessMissingData(allFilesForProcessing);

            // 3.3 第4、5、6列验证和重新计算 - 确保数据符合"1. 基本逻辑重构"要求
            Console.WriteLine("🔍 验证第4、5、6列数据是否符合基本逻辑重构要求...");
            var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles);
            Console.WriteLine($"✅ 第4、5、6列验证和重新计算完成");

            // 保存处理后的数据到Excel文件（包含A2列更新）
            Console.WriteLine("💾 保存处理后的数据并更新A2列...");
            await FileProcessor.SaveProcessedFiles(validatedFiles, options.OutputPath);

            // 3.4 原始文件与已处理文件比较 - 检查数据处理前后的差异（在保存之后进行比较）
            Console.WriteLine("🔍 比较原始文件与修正后文件的数值差异...");
            var originalDirectory = options.InputPath;
            var processedDirectory = options.OutputPath;

            try
            {
                // 使用增强的比较功能，支持详细差异显示和自定义容差
                var comparisonResult = DataProcessor.CompareOriginalAndProcessedFiles(
                    originalDirectory,
                    processedDirectory,
                    showDetailedDifferences: options.ShowDetailedDifferences,
                    tolerance: options.Tolerance,
                    maxDifferencesToShow: options.MaxDifferencesToShow
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

            Console.WriteLine("\n✅ 阶段5数据处理逻辑完成！\n");

            // 显示最终统计和错误报告
            ResultDisplay.ShowFinalStatistics();
        }
    }
}


