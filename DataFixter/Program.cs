using DataFixter.Models;
using DataFixter.Services;
using DataFixter.Excel;
using Serilog;
using DataFixter.Logging;

namespace DataFixter
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 配置Serilog日志
                LoggingConfiguration.ConfigureLogging();
                
                var logger = Log.ForContext<Program>();
                logger.Information("=== DataFixter 数据修正工具启动 ===");
                
                // 显示欢迎信息
                ShowWelcomeMessage();

                // 处理命令行参数
                if (!TryParseArguments(args, out var processedDirectory, out var comparisonDirectory))
                {
                    ShowUsage();
                    return;
                }

                // 验证目录
                if (!ValidateDirectories(processedDirectory, comparisonDirectory))
                {
                    return;
                }

                // 执行批量处理
                var result = ExecuteBatchProcessing(processedDirectory, comparisonDirectory);

                // 显示处理结果
                ShowProcessingResults(result);
                
                logger.Information("=== DataFixter 数据修正工具完成 ===");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "DataFixter 运行时发生致命错误: {Message}", ex.Message);
                Console.WriteLine($"DataFixter 运行时发生致命错误: {ex.Message}");
            }
            finally
            {
                // 关闭日志系统
                LoggingConfiguration.CloseLogging();
            }
        }

        /// <summary>
        /// 显示欢迎信息
        /// </summary>
        private static void ShowWelcomeMessage()
        {
            Console.WriteLine("=== DataFixter 数据修正工具 ===");
            Console.WriteLine("用于修复监测数据中累计变化量计算错误的工具");
            Console.WriteLine();
        }

        /// <summary>
        /// 解析命令行参数
        /// </summary>
        /// <param name="args">命令行参数</param>
        /// <param name="processedDirectory">待处理目录</param>
        /// <param name="comparisonDirectory">对比目录</param>
        /// <returns>是否解析成功</returns>
        private static bool TryParseArguments(string[] args, out string processedDirectory, out string comparisonDirectory)
        {
            processedDirectory = string.Empty;
            comparisonDirectory = string.Empty;

            if (args.Length != 2)
            {
                return false;
            }

            processedDirectory = args[0];
            comparisonDirectory = args[1];
            return true;
        }

        /// <summary>
        /// 验证目录是否存在
        /// </summary>
        /// <param name="processedDirectory">待处理目录</param>
        /// <param name="comparisonDirectory">对比目录</param>
        /// <returns>是否验证通过</returns>
        private static bool ValidateDirectories(string processedDirectory, string comparisonDirectory)
        {
            if (!Directory.Exists(processedDirectory))
            {
                Console.WriteLine($"错误: 待处理目录不存在: {processedDirectory}");
                return false;
            }

            if (!Directory.Exists(comparisonDirectory))
            {
                Console.WriteLine($"错误: 对比目录不存在: {comparisonDirectory}");
                return false;
            }

            Console.WriteLine($"待处理目录: {processedDirectory}");
            Console.WriteLine($"对比目录: {comparisonDirectory}");
            Console.WriteLine();
            
            return true;
        }

        /// <summary>
        /// 执行完整的批量处理流程
        /// </summary>
        /// <param name="processedDirectory">待处理目录</param>
        /// <param name="comparisonDirectory">对比目录</param>
        /// <returns>处理结果</returns>
        private static ProcessingResult ExecuteBatchProcessing(string processedDirectory, string comparisonDirectory)
        {
            var logger = new DualLoggerService(typeof(Program));
            var result = new ProcessingResult();

            try
            {
                logger.LogOperationStart("批量处理", $"待处理目录：{processedDirectory}，对比目录：{comparisonDirectory}");

                // 创建批量处理服务
                var batchProcessor = new BatchProcessingService(logger);
                
                // 执行批量处理
                result = batchProcessor.ProcessBatch(processedDirectory, comparisonDirectory);

                logger.LogOperationComplete("批量处理", "成功完成", 
                    $"处理文件：{result.ProcessedFiles}个，监测点：{result.MonitoringPoints}个，修正记录：{result.CorrectionResult?.AdjustmentRecords.Count ?? 0}条");
            }
            catch (Exception ex)
            {
                logger.ShowError($"批量处理过程中发生异常: {ex.Message}");
                logger.FileError(ex, "批量处理过程中发生异常");
                result.Status = ProcessingStatus.Error;
                result.Message = $"处理过程中发生异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 显示处理结果
        /// </summary>
        /// <param name="result">处理结果</param>
        private static void ShowProcessingResults(ProcessingResult result)
        {
            Console.WriteLine("=== 处理完成 ===");
            Console.WriteLine(result.GetSummary());
            Console.WriteLine();
            Console.WriteLine("详细报告已生成到输出目录");
            Console.WriteLine();
            Console.WriteLine("=== 验证修正结果 ===");
            Console.WriteLine("1. 检查 '修正后' 目录中的Excel文件");
            Console.WriteLine("2. 对比原始文件和修正后文件的数据");
            Console.WriteLine("3. 查看 '修正详细报告.txt' 了解具体修正内容");
            Console.WriteLine("4. 如果文件大小相同，这是正常的Excel格式特性");
            Console.WriteLine("5. 重点检查列3-8的数据是否已按修正逻辑更新");
        }

        /// <summary>
        /// 显示使用说明
        /// </summary>
        private static void ShowUsage()
        {
            Console.WriteLine("使用方法:");
            Console.WriteLine("DataFixter <待处理目录> <对比目录>");
            Console.WriteLine();
            Console.WriteLine("参数说明:");
            Console.WriteLine("  待处理目录: 包含需要修正的Excel文件的目录");
            Console.WriteLine("  对比目录: 包含对比数据的Excel文件的目录");
            Console.WriteLine();
        }
    }
}
