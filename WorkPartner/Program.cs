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
                args = new[] { "E:\\workspace\\gmdi\\tools\\WorkPartner\\excel" };
                // 解析命令行参数与模式分发
                var arguments = CommandLineParser.ParseCommandLineArguments(args);
                if (arguments == null)
                {
                    CommandLineParser.ShowUsage();
                    return;
                }

                if (arguments.CompareMode)
                {
                    await ModeRunner.RunCompareMode(arguments);
                    return;
                }

                if (arguments.CheckLargeValues)
                {
                    await ModeRunner.RunLargeValueCheckMode(arguments);
                    return;
                }

                if (arguments.DataCorrectionMode)
                {
                    await ModeRunner.RunDataCorrectionMode(arguments);
                    return;
                }

                var options = new ProcessingOptions
                {
                    InputPath = arguments.InputPath,
                    OutputPath = arguments.OutputPath,
                    ShowDetailedDifferences = true,
                    Tolerance = 0.001,
                    MaxDifferencesToShow = 10
                };

                var pipeline = new ProcessingPipeline();
                await pipeline.RunAsync(options);
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
