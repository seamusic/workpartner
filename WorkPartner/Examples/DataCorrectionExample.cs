using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 数据修正功能示例
    /// 演示如何使用ProcessDataCorrection方法重新修正已处理文件中的异常数据
    /// </summary>
    public static class DataCorrectionExample
    {
        /// <summary>
        /// 运行数据修正示例
        /// </summary>
        public static void RunDataCorrectionExample()
        {
            Console.WriteLine("=== 数据修正功能示例 ===");
            Console.WriteLine();

            // 示例目录路径
            string originalDirectory = @"E:\workspace\gmdi\tools\WorkPartner\excel";
            string processedDirectory = @"E:\workspace\gmdi\tools\WorkPartner\output";

            Console.WriteLine($"📁 原目录: {originalDirectory}");
            Console.WriteLine($"📁 处理后目录: {processedDirectory}");
            Console.WriteLine();

            try
            {
                // 创建配置
                var config = new DataProcessorConfig
                {
                    CumulativeColumnPrefix = "G",
                    ChangeColumnPrefix = "D",
                    AdjustmentRange = 0.05,
                    RandomSeed = 42,
                    TimeFactorWeight = 1.0,
                    MinimumAdjustment = 0.001
                };

                Console.WriteLine("🔧 开始执行数据修正...");
                Console.WriteLine();

                // 执行数据修正
                var result = DataProcessor.ProcessDataCorrection(originalDirectory, processedDirectory, config);

                // 显示结果
                DisplayCorrectionResult(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 数据修正示例执行失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示修正结果
        /// </summary>
        private static void DisplayCorrectionResult(DataCorrectionResult result)
        {
            Console.WriteLine("=== 数据修正结果 ===");
            Console.WriteLine($"✅ 执行状态: {(result.IsSuccess ? "成功" : "失败")}");
            Console.WriteLine($"⏱️ 处理时间: {result.ProcessingTime.TotalSeconds:F2}秒");
            Console.WriteLine();

            if (result.IsSuccess)
            {
                Console.WriteLine("📊 文件统计:");
                Console.WriteLine($"  原目录文件数: {result.OriginalFilesCount}");
                Console.WriteLine($"  处理后目录文件数: {result.ProcessedFilesCount}");
                Console.WriteLine($"  新补充文件数: {result.SupplementFilesCount}");
                Console.WriteLine($"  有异常数据的文件数: {result.FilesWithAbnormalData}");
                Console.WriteLine($"  总修正次数: {result.TotalCorrections}");
                Console.WriteLine();

                if (result.FilesWithAbnormalData > 0)
                {
                    Console.WriteLine("📋 修正详情:");
                    foreach (var fileCorrection in result.FileCorrections.Where(f => f.HasAbnormalData))
                    {
                        Console.WriteLine($"  📄 {fileCorrection.FileName}:");
                        Console.WriteLine($"    修正了 {fileCorrection.CorrectionsCount} 个异常数据");
                        
                        foreach (var correction in fileCorrection.Corrections)
                        {
                            Console.WriteLine($"    - {correction.DataRowName} 第{correction.ColumnIndex + 1}列:");
                            Console.WriteLine($"      原始值: {correction.OriginalValue:F3}");
                            Console.WriteLine($"      修正值: {correction.CorrectedValue:F3}");
                            Console.WriteLine($"      修正期数: {correction.CorrectionPeriods}");
                            Console.WriteLine($"      修正时间: {correction.CorrectionTime:yyyy-MM-dd HH:mm:ss}");
                            Console.WriteLine();
                        }
                    }
                }
                else
                {
                    Console.WriteLine("✅ 没有发现需要修正的异常数据");
                }
            }
            else
            {
                Console.WriteLine($"❌ 修正失败: {result.ErrorMessage}");
            }
        }

        /// <summary>
        /// 演示命令行使用方式
        /// </summary>
        public static void ShowCommandLineUsage()
        {
            Console.WriteLine("=== 命令行使用方式 ===");
            Console.WriteLine();
            Console.WriteLine("基本用法:");
            Console.WriteLine("  WorkPartner.exe --data-correction <原目录> <处理后目录>");
            Console.WriteLine();
            Console.WriteLine("示例:");
            Console.WriteLine("  WorkPartner.exe --data-correction C:\\excel C:\\output");
            Console.WriteLine("  WorkPartner.exe --data-correction E:\\workspace\\gmdi\\tools\\WorkPartner\\excel E:\\workspace\\gmdi\\tools\\WorkPartner\\output");
            Console.WriteLine();
            Console.WriteLine("功能说明:");
            Console.WriteLine("  1. 读取原目录和处理后目录下的所有Excel文件");
            Console.WriteLine("  2. 识别新补充的数据（不在原目录中的文件）");
            Console.WriteLine("  3. 检查第7、8、9列（索引3、4、5）的值是否超过4");
            Console.WriteLine("  4. 如果发现异常数据，往前5期开始处理:");
            Console.WriteLine("     - 生成每一期的本期变化量（第4、5、6列），变化范围-0.5至0.5");
            Console.WriteLine("     - 计算累计变化量：本期累计 = 上期累计 + 本期变化");
            Console.WriteLine();
        }
    }
}
