using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 大值检查功能使用示例
    /// 展示如何使用 CheckLargeValuesInOutputDirectory 方法
    /// </summary>
    public static class LargeValueCheckExample
    {
        /// <summary>
        /// 基本大值检查示例
        /// </summary>
        public static void BasicLargeValueCheckExample()
        {
            Console.WriteLine("=== 基本大值检查示例 ===");
            
            var outputDirectory = @"C:\Data\Output";
            
            // 基本检查，使用默认阈值4.0
            var result = DataProcessor.CheckLargeValuesInOutputDirectory(outputDirectory);
            
            Console.WriteLine($"检查完成，发现 {result.TotalLargeValues} 个大值数据");
        }
        
        /// <summary>
        /// 自定义阈值检查示例
        /// </summary>
        public static void CustomThresholdExample()
        {
            Console.WriteLine("=== 自定义阈值检查示例 ===");
            
            var outputDirectory = @"C:\Data\Output";
            var threshold = 5.0; // 自定义阈值为5.0
            
            // 使用自定义阈值检查
            var result = DataProcessor.CheckLargeValuesInOutputDirectory(outputDirectory, threshold);
            
            Console.WriteLine($"检查完成，发现 {result.TotalLargeValues} 个大值数据");
        }
        
        /// <summary>
        /// 详细结果分析示例
        /// </summary>
        public static void DetailedAnalysisExample()
        {
            Console.WriteLine("=== 详细结果分析示例 ===");
            
            var outputDirectory = @"C:\Data\Output";
            var threshold = 3.0; // 使用较小的阈值
            
            var result = DataProcessor.CheckLargeValuesInOutputDirectory(outputDirectory, threshold);
            
            if (result.FileResults.Any())
            {
                Console.WriteLine($"\n📊 详细分析结果:");
                Console.WriteLine($"检查目录: {result.OutputDirectory}");
                Console.WriteLine($"检查阈值: {result.Threshold}");
                Console.WriteLine($"检查时间: {result.CheckTime:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"检查文件数: {result.TotalFilesChecked}");
                Console.WriteLine($"包含大值文件数: {result.FilesWithLargeValues}");
                Console.WriteLine($"大值数据总数: {result.TotalLargeValues}");
                
                foreach (var fileResult in result.FileResults)
                {
                    Console.WriteLine($"\n📄 文件: {fileResult.FileName}");
                    Console.WriteLine($"   发现 {fileResult.LargeValues.Count} 个大值数据:");
                    
                    // 按绝对值排序，显示最大的几个
                    var sortedValues = fileResult.LargeValues.OrderByDescending(v => v.AbsoluteValue).Take(10);
                    
                    foreach (var largeValue in sortedValues)
                    {
                        Console.WriteLine($"   - {largeValue.RowName} (第{largeValue.RowIndex}行, {largeValue.ColumnName}列): {largeValue.OriginalValue:F3} (绝对值: {largeValue.AbsoluteValue:F3})");
                    }
                }
            }
            else
            {
                Console.WriteLine("✅ 未发现超过阈值的数据");
            }
        }
        
        /// <summary>
        /// 批量检查示例
        /// </summary>
        public static void BatchCheckExample()
        {
            Console.WriteLine("=== 批量检查示例 ===");
            
            var directories = new[]
            {
                @"C:\Data\Output\Project1",
                @"C:\Data\Output\Project2",
                @"C:\Data\Output\Project3"
            };
            
            var thresholds = new[] { 3.0, 4.0, 5.0 };
            
            foreach (var directory in directories)
            {
                Console.WriteLine($"\n📁 检查目录: {directory}");
                
                foreach (var threshold in thresholds)
                {
                    Console.WriteLine($"  阈值 {threshold}:");
                    var result = DataProcessor.CheckLargeValuesInOutputDirectory(directory, threshold);
                    
                    if (result.TotalLargeValues > 0)
                    {
                        Console.WriteLine($"    发现 {result.TotalLargeValues} 个大值数据");
                    }
                    else
                    {
                        Console.WriteLine($"    未发现超过阈值的数据");
                    }
                }
            }
        }
        
        /// <summary>
        /// 错误处理示例
        /// </summary>
        public static void ErrorHandlingExample()
        {
            Console.WriteLine("=== 错误处理示例 ===");
            
            // 测试不存在的目录
            var nonExistentDirectory = @"C:\NonExistent\Directory";
            var result = DataProcessor.CheckLargeValuesInOutputDirectory(nonExistentDirectory);
            
            if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                Console.WriteLine($"❌ 错误: {result.ErrorMessage}");
            }
            
            // 测试空目录
            var emptyDirectory = @"C:\Empty\Directory";
            if (Directory.Exists(emptyDirectory))
            {
                var emptyResult = DataProcessor.CheckLargeValuesInOutputDirectory(emptyDirectory);
                if (!string.IsNullOrEmpty(emptyResult.ErrorMessage))
                {
                    Console.WriteLine($"❌ 错误: {emptyResult.ErrorMessage}");
                }
            }
        }
        
        /// <summary>
        /// 运行所有示例
        /// </summary>
        public static void RunAllExamples()
        {
            Console.WriteLine("🚀 开始运行大值检查功能示例...\n");
            
            try
            {
                BasicLargeValueCheckExample();
                Console.WriteLine();
                
                CustomThresholdExample();
                Console.WriteLine();
                
                DetailedAnalysisExample();
                Console.WriteLine();
                
                BatchCheckExample();
                Console.WriteLine();
                
                ErrorHandlingExample();
                Console.WriteLine();
                
                Console.WriteLine("✅ 所有示例运行完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 示例运行失败: {ex.Message}");
            }
        }
    }
}
