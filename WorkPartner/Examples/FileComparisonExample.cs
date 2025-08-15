using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 文件比较功能使用示例
    /// 展示如何使用增强的 CompareOriginalAndProcessedFiles 方法
    /// </summary>
    public static class FileComparisonExample
    {
        /// <summary>
        /// 基本比较示例
        /// </summary>
        public static void BasicComparisonExample()
        {
            Console.WriteLine("=== 基本文件比较示例 ===");
            
            var originalDirectory = @"C:\Data\Original";
            var processedDirectory = @"C:\Data\Processed";
            
            // 基本比较，使用默认配置
            var result = DataProcessor.CompareOriginalAndProcessedFiles(
                originalDirectory, 
                processedDirectory
            );
            
            Console.WriteLine($"比较完成，发现 {result.TotalDifferences} 个差异");
        }
        
        /// <summary>
        /// 详细差异分析示例
        /// </summary>
        public static void DetailedComparisonExample()
        {
            Console.WriteLine("=== 详细差异分析示例 ===");
            
            var originalDirectory = @"C:\Data\Original";
            var processedDirectory = @"C:\Data\Processed";
            
            // 启用详细差异显示，设置较小的容差
            var result = DataProcessor.CompareOriginalAndProcessedFiles(
                originalDirectory, 
                processedDirectory,
                showDetailedDifferences: true,  // 显示详细差异
                tolerance: 0.0001,              // 设置严格的容差
                maxDifferencesToShow: 5         // 每个文件最多显示5个差异
            );
            
            Console.WriteLine($"详细分析完成，发现 {result.TotalDifferences} 个差异");
        }
        
        /// <summary>
        /// 自定义容差比较示例
        /// </summary>
        public static void CustomToleranceExample()
        {
            Console.WriteLine("=== 自定义容差比较示例 ===");
            
            var originalDirectory = @"C:\Data\Original";
            var processedDirectory = @"C:\Data\Processed";
            
            // 使用较大的容差，只关注显著差异
            var result = DataProcessor.CompareOriginalAndProcessedFiles(
                originalDirectory, 
                processedDirectory,
                showDetailedDifferences: true,
                tolerance: 0.01,                // 设置较大的容差
                maxDifferencesToShow: 3         // 限制显示数量
            );
            
            Console.WriteLine($"自定义容差分析完成，发现 {result.TotalSignificantDifferences} 个显著差异");
        }
        
        /// <summary>
        /// 批量比较示例
        /// </summary>
        public static void BatchComparisonExample()
        {
            Console.WriteLine("=== 批量比较示例 ===");
            
            var directories = new[]
            {
                (@"C:\Data\Original\Project1", @"C:\Data\Processed\Project1"),
                (@"C:\Data\Original\Project2", @"C:\Data\Processed\Project2"),
                (@"C:\Data\Original\Project3", @"C:\Data\Processed\Project3")
            };
            
            foreach (var (originalDir, processedDir) in directories)
            {
                Console.WriteLine($"\n比较项目: {Path.GetFileName(originalDir)}");
                
                var result = DataProcessor.CompareOriginalAndProcessedFiles(
                    originalDir, 
                    processedDir,
                    showDetailedDifferences: false,  // 批量处理时不显示详细差异
                    tolerance: 0.001
                );
                
                if (result.TotalOriginalValues > 0)
                {
                    var modificationPercentage = (double)result.TotalDifferences / result.TotalOriginalValues * 100;
                    Console.WriteLine($"  修改比例: {modificationPercentage:F2}%");
                }
            }
        }
        
        /// <summary>
        /// 文件名匹配测试示例
        /// 测试处理文件名时间格式差异（如 8 变成 08）
        /// </summary>
        public static void FileNameMatchingTestExample()
        {
            Console.WriteLine("=== 文件名匹配测试示例 ===");
            
            // 模拟原始文件名和已处理文件名的差异
            var testCases = new[]
            {
                ("2025.4.15-8云港城项目.xls", "2025.4.15-08云港城项目.xls"),
                ("2025.4.16-0云港城项目.xls", "2025.4.16-00云港城项目.xls"),
                ("2025.4.17-16云港城项目.xls", "2025.4.17-16云港城项目.xls")
            };
            
            foreach (var (originalName, processedName) in testCases)
            {
                Console.WriteLine($"测试: {originalName} -> {processedName}");
                
                // 这里可以测试文件名匹配逻辑
                var originalParse = FileNameParser.ParseFileName(originalName);
                var processedParse = FileNameParser.ParseFileName(processedName);
                
                if (originalParse?.IsValid == true && processedParse?.IsValid == true)
                {
                    var isMatch = originalParse.Date == processedParse.Date && 
                                 originalParse.Hour == processedParse.Hour &&
                                 originalParse.ProjectName == processedParse.ProjectName;
                    
                    Console.WriteLine($"  匹配结果: {(isMatch ? "✅ 成功" : "❌ 失败")}");
                }
            }
        }
        
        /// <summary>
        /// 运行所有示例
        /// </summary>
        public static void RunAllExamples()
        {
            Console.WriteLine("🚀 开始运行文件比较功能示例...\n");
            
            try
            {
                BasicComparisonExample();
                Console.WriteLine();
                
                DetailedComparisonExample();
                Console.WriteLine();
                
                CustomToleranceExample();
                Console.WriteLine();
                
                BatchComparisonExample();
                Console.WriteLine();
                
                FileNameMatchingTestExample();
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
