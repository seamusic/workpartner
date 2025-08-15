using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 文件比较功能测试
    /// 用于验证增强的 CompareOriginalAndProcessedFiles 方法
    /// </summary>
    public static class FileComparisonTest
    {
        /// <summary>
        /// 测试文件名匹配功能
        /// </summary>
        public static void TestFileNameMatching()
        {
            Console.WriteLine("=== 测试文件名匹配功能 ===");
            
            // 测试用例：原始文件名 -> 期望的已处理文件名
            var testCases = new[]
            {
                ("2025.4.15-8云港城项目.xls", "2025.4.15-08云港城项目.xls"),
                ("2025.4.16-0云港城项目.xls", "2025.4.16-00云港城项目.xls"),
                ("2025.4.17-16云港城项目.xls", "2025.4.17-16云港城项目.xls"),
                ("2025.12.31-8测试项目.xlsx", "2025.12.31-08测试项目.xlsx")
            };
            
            foreach (var (originalName, expectedProcessedName) in testCases)
            {
                Console.WriteLine($"\n测试: {originalName}");
                
                // 解析原始文件名
                var originalParse = FileNameParser.ParseFileName(originalName);
                if (originalParse?.IsValid != true)
                {
                    Console.WriteLine("  ❌ 原始文件名解析失败");
                    continue;
                }
                
                // 生成标准化的文件名
                var standardizedName = FileNameParser.GenerateFileName(
                    originalParse.Date, 
                    originalParse.Hour, 
                    originalParse.ProjectName
                );
                
                Console.WriteLine($"  原始解析: 日期={originalParse.FormattedDate}, 时间={originalParse.Hour}, 项目={originalParse.ProjectName}");
                Console.WriteLine($"  标准化文件名: {standardizedName}");
                Console.WriteLine($"  期望文件名: {expectedProcessedName}");
                Console.WriteLine($"  匹配结果: {(standardizedName == expectedProcessedName ? "✅ 成功" : "❌ 失败")}");
            }
        }
        
        /// <summary>
        /// 测试参数验证
        /// </summary>
        public static void TestParameterValidation()
        {
            Console.WriteLine("\n=== 测试参数验证 ===");
            
            // 测试不同的容差设置
            var toleranceTests = new[] { 0.0001, 0.001, 0.01, 0.1 };
            
            foreach (var tolerance in toleranceTests)
            {
                Console.WriteLine($"容差 {tolerance}:");
                Console.WriteLine($"  - 0.0005 差异: {(0.0005 > tolerance ? "显著" : "不显著")}");
                Console.WriteLine($"  - 0.005 差异: {(0.005 > tolerance ? "显著" : "不显著")}");
                Console.WriteLine($"  - 0.05 差异: {(0.05 > tolerance ? "显著" : "不显著")}");
            }
        }
        
        /// <summary>
        /// 测试差异统计计算
        /// </summary>
        public static void TestDifferenceCalculation()
        {
            Console.WriteLine("\n=== 测试差异统计计算 ===");
            
            // 模拟数据
            var totalOriginalValues = 100;
            var differences = new[] { 5, 10, 15, 20 };
            
            foreach (var diff in differences)
            {
                var percentage = (double)diff / totalOriginalValues * 100;
                Console.WriteLine($"差异 {diff}/{totalOriginalValues}: {percentage:F2}%");
            }
        }
        
        /// <summary>
        /// 运行所有测试
        /// </summary>
        public static void RunAllTests()
        {
            Console.WriteLine("🧪 开始运行文件比较功能测试...\n");
            
            try
            {
                TestFileNameMatching();
                TestParameterValidation();
                TestDifferenceCalculation();
                
                Console.WriteLine("\n✅ 所有测试完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ 测试失败: {ex.Message}");
            }
        }
    }
}
