using WorkPartner.Utils;
using WorkPartner.Models;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 命令行参数测试
    /// 用于验证新的命令行参数功能
    /// </summary>
    public static class CommandLineTest
    {
        /// <summary>
        /// 测试命令行参数解析
        /// </summary>
        public static void TestCommandLineParsing()
        {
            Console.WriteLine("=== 测试命令行参数解析 ===");
            
            // 测试用例
            var testCases = new[]
            {
                // 基本数据处理模式
                new { Args = new[] { "C:\\excel" }, Description = "基本数据处理模式" },
                new { Args = new[] { "C:\\excel", "-o", "C:\\output" }, Description = "指定输出目录" },
                new { Args = new[] { "C:\\excel", "-v" }, Description = "详细输出模式" },
                
                // 文件比较模式
                new { Args = new[] { "-c", "C:\\original", "C:\\processed" }, Description = "基本比较模式" },
                new { Args = new[] { "-v", "C:\\original", "C:\\processed" }, Description = "简化比较模式" },
                new { Args = new[] { "-c", "C:\\original", "C:\\processed", "--detailed" }, Description = "详细比较模式" },
                new { Args = new[] { "-c", "C:\\original", "C:\\processed", "--tolerance", "0.01" }, Description = "自定义容差" },
                new { Args = new[] { "-c", "C:\\original", "C:\\processed", "--max-differences", "5" }, Description = "限制差异数量" },
                new { Args = new[] { "-c", "C:\\original", "C:\\processed", "--detailed", "--tolerance", "0.001", "--max-differences", "20" }, Description = "完整比较模式" },
                
                // 帮助模式
                new { Args = new[] { "-h" }, Description = "帮助模式" },
                new { Args = new[] { "--help" }, Description = "帮助模式" }
            };
            
            foreach (var testCase in testCases)
            {
                Console.WriteLine($"\n测试: {testCase.Description}");
                Console.WriteLine($"参数: {string.Join(" ", testCase.Args)}");
                
                try
                {
                    // 这里我们模拟参数解析逻辑
                    var result = SimulateCommandLineParsing(testCase.Args);
                    Console.WriteLine($"结果: {result}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"错误: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 模拟命令行参数解析
        /// </summary>
        private static string SimulateCommandLineParsing(string[] args)
        {
            if (args.Length == 0)
                return "无参数";
                
            var firstArg = args[0].ToLower();
            
            switch (firstArg)
            {
                case "-h":
                case "--help":
                    return "显示帮助信息";
                    
                case "-c":
                case "--compare":
                    if (args.Length >= 3)
                    {
                        var originalPath = args[1];
                        var processedPath = args[2];
                        var options = new List<string>();
                        
                        // 解析额外选项
                        for (int i = 3; i < args.Length; i++)
                        {
                            switch (args[i].ToLower())
                            {
                                case "--detailed":
                                    options.Add("详细差异显示");
                                    break;
                                case "--tolerance":
                                    if (i + 1 < args.Length)
                                    {
                                        options.Add($"容差: {args[++i]}");
                                    }
                                    break;
                                case "--max-differences":
                                    if (i + 1 < args.Length)
                                    {
                                        options.Add($"最大差异数: {args[++i]}");
                                    }
                                    break;
                            }
                        }
                        
                        var optionsStr = options.Any() ? $" ({string.Join(", ", options)})" : "";
                        return $"比较模式: {originalPath} -> {processedPath}{optionsStr}";
                    }
                    return "比较模式参数不足";
                    
                case "-v":
                    if (args.Length >= 3)
                    {
                        var originalPath = args[1];
                        var processedPath = args[2];
                        return $"简化比较模式: {originalPath} -> {processedPath} (详细输出)";
                    }
                    return "简化比较模式参数不足";
                    
                default:
                    // 数据处理模式
                    var inputPath = args[0];
                    var outputPath = "默认输出路径";
                    var verbose = false;
                    
                    for (int i = 1; i < args.Length; i++)
                    {
                        switch (args[i].ToLower())
                        {
                            case "-o":
                            case "--output":
                                if (i + 1 < args.Length)
                                {
                                    outputPath = args[++i];
                                }
                                break;
                            case "-v":
                            case "--verbose":
                                verbose = true;
                                break;
                        }
                    }
                    
                    var verboseStr = verbose ? " (详细输出)" : "";
                    return $"数据处理模式: {inputPath} -> {outputPath}{verboseStr}";
            }
        }
        
        /// <summary>
        /// 测试文件名匹配功能
        /// </summary>
        public static void TestFileNameMatching()
        {
            Console.WriteLine("\n=== 测试文件名匹配功能 ===");
            
            var testCases = new[]
            {
                ("2025.4.15-8云港城项目.xls", "2025.4.15-08云港城项目.xls"),
                ("2025.4.16-0云港城项目.xls", "2025.4.16-00云港城项目.xls"),
                ("2025.4.17-16云港城项目.xls", "2025.4.17-16云港城项目.xls")
            };
            
            foreach (var (originalName, expectedName) in testCases)
            {
                Console.WriteLine($"\n测试: {originalName} -> {expectedName}");
                
                var parseResult = FileNameParser.ParseFileName(originalName);
                if (parseResult?.IsValid == true)
                {
                    var standardizedName = FileNameParser.GenerateFileName(
                        parseResult.Date, 
                        parseResult.Hour, 
                        parseResult.ProjectName
                    );
                    
                    var isMatch = standardizedName == expectedName;
                    Console.WriteLine($"  标准化文件名: {standardizedName}");
                    Console.WriteLine($"  匹配结果: {(isMatch ? "✅ 成功" : "❌ 失败")}");
                }
                else
                {
                    Console.WriteLine("  ❌ 文件名解析失败");
                }
            }
        }
        
        /// <summary>
        /// 测试数据修正参数解析
        /// </summary>
        public static void TestDataCorrectionParsing()
        {
            Console.WriteLine("=== 测试数据修正参数解析 ===");
            
            // 测试数据修正模式
            var testArgs = new[] { "--data-correction", "C:\\original", "C:\\processed" };
            var arguments = CommandLineParser.ParseCommandLineArguments(testArgs);
            
            if (arguments != null)
            {
                Console.WriteLine($"✅ 参数解析成功");
                Console.WriteLine($"数据修正模式: {arguments.DataCorrectionMode}");
                Console.WriteLine($"原目录: {arguments.CorrectionOriginalPath}");
                Console.WriteLine($"处理后目录: {arguments.CorrectionProcessedPath}");
            }
            else
            {
                Console.WriteLine("❌ 参数解析失败");
            }
            
            Console.WriteLine();
        }
        
        /// <summary>
        /// 运行所有测试
        /// </summary>
        public static void RunAllTests()
        {
            Console.WriteLine("🧪 开始运行命令行参数测试...\n");
            
            try
            {
                TestCommandLineParsing();
                TestFileNameMatching();
                
                Console.WriteLine("\n✅ 所有测试完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ 测试失败: {ex.Message}");
            }
        }
    }
}
