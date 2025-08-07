using WorkPartner.Utils;
using System;

namespace WorkPartner.Examples
{
    /// <summary>
    /// Excel配置使用示例
    /// 展示如何使用ExcelConfiguration类来管理Excel读取配置
    /// </summary>
    public class ExcelConfigurationExample
    {
        /// <summary>
        /// 演示基本配置使用
        /// </summary>
        public static void DemonstrateBasicUsage()
        {
            Console.WriteLine("=== Excel配置管理示例 ===");
            
            // 获取配置实例
            var config = ExcelConfiguration.Instance;
            
            // 显示当前配置
            Console.WriteLine($"当前配置: {config.GetConfigurationSummary()}");
            
            // 修改配置
            config.StartRow = 10;
            config.EndRow = 100;
            config.StartCol = 3;
            config.EndCol = 8;
            config.NameCol = 2;
            
            Console.WriteLine($"修改后配置: {config.GetConfigurationSummary()}");
            
            // 配置已通过appsettings.json管理，无需手动保存
            Console.WriteLine("配置已通过appsettings.json管理");
            
            // 重置为默认配置
            config.ResetToDefault();
            Console.WriteLine($"重置后配置: {config.GetConfigurationSummary()}");
        }
        
        /// <summary>
        /// 演示从Excel文件动态读取配置
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径</param>
        public static void DemonstrateDynamicConfiguration(string excelFilePath)
        {
            Console.WriteLine("=== 动态配置读取示例 ===");
            
            var config = ExcelConfiguration.Instance;
            
            // 显示原始配置
            Console.WriteLine($"原始配置: {config.GetConfigurationSummary()}");
            
            // 从Excel文件读取配置
            if (config.LoadConfigurationFromExcel(excelFilePath))
            {
                Console.WriteLine($"从Excel文件读取的配置: {config.GetConfigurationSummary()}");
                
                // 动态配置仅在内存中，如需持久化请修改appsettings.json
                Console.WriteLine("动态配置仅在内存中，如需持久化请修改appsettings.json");
            }
            else
            {
                Console.WriteLine("无法从Excel文件读取配置，使用默认配置");
            }
        }
        
        /// <summary>
        /// 演示配置验证
        /// </summary>
        public static void DemonstrateConfigurationValidation()
        {
            Console.WriteLine("=== 配置验证示例 ===");
            
            var config = ExcelConfiguration.Instance;
            
            // 测试有效配置
            Console.WriteLine($"有效配置验证: {config.ValidateConfiguration()}");
            
            // 测试无效配置
            config.StartRow = -1;
            config.EndRow = 0;
            Console.WriteLine($"无效配置验证: {config.ValidateConfiguration()}");
            
            // 配置会被自动重置为默认值
            Console.WriteLine($"重置后配置: {config.GetConfigurationSummary()}");
        }
        
        /// <summary>
        /// 演示配置的扩展性
        /// </summary>
        public static void DemonstrateExtensibility()
        {
            Console.WriteLine("=== 配置扩展性示例 ===");
            
            var config = ExcelConfiguration.Instance;
            
            // 模拟不同Excel格式的配置
            var configurations = new[]
            {
                new { Name = "标准格式", StartRow = 5, EndRow = 368, StartCol = 4, EndCol = 9, NameCol = 2 },
                new { Name = "紧凑格式", StartRow = 3, EndRow = 200, StartCol = 2, EndCol = 7, NameCol = 1 },
                new { Name = "扩展格式", StartRow = 10, EndRow = 500, StartCol = 5, EndCol = 12, NameCol = 3 }
            };
            
            foreach (var format in configurations)
            {
                config.StartRow = format.StartRow;
                config.EndRow = format.EndRow;
                config.StartCol = format.StartCol;
                config.EndCol = format.EndCol;
                config.NameCol = format.NameCol;
                
                Console.WriteLine($"{format.Name}: {config.GetConfigurationSummary()}");
            }
        }
        
        /// <summary>
        /// 运行所有示例
        /// </summary>
        public static void RunAllExamples()
        {
            try
            {
                DemonstrateBasicUsage();
                Console.WriteLine();
                
                DemonstrateConfigurationValidation();
                Console.WriteLine();
                
                DemonstrateExtensibility();
                Console.WriteLine();
                
                // 如果有Excel文件，演示动态配置
                var sampleExcelPath = "path/to/sample.xlsx"; // 需要替换为实际路径
                if (System.IO.File.Exists(sampleExcelPath))
                {
                    DemonstrateDynamicConfiguration(sampleExcelPath);
                }
                
                Console.WriteLine("所有示例运行完成！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"示例运行出错: {ex.Message}");
            }
        }
    }
} 