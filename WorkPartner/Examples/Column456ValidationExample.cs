using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 第4、5、6列验证和重新计算示例
    /// </summary>
    public class Column456ValidationExample
    {
        /// <summary>
        /// 演示完整的处理流程：先补充缺失数据，再验证和重新计算第4、5、6列
        /// </summary>
        public static void DemonstrateCompleteWorkflow()
        {
            Console.WriteLine("🚀 开始演示完整的处理流程...");
            
            // 1. 加载Excel文件
            var allFilesForProcessing = LoadExcelFiles();
            Console.WriteLine($"📁 加载了 {allFilesForProcessing.Count} 个Excel文件");
            
            // 2. 执行缺失数据补充
            Console.WriteLine("🔄 执行缺失数据补充...");
            var processedFiles = DataProcessor.ProcessMissingData(allFilesForProcessing);
            Console.WriteLine($"✅ 缺失数据补充完成，处理了 {processedFiles.Count} 个文件");
            
            // 3. 验证并重新计算第4、5、6列的值
            Console.WriteLine("🔍 验证并重新计算第4、5、6列的值...");
            var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles);
            Console.WriteLine($"✅ 第4、5、6列验证和重新计算完成");
            
            // 4. 保存处理后的文件
            SaveProcessedFiles(validatedFiles);
            Console.WriteLine("💾 所有处理后的文件已保存");
            
            Console.WriteLine("🎉 完整处理流程演示完成！");
        }
        
        /// <summary>
        /// 仅验证和重新计算第4、5、6列（适用于数据已补充的情况）
        /// </summary>
        public static void ValidateColumns456Only()
        {
            Console.WriteLine("🔍 开始验证和重新计算第4、5、6列...");
            
            // 加载已处理的文件
            var processedFiles = LoadProcessedFiles();
            Console.WriteLine($"📁 加载了 {processedFiles.Count} 个已处理的文件");
            
            // 验证并重新计算第4、5、6列
            var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles);
            Console.WriteLine($"✅ 第4、5、6列验证和重新计算完成");
            
            // 保存验证后的文件
            SaveProcessedFiles(validatedFiles);
            Console.WriteLine("💾 验证后的文件已保存");
        }
        
        /// <summary>
        /// 使用自定义配置进行验证
        /// </summary>
        public static void ValidateWithCustomConfig()
        {
            Console.WriteLine("⚙️ 使用自定义配置进行验证...");
            
            // 创建自定义配置
            var customConfig = new DataProcessorConfig
            {
                ColumnValidationTolerance = 0.005, // 更严格的误差容忍度（0.5%）
                EnableDetailedLogging = true,
                EnablePerformanceMonitoring = true
            };
            
            // 加载已处理的文件
            var processedFiles = LoadProcessedFiles();
            Console.WriteLine($"📁 加载了 {processedFiles.Count} 个已处理的文件");
            
            // 使用自定义配置验证
            var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles, customConfig);
            Console.WriteLine($"✅ 使用自定义配置验证完成");
            
            // 保存验证后的文件
            SaveProcessedFiles(validatedFiles);
            Console.WriteLine("💾 验证后的文件已保存");
        }
        
        /// <summary>
        /// 加载Excel文件（示例实现）
        /// </summary>
        private static List<ExcelFile> LoadExcelFiles()
        {
            // 这里应该是实际的Excel文件加载逻辑
            // 为了示例，返回空列表
            return new List<ExcelFile>();
        }
        
        /// <summary>
        /// 加载已处理的文件（示例实现）
        /// </summary>
        private static List<ExcelFile> LoadProcessedFiles()
        {
            // 这里应该是实际的已处理文件加载逻辑
            // 为了示例，返回空列表
            return new List<ExcelFile>();
        }
        
        /// <summary>
        /// 保存处理后的文件（示例实现）
        /// </summary>
        private static void SaveProcessedFiles(List<ExcelFile> files)
        {
            // 这里应该是实际的文件保存逻辑
            // 为了示例，只是输出信息
            Console.WriteLine($"💾 保存 {files.Count} 个处理后的文件");
        }
    }
}
