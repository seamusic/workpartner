using WorkPartner.Services;
using System;
using System.IO;

namespace WorkPartner.Examples
{
    /// <summary>
    /// .xls 文件处理使用示例
    /// </summary>
    public class XlsUsageExample
    {
        private readonly IExcelService _excelService;
        
        public XlsUsageExample()
        {
            _excelService = new ExcelService();
        }
        
        /// <summary>
        /// 处理 Excel 文件（支持 .xls 和 .xlsx）
        /// </summary>
        /// <param name="filePath">Excel 文件路径</param>
        public void ProcessExcelFile(string filePath)
        {
            try
            {
                Console.WriteLine($"开始处理文件: {Path.GetFileName(filePath)}");
                
                // 1. 验证文件
                if (!_excelService.ValidateExcelFile(filePath))
                {
                    Console.WriteLine("❌ 文件格式无效或文件损坏");
                    return;
                }
                
                var extension = Path.GetExtension(filePath).ToLower();
                Console.WriteLine($"✓ 检测到 {extension} 格式文件");
                
                // 2. 检查文件是否被占用
                if (_excelService.IsFileLocked(filePath))
                {
                    Console.WriteLine("❌ 文件被其他程序占用，请关闭后重试");
                    return;
                }
                
                // 3. 读取文件数据
                Console.WriteLine("正在读取文件数据...");
                var excelFile = _excelService.ReadExcelFile(filePath);
                
                // 4. 显示文件信息
                Console.WriteLine($"文件信息:");
                Console.WriteLine($"  文件名: {excelFile.FileName}");
                Console.WriteLine($"  文件大小: {excelFile.FileSize:N0} 字节");
                Console.WriteLine($"  最后修改: {excelFile.LastModified:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"  数据行数: {excelFile.DataRows.Count}");
                
                // 5. 显示前几行数据示例
                Console.WriteLine("\n数据预览 (前5行):");
                for (int i = 0; i < Math.Min(5, excelFile.DataRows.Count); i++)
                {
                    var row = excelFile.DataRows[i];
                    Console.WriteLine($"  行{row.RowIndex}: {row.Name} ({row.Values.Count} 个数值)");
                }
                
                Console.WriteLine("✓ 文件处理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 处理文件时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 批量处理文件夹中的 Excel 文件
        /// </summary>
        /// <param name="folderPath">文件夹路径</param>
        public void ProcessExcelFolder(string folderPath)
        {
            try
            {
                Console.WriteLine($"扫描文件夹: {folderPath}");
                
                var excelFiles = Directory.GetFiles(folderPath, "*.xls*", SearchOption.TopDirectoryOnly);
                
                Console.WriteLine($"找到 {excelFiles.Length} 个 Excel 文件");
                
                foreach (var file in excelFiles)
                {
                    Console.WriteLine($"\n--- 处理文件 {Path.GetFileName(file)} ---");
                    ProcessExcelFile(file);
                }
                
                Console.WriteLine("\n✓ 批量处理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 批量处理时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 文件格式转换示例（可选功能）
        /// </summary>
        /// <param name="xlsFilePath">.xls 文件路径</param>
        /// <param name="outputPath">输出 .xlsx 文件路径</param>
        public void ConvertXlsToXlsx(string xlsFilePath, string outputPath)
        {
            try
            {
                Console.WriteLine($"转换文件: {Path.GetFileName(xlsFilePath)} -> {Path.GetFileName(outputPath)}");
                
                // 读取 .xls 文件
                var excelFile = _excelService.ReadExcelFile(xlsFilePath);
                
                // 保存为 .xlsx 格式
                var success = _excelService.SaveExcelFile(excelFile, outputPath);
                
                if (success)
                {
                    Console.WriteLine("✓ 转换完成");
                }
                else
                {
                    Console.WriteLine("❌ 转换失败");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 转换时出错: {ex.Message}");
            }
        }
    }
}