using WorkPartner.Services;
using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner.Examples
{
    /// <summary>
    /// 数据行数验证示例
    /// 演示当Excel文件的数据行数不一致时的处理逻辑
    /// </summary>
    public class DataRowCountValidationExample
    {
        private readonly IExcelService _excelService;

        public DataRowCountValidationExample()
        {
            _excelService = new ExcelService();
        }

        /// <summary>
        /// 演示数据行数验证功能
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        public void DemonstrateRowCountValidation(string filePath)
        {
            Console.WriteLine("=== 数据行数验证示例 ===");
            Console.WriteLine($"正在读取文件: {Path.GetFileName(filePath)}");
            Console.WriteLine();

            try
            {
                // 读取Excel文件，如果数据行数不一致会提示用户
                var excelFile = _excelService.ReadExcelFile(filePath);
                
                Console.WriteLine($"文件读取成功！");
                Console.WriteLine($"实际读取的数据行数: {excelFile.DataRows.Count}");
                Console.WriteLine($"预期数据行数: 364 (B5-B368行)");
                Console.WriteLine();

                if (excelFile.DataRows.Count == 364)
                {
                    Console.WriteLine("✅ 数据行数符合预期");
                }
                else
                {
                    Console.WriteLine("⚠️  数据行数不一致，但用户选择继续处理");
                }

                // 显示前几行数据作为示例
                Console.WriteLine("\n前5行数据示例:");
                for (int i = 0; i < Math.Min(5, excelFile.DataRows.Count); i++)
                {
                    var row = excelFile.DataRows[i];
                    Console.WriteLine($"第{row.RowIndex}行: {row.Name} - 数据值: [{string.Join(", ", row.Values.Select(v => v?.ToString() ?? "null"))}]");
                }
            }
            catch (WorkPartnerException ex) when (ex.Category == "UserCancelled")
            {
                Console.WriteLine("❌ 用户取消了处理操作");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 读取文件时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量验证多个文件的数据行数
        /// </summary>
        /// <param name="directoryPath">包含Excel文件的目录路径</param>
        public void BatchValidateRowCounts(string directoryPath)
        {
            Console.WriteLine("=== 批量数据行数验证 ===");
            Console.WriteLine($"正在扫描目录: {directoryPath}");
            Console.WriteLine();

            var excelFiles = Directory.GetFiles(directoryPath, "*.xls")
                .Concat(Directory.GetFiles(directoryPath, "*.xlsx"))
                .ToList();

            Console.WriteLine($"找到 {excelFiles.Count} 个Excel文件");
            Console.WriteLine();

            var validationResults = new List<(string fileName, int rowCount, bool isExpected)>();

            foreach (var filePath in excelFiles)
            {
                try
                {
                    var excelFile = _excelService.ReadExcelFile(filePath);
                    var isExpected = excelFile.DataRows.Count == 364;
                    validationResults.Add((Path.GetFileName(filePath), excelFile.DataRows.Count, isExpected));
                    
                    Console.WriteLine($"✅ {Path.GetFileName(filePath)}: {excelFile.DataRows.Count} 行");
                }
                catch (WorkPartnerException ex) when (ex.Category == "UserCancelled")
                {
                    Console.WriteLine($"❌ {Path.GetFileName(filePath)}: 用户取消");
                    validationResults.Add((Path.GetFileName(filePath), 0, false));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ {Path.GetFileName(filePath)}: 错误 - {ex.Message}");
                    validationResults.Add((Path.GetFileName(filePath), 0, false));
                }
            }

            // 显示验证统计
            Console.WriteLine();
            Console.WriteLine("=== 验证统计 ===");
            var expectedCount = validationResults.Count(r => r.isExpected);
            var unexpectedCount = validationResults.Count(r => !r.isExpected);
            
            Console.WriteLine($"符合预期的文件: {expectedCount}");
            Console.WriteLine($"不符合预期的文件: {unexpectedCount}");
            Console.WriteLine($"总文件数: {validationResults.Count}");
        }
    }
} 