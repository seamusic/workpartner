using DataFixter.Excel;
using DataFixter.Logging;
using DataFixter.Services;
using Serilog;

namespace DataFixter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 配置日志系统
                LoggingConfiguration.ConfigureLogging();
                
                Log.Information("=== DataFixter 数据修正工具启动 ===");
                Log.Information("版本: 1.0.0");
                Log.Information("目标框架: .NET 8.0");
                
                // 检查命令行参数
                if (args.Length != 2)
                {
                    ShowUsage();
                    return;
                }

                var processedDirectory = args[0];
                var referenceDirectory = args[1];
                
                Log.Information("待处理数据目录: {ProcessedDir}", processedDirectory);
                Log.Information("对比参考目录: {ReferenceDir}", referenceDirectory);
                
                // 验证目录是否存在
                if (!Directory.Exists(processedDirectory))
                {
                    Log.Error("待处理数据目录不存在: {ProcessedDir}", processedDirectory);
                    Console.WriteLine($"错误: 待处理数据目录不存在: {processedDirectory}");
                    return;
                }
                
                if (!Directory.Exists(referenceDirectory))
                {
                    Log.Error("对比参考目录不存在: {ReferenceDir}", referenceDirectory);
                    Console.WriteLine($"错误: 对比参考目录不存在: {referenceDirectory}");
                    return;
                }
                
                // 批量处理两个目录中的所有Excel文件
                ProcessDirectories(processedDirectory, referenceDirectory);
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "程序执行过程中发生严重错误");
            }
            finally
            {
                Log.Information("=== DataFixter 数据修正工具退出 ===");
                LoggingConfiguration.CloseLogging();
            }
        }

        /// <summary>
        /// 显示使用说明
        /// </summary>
        private static void ShowUsage()
        {
            Console.WriteLine("DataFixter - Excel数据批量修正工具");
            Console.WriteLine();
            Console.WriteLine("使用方法:");
            Console.WriteLine("  DataFixter <待处理数据目录> <对比参考目录>");
            Console.WriteLine();
            Console.WriteLine("参数说明:");
            Console.WriteLine("  待处理数据目录: 包含需要修正的Excel文件的目录路径");
            Console.WriteLine("  对比参考目录: 包含对比数据的Excel文件的目录路径");
            Console.WriteLine();
            Console.WriteLine("示例:");
            Console.WriteLine("  DataFixter \"E:\\workspace\\gmdi\\tools\\WorkPartner\\excel\\processed\" \"E:\\workspace\\gmdi\\tools\\WorkPartner\\excel\"");
            Console.WriteLine();
            Console.WriteLine("功能说明:");
            Console.WriteLine("  工具将批量处理待处理目录中的所有Excel文件");
            Console.WriteLine("  使用对比参考目录中的数据进行验证和参考");
            Console.WriteLine("  自动修正数据逻辑问题并生成修正后的文件");
        }

        /// <summary>
        /// 批量处理两个目录中的所有Excel文件
        /// </summary>
        /// <param name="processedDirectory">待处理数据目录</param>
        /// <param name="referenceDirectory">对比参考目录</param>
        private static void ProcessDirectories(string processedDirectory, string referenceDirectory)
        {
            try
            {
                Log.Information("开始批量处理目录: {ProcessedDir} 和 {ReferenceDir}", processedDirectory, referenceDirectory);
                
                // 获取两个目录中的所有Excel文件
                var processedFiles = GetExcelFiles(processedDirectory);
                var referenceFiles = GetExcelFiles(referenceDirectory);
                
                Log.Information("找到待处理文件: {ProcessedCount}个", processedFiles.Count);
                Log.Information("找到参考文件: {ReferenceCount}个", referenceFiles.Count);
                
                if (processedFiles.Count == 0)
                {
                    Log.Warning("待处理目录中没有找到Excel文件");
                    Console.WriteLine("警告: 待处理目录中没有找到Excel文件");
                    return;
                }
                
                if (referenceFiles.Count == 0)
                {
                    Log.Warning("参考目录中没有找到Excel文件");
                    Console.WriteLine("警告: 参考目录中没有找到Excel文件");
                    return;
                }
                
                // 显示文件列表
                Console.WriteLine($"\n=== 待处理文件列表 ===");
                foreach (var file in processedFiles)
                {
                    Console.WriteLine($"  {Path.GetFileName(file)}");
                }
                
                Console.WriteLine($"\n=== 参考文件列表 ===");
                foreach (var file in referenceFiles)
                {
                    Console.WriteLine($"  {Path.GetFileName(file)}");
                }
                
                // 开始批量处理
                Console.WriteLine($"\n=== 开始批量处理 ===");
                var totalFiles = processedFiles.Count;
                var processedCount = 0;
                var successCount = 0;
                var errorCount = 0;
                
                foreach (var filePath in processedFiles)
                {
                    processedCount++;
                    var fileName = Path.GetFileName(filePath);
                    
                    Console.WriteLine($"\n[{processedCount}/{totalFiles}] 处理文件: {fileName}");
                    Log.Information("开始处理文件 [{Current}/{Total}]: {FileName}", processedCount, totalFiles, fileName);
                    
                    try
                    {
                        // 处理单个文件
                        if (ProcessSingleFile(filePath, referenceFiles))
                        {
                            successCount++;
                            Console.WriteLine($"  ✓ 文件处理成功: {fileName}");
                        }
                        else
                        {
                            errorCount++;
                            Console.WriteLine($"  ✗ 文件处理失败: {fileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        Log.Error(ex, "处理文件时发生错误: {FileName}", fileName);
                        Console.WriteLine($"  ✗ 文件处理异常: {fileName} - {ex.Message}");
                    }
                }
                
                // 输出处理结果汇总
                Console.WriteLine($"\n=== 批量处理完成 ===");
                Console.WriteLine($"总文件数: {totalFiles}");
                Console.WriteLine($"成功处理: {successCount}");
                Console.WriteLine($"处理失败: {errorCount}");
                Console.WriteLine($"成功率: {(double)successCount / totalFiles * 100:F1}%");
                
                Log.Information("批量处理完成: 总计{Total}, 成功{Success}, 失败{Failed}", 
                    totalFiles, successCount, errorCount);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "批量处理目录时发生错误");
                Console.WriteLine($"错误: 批量处理目录时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取目录中的所有Excel文件
        /// </summary>
        /// <param name="directory">目录路径</param>
        /// <returns>Excel文件路径列表</returns>
        private static List<string> GetExcelFiles(string directory)
        {
            var excelFiles = new List<string>();
            
            try
            {
                // 支持.xls和.xlsx格式
                var extensions = new[] { "*.xls", "*.xlsx" };
                
                foreach (var extension in extensions)
                {
                    var files = Directory.GetFiles(directory, extension, SearchOption.TopDirectoryOnly);
                    excelFiles.AddRange(files);
                }
                
                // 按文件名排序
                excelFiles.Sort();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取目录文件列表失败: {Directory}", directory);
            }
            
            return excelFiles;
        }

        /// <summary>
        /// 处理单个Excel文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="referenceFiles">参考文件列表</param>
        /// <returns>是否处理成功</returns>
        private static bool ProcessSingleFile(string filePath, List<string> referenceFiles)
        {
            try
            {
                using var excelProcessor = new ExcelProcessor(filePath);
                if (!excelProcessor.OpenFile())
                {
                    Log.Error("无法打开Excel文件: {FilePath}", filePath);
                    return false;
                }

                var dataService = new DataProcessingService(excelProcessor);
                
                // 获取文件的行数信息
                var rowCount = excelProcessor.GetRowCount();
                var columnCount = excelProcessor.GetColumnCount();
                
                Log.Information("文件信息: 行数{RowCount}, 列数{ColumnCount}", rowCount, columnCount);
                
                // 根据实际需求调整数据范围（这里假设从第5行开始，到第364行结束）
                var startRow = 4; // 第5行（索引从0开始）
                var endRow = Math.Min(363, rowCount - 1); // 第364行或文件末尾
                
                if (endRow < startRow)
                {
                    Log.Warning("文件行数不足，无法处理: {FilePath}, 行数: {RowCount}", filePath, rowCount);
                    return false;
                }
                
                // 验证数据完整性
                var requiredColumns = new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 }; // A-L列
                var validationResult = dataService.ValidateData(startRow, endRow, requiredColumns);
                
                // 输出验证结果
                Console.WriteLine($"    数据验证: 总行{validationResult.TotalRows}, 有效{validationResult.ValidRows}, 无效{validationResult.InvalidRows}");
                
                if (validationResult.MissingDataRows.Count > 0)
                {
                    Console.WriteLine($"    缺失数据行: {string.Join(", ", validationResult.MissingDataRows)}");
                }
                
                if (validationResult.InvalidDataRows.Count > 0)
                {
                    Console.WriteLine($"    无效数据行: {string.Join(", ", validationResult.InvalidDataRows)}");
                }
                
                // 如果有无效数据，尝试修正
                if (validationResult.InvalidRows > 0)
                {
                    Console.WriteLine("    开始修正数据...");
                    
                    // 设置默认值
                    var defaultValues = new Dictionary<int, string>
                    {
                        { 0, "0" },           // A列序号
                        { 1, "默认点名" },     // B列点名
                        { 2, "0" },           // C列里程
                        { 3, "0" },           // D列本期变化量X
                        { 4, "0" },           // E列本期变化量Y
                        { 5, "0" },           // F列本期变化量Z
                        { 6, "0" },           // G列累计变化量X
                        { 7, "0" },           // H列累计变化量Y
                        { 8, "0" },           // I列累计变化量Z
                        { 9, "0" },           // J列日变化量X
                        { 10, "0" },          // K列日变化量Y
                        { 11, "0" }           // L列日变化量Z
                    };
                    
                    // 修正缺失数据
                    var correctionResult = dataService.CorrectMissingData(
                        validationResult.MissingDataRows, 
                        requiredColumns.ToList(), 
                        defaultValues);
                    
                    Console.WriteLine($"    修正完成: 总计{correctionResult.TotalCorrections}, 成功{correctionResult.SuccessfulCorrections}, 失败{correctionResult.FailedCorrections}");
                    
                    // 保存修正后的文件
                    var outputPath = Path.Combine(
                        Path.GetDirectoryName(filePath) ?? ".",
                        Path.GetFileNameWithoutExtension(filePath) + "_corrected" + Path.GetExtension(filePath));
                    
                    if (excelProcessor.SaveFile(outputPath))
                    {
                        Log.Information("修正后的文件已保存: {OutputPath}", outputPath);
                        Console.WriteLine($"    修正文件已保存: {Path.GetFileName(outputPath)}");
                    }
                    else
                    {
                        Log.Error("保存修正后的文件失败: {FilePath}", filePath);
                        Console.WriteLine("    错误: 保存修正文件失败");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine("    数据完整，无需修正");
                }
                
                // 生成数据统计报告
                var statisticsReport = dataService.GenerateStatistics(startRow, endRow);
                Console.WriteLine($"    数据统计: 总行{statisticsReport.TotalRows}, 总列{statisticsReport.ColumnStatistics.Count}");
                
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "处理文件失败: {FilePath}", filePath);
                return false;
            }
        }
    }
}
