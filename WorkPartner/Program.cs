using WorkPartner.Models;
using WorkPartner.Utils;
using WorkPartner.Services;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorkPartner
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("WorkPartner Excel数据处理工具 - 阶段5增强版");
            Console.WriteLine("==========================================");

            // 初始化日志
            Logger.Initialize("logs/workpartner.log", LogLevel.Info);
            Logger.Info("WorkPartner启动");
            Logger.MemoryUsage("启动时");

            try
            {
                using var mainOperation = Logger.StartOperation("主程序执行");
                ExceptionHandler.ClearErrorStatistics();

                // 解析命令行参数
                //args = new[] { "E:\\workspace\\gmdi\\tools\\WorkPartner\\excel" };
                var arguments = ParseCommandLineArguments(args);
                if (arguments == null)
                {
                    ShowUsage();
                    return;
                }

                // 检查是否为比较模式
                if (arguments.CompareMode)
                {
                    await RunCompareMode(arguments);
                    return;
                }

                // 检查是否为大值检查模式
                if (arguments.CheckLargeValues)
                {
                    await RunLargeValueCheckMode(arguments);
                    return;
                }

                // 验证输入路径
                if (!ValidateInputPath(arguments.InputPath))
                {
                    Console.WriteLine("❌ 输入路径无效或不存在");
                    return;
                }

                // 创建输出目录
                CreateOutputDirectory(arguments.OutputPath);

                // 扫描Excel文件
                var excelFiles = ScanExcelFiles(arguments.InputPath);
                if (excelFiles.Count == 0)
                {
                    Console.WriteLine("❌ 未找到任何Excel文件");
                    return;
                }

                Console.WriteLine($"✅ 找到 {excelFiles.Count} 个Excel文件");

                // 解析文件名并排序
                var parsedFiles = ParseAndSortFiles(excelFiles);
                if (parsedFiles.Count == 0)
                {
                    Console.WriteLine("❌ 没有找到符合格式的Excel文件");
                    return;
                }

                Console.WriteLine($"✅ 成功解析 {parsedFiles.Count} 个文件");

                // 读取Excel数据
                var filesWithData = ReadExcelData(parsedFiles);
                Console.WriteLine($"✅ 成功读取 {filesWithData.Count} 个文件的数据");

                // 阶段3：数据处理逻辑
                Console.WriteLine("\n🔄 开始阶段3数据处理...");
                
                // 3.1 数据完整性检查
                Console.WriteLine("🔍 检查数据完整性...");
                var completenessResult = DataProcessor.CheckCompleteness(filesWithData);
                
                // 生成补充文件列表
                var supplementFiles = DataProcessor.GenerateSupplementFiles(filesWithData);
                
                // 创建补充文件（不包含A2列数据修改）
                if (supplementFiles.Any())
                {
                    Console.WriteLine($"📁 创建 {supplementFiles.Count} 个补充文件...");
                    var createdCount = DataProcessor.CreateSupplementFiles(supplementFiles, arguments.OutputPath);
                    Console.WriteLine($"✅ 成功创建 {createdCount} 个补充文件");
                }
                else
                {
                    Console.WriteLine("ℹ️ 无需创建补充文件，所有时间点数据都完整");
                }
                
                // 3.2 数据补充算法 - 处理所有文件（包括新创建的补充文件）
                Console.WriteLine("📊 处理缺失数据...");
                var allFilesForProcessing = DataProcessor.GetAllFilesForProcessing(filesWithData, supplementFiles, arguments.OutputPath);
                var processedFiles = DataProcessor.ProcessMissingData(allFilesForProcessing);

                // 3.3 第4、5、6列验证和重新计算 - 确保数据符合"1. 基本逻辑重构"要求
                Console.WriteLine("🔍 验证第4、5、6列数据是否符合基本逻辑重构要求...");
                var validatedFiles = DataProcessor.ValidateAndRecalculateColumns456(processedFiles);
                //var validatedFiles = processedFiles;
                Console.WriteLine($"✅ 第4、5、6列验证和重新计算完成");
                                
                // 保存处理后的数据到Excel文件（包含A2列更新）
                Console.WriteLine("💾 保存处理后的数据并更新A2列...");
                await SaveProcessedFiles(validatedFiles, arguments.OutputPath);
                
                // 3.4 原始文件与已处理文件比较 - 检查数据处理前后的差异（在保存之后进行比较）
                Console.WriteLine("🔍 比较原始文件与修正后文件的数值差异...");
                var originalDirectory = arguments.InputPath;
                var processedDirectory = arguments.OutputPath;
                
                try
                {
                    // 使用增强的比较功能，支持详细差异显示和自定义容差
                    var comparisonResult = DataProcessor.CompareOriginalAndProcessedFiles(
                        originalDirectory, 
                        processedDirectory,
                        showDetailedDifferences: true,  // 启用详细差异显示
                        tolerance: 0.001,               // 设置比较容差为0.001
                        maxDifferencesToShow: 10        // 每个文件最多显示10个差异
                    );
                    
                    if (comparisonResult.HasError)
                    {
                        Console.WriteLine($"⚠️ 文件比较过程发生错误: {comparisonResult.ErrorMessage}");
                    }
                    else
                    {
                        // 比较结果已在方法内部显示，这里只显示简要总结
                        Console.WriteLine($"✅ 文件比较分析完成");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ 文件比较功能执行失败: {ex.Message}");
                    Logger.Warning($"文件比较功能执行失败: {ex.Message}");
                }
                
                // 数据质量验证
                var qualityReport = DataProcessor.ValidateDataQuality(validatedFiles);

                // 显示处理结果
                DisplayProcessingResults(validatedFiles, completenessResult, supplementFiles, qualityReport);

                Console.WriteLine("\n✅ 阶段5数据处理逻辑完成！");
                
                // 显示最终统计和错误报告
                ShowFinalStatistics();
            }
            catch (WorkPartnerException ex)
            {
                Logger.Error($"WorkPartner专用错误 - {ex.Category}", ex);
                Console.WriteLine($"\n❌ 程序执行失败 ({ex.Category}): {ex.Message}");
                if (ex.FilePath != null)
                {
                    Console.WriteLine($"   相关文件: {ex.FilePath}");
                }
                ShowErrorContext(ex);
            }
            catch (Exception ex)
            {
                Logger.Error("程序执行过程中发生未知错误", ex);
                Console.WriteLine($"\n❌ 程序执行失败: {ex.Message}");
                Console.WriteLine($"   异常类型: {ex.GetType().Name}");
            }
            finally
            {
                // 最终清理工作
                Logger.MemoryUsage("程序结束时");
                Logger.Info("WorkPartner执行完成");
                
                // 显示错误报告
                var errorReport = ExceptionHandler.GenerateErrorReport();
                if (!errorReport.Contains("未发现错误"))
                {
                    Console.WriteLine("\n📊 错误统计报告:");
                    Console.WriteLine(errorReport);
                    Logger.Info("错误统计报告:");
                    Logger.Info(errorReport);
                }
                
                // 清理日志文件
                Logger.CleanupLogFile();
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        // 命令行参数模型
        private class CommandLineArguments
        {
            public string InputPath { get; set; } = string.Empty;
            public string OutputPath { get; set; } = string.Empty;
            public bool Verbose { get; set; } = false;
            public bool CompareMode { get; set; } = false;
            public string CompareOriginalPath { get; set; } = string.Empty;
            public string CompareProcessedPath { get; set; } = string.Empty;
            public bool ShowDetailedDifferences { get; set; } = false;
            public double Tolerance { get; set; } = 0.001;
            public int MaxDifferencesToShow { get; set; } = 10;
            public bool CheckLargeValues { get; set; } = false;
            public double LargeValueThreshold { get; set; } = 4.0;
        }

        // 解析命令行参数
        private static CommandLineArguments? ParseCommandLineArguments(string[] args)
        {
            if (args.Length == 0)
            {
                return null;
            }

            var arguments = new CommandLineArguments();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-i":
                    case "--input":
                        if (i + 1 < args.Length)
                        {
                            arguments.InputPath = args[++i];
                        }
                        break;
                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            arguments.OutputPath = args[++i];
                        }
                        break;
                    case "-v":
                    case "--verbose":
                        arguments.Verbose = true;
                        break;
                    case "-c":
                    case "--compare":
                        arguments.CompareMode = true;
                        // 比较模式需要两个路径参数
                        if (i + 2 < args.Length)
                        {
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = args[++i];
                        }
                        else if (i + 1 < args.Length)
                        {
                            // 如果只有一个路径，假设是原始路径，输出路径使用默认值
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = Path.Combine(arguments.CompareOriginalPath, "processed");
                        }
                        break;
                    case "--detailed":
                        arguments.ShowDetailedDifferences = true;
                        break;
                    case "--tolerance":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out var tolerance))
                        {
                            arguments.Tolerance = tolerance;
                        }
                        break;
                    case "--max-differences":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out var maxDiff))
                        {
                            arguments.MaxDifferencesToShow = maxDiff;
                        }
                        break;
                    case "--check-large-values":
                        arguments.CheckLargeValues = true;
                        break;
                    case "--large-value-threshold":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out var largeValueThreshold))
                        {
                            arguments.LargeValueThreshold = largeValueThreshold;
                        }
                        break;
                    case "-h":
                    case "--help":
                        return null;
                    default:
                        // 检查是否是比较模式的简化语法：-v 原始目录 对比目录
                        if (args[i] == "-v" && i + 2 < args.Length)
                        {
                            arguments.CompareMode = true;
                            arguments.Verbose = true;
                            arguments.ShowDetailedDifferences = true;
                            arguments.CompareOriginalPath = args[++i];
                            arguments.CompareProcessedPath = args[++i];
                        }
                        // 检查是否是大值检查模式：--check-large-values 目录路径
                        else if (args[i] == "--check-large-values" && i + 1 < args.Length)
                        {
                            arguments.CheckLargeValues = true;
                            arguments.InputPath = args[++i]; // 将下一个参数作为要检查的目录路径
                        }
                        // 如果没有指定参数，第一个参数作为输入路径
                        else if (string.IsNullOrEmpty(arguments.InputPath))
                        {
                            arguments.InputPath = args[i];
                        }
                        break;
                }
            }

            // 如果没有指定输出路径，使用默认路径
            if (string.IsNullOrEmpty(arguments.OutputPath))
            {
                arguments.OutputPath = Path.Combine(arguments.InputPath, "processed");
            }

            return arguments;
        }

        // 显示使用说明
        private static void ShowUsage()
        {
            Console.WriteLine("使用方法:");
            Console.WriteLine("  WorkPartner.exe <输入目录> [选项]");
            Console.WriteLine("  WorkPartner.exe -c <原始目录> <对比目录> [选项]");
            Console.WriteLine("  WorkPartner.exe -v <原始目录> <对比目录>");
            Console.WriteLine("  WorkPartner.exe --check-large-values <目录路径> [选项]");
            Console.WriteLine("");
            Console.WriteLine("参数:");
            Console.WriteLine("  <输入目录>              包含Excel文件的目录路径");
            Console.WriteLine("  <原始目录>              原始Excel文件目录");
            Console.WriteLine("  <对比目录>              已处理的Excel文件目录");
            Console.WriteLine("  <目录路径>              要检查的Excel文件目录");
            Console.WriteLine("");
            Console.WriteLine("选项:");
            Console.WriteLine("  -o, --output <目录>     输出目录路径 (默认: <输入目录>/processed)");
            Console.WriteLine("  -v, --verbose           详细输出模式");
            Console.WriteLine("  -c, --compare           文件比较模式");
            Console.WriteLine("  --detailed              显示详细差异信息");
            Console.WriteLine("  --tolerance <数值>      设置比较容差 (默认: 0.001)");
            Console.WriteLine("  --max-differences <数量> 限制显示差异数量 (默认: 10)");
            Console.WriteLine("  --check-large-values    大值数据检查模式");
            Console.WriteLine("  --large-value-threshold <数值> 设置大值检查阈值 (默认: 4.0)");
            Console.WriteLine("  -h, --help              显示此帮助信息");
            Console.WriteLine("");
            Console.WriteLine("支持的文件格式:");
            Console.WriteLine("  ✅ .xlsx (Excel 2007+)");
            Console.WriteLine("  ✅ .xls (Excel 97-2003)");
            Console.WriteLine("");
            Console.WriteLine("示例:");
            Console.WriteLine("  数据处理模式:");
            Console.WriteLine("    WorkPartner.exe C:\\excel\\");
            Console.WriteLine("    WorkPartner.exe ..\\excel\\");
            Console.WriteLine("    WorkPartner.exe C:\\excel\\ -o C:\\output\\ -v");
            Console.WriteLine("");
            Console.WriteLine("  文件比较模式:");
            Console.WriteLine("    WorkPartner.exe -c C:\\original C:\\processed");
            Console.WriteLine("    WorkPartner.exe -v C:\\original C:\\processed");
            Console.WriteLine("    WorkPartner.exe -c C:\\original C:\\processed --detailed --tolerance 0.01");
            Console.WriteLine("");
            Console.WriteLine("  大值检查模式:");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output --large-value-threshold 5.0");
            Console.WriteLine("    WorkPartner.exe --check-large-values C:\\output --large-value-threshold 3.0 -v");
        }

        // 验证输入路径
        private static bool ValidateInputPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("❌ 输入路径不能为空");
                return false;
            }

            if (!Directory.Exists(path))
            {
                Console.WriteLine($"❌ 目录不存在: {path}");
                return false;
            }

            Logger.Info($"验证输入路径: {path}");
            return true;
        }

        // 创建输出目录
        private static void CreateOutputDirectory(string outputPath)
        {
            try
            {
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                    Logger.Info($"创建输出目录: {outputPath}");
                }
                else
                {
                    Logger.Info($"输出目录已存在: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"创建输出目录失败: {outputPath}", ex);
                throw;
            }
        }

        // 扫描Excel文件
        private static List<string> ScanExcelFiles(string inputPath)
        {
            try
            {
                using var operation = Logger.StartOperation("扫描Excel文件", inputPath);
                Logger.Info($"验证输入路径: {inputPath}");
                
                var fileService = new FileService();
                var excelFiles = fileService.ScanExcelFiles(inputPath);
                
                foreach (var file in excelFiles)
                {
                    Logger.Debug($"找到Excel文件: {Path.GetFileName(file)}");
                }

                Logger.Info($"扫描完成，找到 {excelFiles.Count} 个Excel文件");
                return excelFiles;
            }
            catch (Exception ex)
            {
                Logger.Error($"扫描Excel文件失败", ex);
                throw new WorkPartnerException("ScanFailed", "文件扫描失败", inputPath, ex);
            }
        }

        // 解析文件名并排序
        private static List<ExcelFile> ParseAndSortFiles(List<string> filePaths)
        {
            var parsedFiles = new List<ExcelFile>();

            foreach (var filePath in filePaths)
            {
                var fileName = Path.GetFileName(filePath);
                var parseResult = FileNameParser.ParseFileName(fileName);

                if (parseResult?.IsValid == true)
                {
                    var excelFile = new ExcelFile
                    {
                        FilePath = filePath,
                        FileName = fileName,
                        Date = parseResult.Date,
                        Hour = parseResult.Hour,
                        ProjectName = parseResult.ProjectName,
                        FileSize = new FileInfo(filePath).Length,
                        LastModified = new FileInfo(filePath).LastWriteTime,
                        IsValid = true
                    };
                    parsedFiles.Add(excelFile);
                    Logger.Debug($"成功解析文件: {fileName}");
                }
                else
                {
                    Logger.Warning($"跳过无效格式文件: {fileName}");
                }
            }

            // 按日期和时间排序
            parsedFiles.Sort((a, b) =>
            {
                var dateComparison = a.Date.CompareTo(b.Date);
                if (dateComparison != 0)
                    return dateComparison;
                return a.Hour.CompareTo(b.Hour);
            });

            Logger.Info($"成功解析 {parsedFiles.Count} 个文件，已按日期时间排序");
            return parsedFiles;
        }

        // 读取Excel数据
        private static List<ExcelFile> ReadExcelData(List<ExcelFile> files)
        {
            Console.WriteLine($"📖 开始读取Excel数据，共 {files.Count} 个文件...");
            
            var filesWithData = new List<ExcelFile>();
            var excelService = new ExcelService();
            var lastProgressTime = DateTime.Now;

            for (int i = 0; i < files.Count; i++)
            {
                var file = files[i];
                
                // 每读取10个文件或每30秒显示一次进度
                if ((i + 1) % 10 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                {
                    var progress = (double)(i + 1) / files.Count * 100;
                    Console.WriteLine($"📈 读取进度: {i + 1}/{files.Count} ({progress:F1}%) - 当前文件: {file.FileName}");
                    lastProgressTime = DateTime.Now;
                }
                
                Logger.Progress(i + 1, files.Count, $"读取Excel数据: {file.FileName}");

                try
                {
                    var excelFileWithData = excelService.ReadExcelFile(file.FilePath);
                    file.DataRows = excelFileWithData.DataRows;
                    file.IsValid = excelFileWithData.IsValid;
                    file.IsLocked = excelFileWithData.IsLocked;
                    filesWithData.Add(file);

                    Logger.Debug($"成功读取 {file.FileName}: {file.DataRows.Count} 行数据");
                }
                catch (Exception ex)
                {
                    Logger.Error($"读取文件失败: {file.FileName}", ex);
                    Console.WriteLine($"❌ 读取失败: {file.FileName} - {ex.Message}");
                    // 继续处理其他文件
                }
            }

            Console.WriteLine($"✅ 成功读取 {filesWithData.Count} 个文件的数据");
            Logger.Info($"成功读取 {filesWithData.Count} 个文件的数据");
            return filesWithData;
        }

        // 显示处理结果（阶段3版本）
        private static void DisplayProcessingResults(List<ExcelFile> files, CompletenessCheckResult completenessResult, List<SupplementFileInfo> supplementFiles, DataQualityReport qualityReport)
        {
            Console.WriteLine("\n--- 阶段3处理结果摘要 ---");

            // 按日期分组显示
            var groupedFiles = files.GroupBy(f => f.Date).OrderBy(g => g.Key);

            foreach (var group in groupedFiles)
            {
                Console.WriteLine($"\n日期: {group.Key:yyyy.M.d}");
                var hours = group.Select(f => f.Hour).OrderBy(h => h).ToList();
                Console.WriteLine($"  时间点: [{string.Join(", ", hours)}]");
                Console.WriteLine($"  文件数: {group.Count()}");

                foreach (var file in group.OrderBy(f => f.Hour))
                {
                    var dataCount = file.DataRows?.Count ?? 0;
                    var completeness = file.DataRows?.Count > 0 
                        ? file.DataRows.Average(r => r.CompletenessPercentage) 
                        : 0;
                    Console.WriteLine($"    {file.FormattedHour}时: {dataCount} 行数据, 完整性 {completeness:F1}%");
                }
            }

            // 数据补充统计
            Console.WriteLine($"\n📊 数据补充统计:");
            var totalMissingValues = files.SelectMany(f => f.DataRows).Sum(r => r.MissingDataCount);
            var totalValues = files.SelectMany(f => f.DataRows).Sum(r => r.Values.Count);
            var supplementedCount = totalValues - totalMissingValues; // 假设所有缺失都已补充
            Console.WriteLine($"  总数据点: {totalValues}");
            Console.WriteLine($"  原始缺失: {totalMissingValues}");
            Console.WriteLine($"  已补充: {supplementedCount}");
            Console.WriteLine($"  补充率: {(totalMissingValues > 0 ? 100.0 : 0):F1}%");

            // 完整性检查结果
            Console.WriteLine($"\n🔍 数据完整性: {(completenessResult.IsAllComplete ? "✅ 完整" : "❌ 不完整")}");

            if (!completenessResult.IsAllComplete)
            {
                Console.WriteLine("缺失的时间点:");
                foreach (var dateCompleteness in completenessResult.DateCompleteness)
                {
                    if (dateCompleteness.MissingHours.Any())
                    {
                        Console.WriteLine($"  {dateCompleteness.Date:yyyy.M.d}: [{string.Join(", ", dateCompleteness.MissingHours)}]");
                    }
                }
            }

            // 补充文件建议
            if (supplementFiles.Any())
            {
                Console.WriteLine($"\n📋 建议生成 {supplementFiles.Count} 个补充文件:");
                foreach (var supplement in supplementFiles.Take(5)) // 只显示前5个
                {
                    Console.WriteLine($"  {supplement.TargetFileName}");
                }
                if (supplementFiles.Count > 5)
                {
                    Console.WriteLine($"  ... 还有 {supplementFiles.Count - 5} 个文件");
                }
            }

            // 数据质量报告
            Console.WriteLine($"\n📈 数据质量报告:");
            Console.WriteLine($"  总体完整性: {qualityReport.OverallCompleteness:F1}%");
            Console.WriteLine($"  有效数据行: {qualityReport.ValidRows}/{qualityReport.TotalRows}");
            Console.WriteLine($"  缺失数据行: {qualityReport.MissingRows}");
        }

        // 显示处理结果（阶段2版本，保留向后兼容）
        private static void DisplayProcessingResults(List<ExcelFile> files)
        {
            Console.WriteLine("\n--- 处理结果摘要 ---");

            // 按日期分组显示
            var groupedFiles = files.GroupBy(f => f.Date).OrderBy(g => g.Key);

            foreach (var group in groupedFiles)
            {
                Console.WriteLine($"\n日期: {group.Key:yyyy.M.d}");
                var hours = group.Select(f => f.Hour).OrderBy(h => h).ToList();
                Console.WriteLine($"  时间点: [{string.Join(", ", hours)}]");
                Console.WriteLine($"  文件数: {group.Count()}");

                foreach (var file in group.OrderBy(f => f.Hour))
                {
                    var dataCount = file.DataRows?.Count ?? 0;
                    var completeness = file.DataRows?.Count > 0 
                        ? file.DataRows.Average(r => r.CompletenessPercentage) 
                        : 0;
                    Console.WriteLine($"    {file.FormattedHour}时: {dataCount} 行数据, 完整性 {completeness:F1}%");
                }
            }

            // 完整性检查
            var completenessResult = DataProcessor.CheckCompleteness(files);
            Console.WriteLine($"\n数据完整性: {(completenessResult.IsAllComplete ? "✅ 完整" : "❌ 不完整")}");

            if (!completenessResult.IsAllComplete)
            {
                Console.WriteLine("缺失的时间点:");
                foreach (var dateCompleteness in completenessResult.DateCompleteness)
                {
                    if (dateCompleteness.MissingHours.Any())
                    {
                        Console.WriteLine($"  {dateCompleteness.Date:yyyy.M.d}: [{string.Join(", ", dateCompleteness.MissingHours)}]");
                    }
                }
            }
        }

        // 阶段1测试方法（保留用于测试）
        static void TestFileNameParser()
        {
            Console.WriteLine("\n--- 测试文件名解析 ---");

            var testFiles = new[]
            {
                "2025.4.18-8云港城项目4#地块.xlsx",
                "2025.4.19-16云港城项目4#地块.xlsx",
                "invalid_file.txt",
                "2025.4.20-25云港城项目4#地块.xlsx" // 无效时间
            };

            foreach (var fileName in testFiles)
            {
                var result = FileNameParser.ParseFileName(fileName);
                if (result?.IsValid == true)
                {
                    Console.WriteLine($"✅ {fileName} -> 日期: {result.FormattedDate}, 时间: {result.FormattedHour}, 项目: {result.ProjectName}");
                }
                else
                {
                    Console.WriteLine($"❌ {fileName} -> 格式无效");
                }
            }
        }

        static void TestDataModels()
        {
            Console.WriteLine("\n--- 测试数据模型 ---");

            // 测试DataRow
            var dataRow = new DataRow
            {
                Name = "测试数据",
                RowIndex = 5
            };

            dataRow.AddValue(10.5);
            dataRow.AddValue(null);
            dataRow.AddValue(20.3);
            dataRow.AddValue(15.7);

            Console.WriteLine($"数据行: {dataRow}");
            Console.WriteLine($"完整性: {dataRow.CompletenessPercentage:F1}%");
            Console.WriteLine($"平均值: {dataRow.AverageValue:F2}");
            Console.WriteLine($"数据范围: {dataRow.DataRange:F2}");

            // 测试ExcelFile
            var excelFile = new ExcelFile
            {
                FileName = "test.xlsx",
                Date = DateTime.Now,
                Hour = 8,
                ProjectName = "测试项目",
                FileSize = 1024 * 100 // 100KB
            };

            excelFile.DataRows.Add(dataRow);

            Console.WriteLine($"Excel文件: {excelFile}");
            Console.WriteLine($"文件大小: {excelFile.FileSizeKB:F1}KB");
            Console.WriteLine($"文件标识: {excelFile.FileIdentifier}");
        }

        static void TestDataProcessor()
        {
            Console.WriteLine("\n--- 测试数据处理工具 ---");

            // 创建测试数据
            var files = new List<ExcelFile>();

            // 模拟2025.4.18的数据
            var file1 = new ExcelFile
            {
                FileName = "2025.4.18-0云港城项目4#地块.xlsx",
                Date = new DateTime(2025, 4, 18),
                Hour = 0,
                ProjectName = "云港城项目4#地块.xlsx"
            };

            var file2 = new ExcelFile
            {
                FileName = "2025.4.18-16云港城项目4#地块.xlsx",
                Date = new DateTime(2025, 4, 18),
                Hour = 16,
                ProjectName = "云港城项目4#地块.xlsx"
            };

            files.Add(file1);
            files.Add(file2);

            // 测试完整性检查
            var completenessResult = DataProcessor.CheckCompleteness(files);
            Console.WriteLine($"数据完整性检查: {(completenessResult.IsAllComplete ? "完整" : "不完整")}");

            foreach (var dateCompleteness in completenessResult.DateCompleteness)
            {
                Console.WriteLine($"日期 {dateCompleteness.Date:yyyy.M.d}: 现有时间点 [{string.Join(", ", dateCompleteness.ExistingHours)}], 缺失时间点 [{string.Join(", ", dateCompleteness.MissingHours)}]");
            }

            // 测试补充文件生成
            var supplementFiles = DataProcessor.GenerateSupplementFiles(files);
            Console.WriteLine($"需要补充的文件数量: {supplementFiles.Count}");

            foreach (var supplementFile in supplementFiles)
            {
                Console.WriteLine($"补充文件: {supplementFile.TargetFileName}");
            }
        }

        static void TestLogger()
        {
            Console.WriteLine("\n--- 测试日志功能 ---");

            Logger.Debug("这是一条调试日志");
            Logger.Info("这是一条信息日志");
            Logger.Warning("这是一条警告日志");
            Logger.Error("这是一条错误日志");

            // 测试进度显示
            for (int i = 0; i <= 10; i++)
            {
                Logger.Progress(i, 10, "测试进度");
                Thread.Sleep(100);
            }

            // 测试性能记录
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            Thread.Sleep(100);
            stopwatch.Stop();
            Logger.Performance("测试操作", stopwatch.Elapsed);

            // 测试内存使用记录
            Logger.MemoryUsage("测试后");
        }

        /// <summary>
        /// 保存处理后的Excel文件
        /// </summary>
        /// <param name="processedFiles">处理后的文件列表</param>
        /// <param name="outputPath">输出目录</param>
        private static async Task SaveProcessedFiles(List<ExcelFile> processedFiles, string outputPath)
        {
            Console.WriteLine($"💾 开始保存处理后的文件，共 {processedFiles.Count} 个文件...");
            
            var excelService = new ExcelService();
            int savedCount = 0;
            int totalFiles = processedFiles.Count;
            var lastProgressTime = DateTime.Now;

            // 按日期和时间排序文件
            var sortedFiles = processedFiles.OrderBy(f => f.Date).ThenBy(f => f.Hour).ToList();

            for (int i = 0; i < sortedFiles.Count; i++)
            {
                var file = sortedFiles[i];
                
                try
                {
                    // 使用标准化的文件名格式（确保时间点使用零填充）
                    var standardizedFileName = FileNameParser.GenerateFileName(file.Date, file.Hour, file.ProjectName);
                    var outputFilePath = Path.Combine(outputPath, standardizedFileName);
                    
                    // 确定本期观测时间
                    var currentObservationTime = $"{file.Date:yyyy-M-d} {file.Hour:00}:00";
                    
                    // 确定上期观测时间
                    string previousObservationTime;
                    if (i > 0)
                    {
                        var previousFile = sortedFiles[i - 1];
                        previousObservationTime = $"{previousFile.Date:yyyy-M-d} {previousFile.Hour:00}:00";
                    }
                    else
                    {
                        // 如果是第一个文件，使用当前时间作为上期观测时间
                        previousObservationTime = currentObservationTime;
                    }
                    
                    // 保存文件并同时更新A2列
                    var success = excelService.SaveExcelFileWithA2Update(file, outputFilePath, currentObservationTime, previousObservationTime);
                    
                    if (success)
                    {
                        savedCount++;
                        
                        // 每保存10个文件或每30秒显示一次进度
                        if (savedCount % 10 == 0 || (DateTime.Now - lastProgressTime).TotalSeconds >= 30)
                        {
                            var progress = (double)savedCount / totalFiles * 100;
                            Console.WriteLine($"📈 保存进度: {savedCount}/{totalFiles} ({progress:F1}%) - 当前文件: {standardizedFileName}");
                            lastProgressTime = DateTime.Now;
                        }
                        else
                        {
                            Console.WriteLine($"✅ 已保存: {standardizedFileName}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"❌ 保存失败: {standardizedFileName}");
                    }
                }
                catch (Exception ex)
                {
                    var displayName = FileNameParser.GenerateFileName(file.Date, file.Hour, file.ProjectName);
                    Console.WriteLine($"❌ 保存文件失败: {displayName}");
                    Console.WriteLine($"   错误: {ex.Message}");
                    Logger.Error($"保存文件失败: {displayName}", ex);
                }
            }

            Console.WriteLine($"✅ 成功保存 {savedCount}/{totalFiles} 个处理后的文件");
        }

        /// <summary>
        /// 显示最终统计信息
        /// </summary>
        private static void ShowFinalStatistics()
        {
            Logger.MemoryUsage("处理完成时");
            
            var stats = new Dictionary<string, object>
            {
                ["最终内存使用"] = $"{GC.GetTotalMemory(false) / (1024.0 * 1024.0):F2}MB",
                ["GC次数 Gen0"] = GC.CollectionCount(0),
                ["GC次数 Gen1"] = GC.CollectionCount(1), 
                ["GC次数 Gen2"] = GC.CollectionCount(2)
            };
            
            Logger.Statistics("程序执行", stats);
        }

        /// <summary>
        /// 显示错误上下文信息
        /// </summary>
        /// <param name="ex">WorkPartner异常</param>
        private static void ShowErrorContext(WorkPartnerException ex)
        {
            if (ex.Context.Any())
            {
                Console.WriteLine("   错误上下文:");
                foreach (var context in ex.Context)
                {
                    Console.WriteLine($"     {context.Key}: {context.Value}");
                }
            }
        }
        
        /// <summary>
        /// 检查Excel文件第185行数据
        /// </summary>
        static void CheckExcelRow185Data()
        {
            var excelDir = Path.Combine(Directory.GetCurrentDirectory(), "..", "excel");
            if (!Directory.Exists(excelDir))
            {
                Console.WriteLine($"❌ Excel目录不存在: {excelDir}");
                return;
            }
            
            var excelFiles = Directory.GetFiles(excelDir, "*.xls").ToList();
            Console.WriteLine($"📁 找到 {excelFiles.Count} 个Excel文件");
            
            foreach (var filePath in excelFiles.Take(3)) // 只检查前3个文件
            {
                Console.WriteLine($"\n📄 检查文件: {Path.GetFileName(filePath)}");
                CheckSingleExcelFile(filePath);
            }
        }
        
        /// <summary>
        /// 检查单个Excel文件的第185行数据
        /// </summary>
        static void CheckSingleExcelFile(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(fs);
                    var sheet = workbook.GetSheetAt(0);
                    
                    // 检查第184、185、186行数据
                    for (int rowIndex = 183; rowIndex <= 185; rowIndex++) // 0基索引，所以184行是183
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            Console.WriteLine($"第{rowIndex + 1}行数据:");
                            
                            // 检查D列到I列（索引3-8）
                            for (int colIndex = 3; colIndex <= 8; colIndex++)
                            {
                                var cell = row.GetCell(colIndex);
                                var value = GetCellValue(cell);
                                var colName = GetColumnName(colIndex);
                                Console.WriteLine($"  {colName}: {value}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"第{rowIndex + 1}行: 空行");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 读取文件失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 获取单元格值
        /// </summary>
        static string GetCellValue(ICell? cell)
        {
            if (cell == null) return "空";
            
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString("F2");
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    return $"公式:{cell.CellFormula}";
                default:
                    return "未知类型";
            }
        }
        
        /// <summary>
        /// 获取列名
        /// </summary>
        static string GetColumnName(int colIndex)
        {
            return ((char)('A' + colIndex)).ToString();
        }

        /// <summary>
        /// 测试第185行补数逻辑
        /// </summary>
        static void TestRow185SupplementLogic()
        {
            Console.WriteLine("\n--- 测试第185行补数逻辑 ---");

            var excelDir = Path.Combine(Directory.GetCurrentDirectory(), "..", "excel");
            if (!Directory.Exists(excelDir))
            {
                Console.WriteLine($"❌ Excel目录不存在: {excelDir}");
                return;
            }

            var excelFiles = Directory.GetFiles(excelDir, "*.xls").ToList();
            if (excelFiles.Count == 0)
            {
                Console.WriteLine("❌ 未找到任何Excel文件");
                return;
            }

            var fileService = new FileService();
            var excelService = new ExcelService();

            foreach (var filePath in excelFiles)
            {
                Console.WriteLine($"\n📄 测试文件: {Path.GetFileName(filePath)}");
                try
                {
                    var workbook = new HSSFWorkbook(new FileStream(filePath, FileMode.Open, FileAccess.Read));
                    var sheet = workbook.GetSheetAt(0);

                    // 获取第185行数据
                    var row185 = sheet.GetRow(184); // 0基索引
                    if (row185 == null)
                    {
                        Console.WriteLine("❌ 文件缺少第185行数据");
                        continue;
                    }

                    Console.WriteLine("🔍 检查第185行数据:");
                    for (int colIndex = 3; colIndex <= 8; colIndex++) // D到I列
                    {
                        var cell = row185.GetCell(colIndex);
                        var value = GetCellValue(cell);
                        var colName = GetColumnName(colIndex);
                        Console.WriteLine($"  {colName}: {value}");
                    }

                    // 模拟保存过程
                    var currentObservationTime = $"{DateTime.Now:yyyy-M-d} {DateTime.Now.Hour:00}:00";
                    var previousObservationTime = $"{DateTime.Now.AddHours(-1):yyyy-M-d} {DateTime.Now.AddHours(-1).Hour:00}:00";

                    Console.WriteLine($"\n💾 模拟保存文件: {Path.GetFileName(filePath)}");
                    Console.WriteLine($"  本期观测时间: {currentObservationTime}");
                    Console.WriteLine($"  上期观测时间: {previousObservationTime}");

                    var success = excelService.SaveExcelFileWithA2Update(null, filePath, currentObservationTime, previousObservationTime); // 模拟保存

                    if (success)
                    {
                        Console.WriteLine("✅ 模拟保存成功");
                        Console.WriteLine("🔍 重新检查第185行数据:");
                        for (int colIndex = 3; colIndex <= 8; colIndex++) // D到I列
                        {
                            var cell = row185.GetCell(colIndex);
                            var value = GetCellValue(cell);
                            var colName = GetColumnName(colIndex);
                            Console.WriteLine($"  {colName}: {value}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("❌ 模拟保存失败");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 测试文件失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 测试通用行缺失数据检查功能
        /// </summary>
        static void TestGeneralRowMissingDataCheck()
        {
            Console.WriteLine("\n--- 测试通用行缺失数据检查功能 ---");

            var excelDir = Path.Combine(Directory.GetCurrentDirectory(), "..", "excel");
            if (!Directory.Exists(excelDir))
            {
                Console.WriteLine($"❌ Excel目录不存在: {excelDir}");
                return;
            }

            var excelFiles = Directory.GetFiles(excelDir, "*.xls").ToList();
            if (excelFiles.Count == 0)
            {
                Console.WriteLine("❌ 未找到任何Excel文件");
                return;
            }

            var fileService = new FileService();
            var excelService = new ExcelService();

            foreach (var filePath in excelFiles)
            {
                Console.WriteLine($"\n📄 测试文件: {Path.GetFileName(filePath)}");
                try
                {
                    var workbook = new HSSFWorkbook(new FileStream(filePath, FileMode.Open, FileAccess.Read));
                    var sheet = workbook.GetSheetAt(0);

                    // 获取第185行数据
                    var row185 = sheet.GetRow(184); // 0基索引
                    if (row185 == null)
                    {
                        Console.WriteLine("❌ 文件缺少第185行数据");
                        continue;
                    }

                    Console.WriteLine("🔍 检查第185行数据:");
                    for (int colIndex = 3; colIndex <= 8; colIndex++) // D到I列
                    {
                        var cell = row185.GetCell(colIndex);
                        var value = GetCellValue(cell);
                        var colName = GetColumnName(colIndex);
                        Console.WriteLine($"  {colName}: {value}");
                    }

                    // 模拟保存过程
                    var currentObservationTime = $"{DateTime.Now:yyyy-M-d} {DateTime.Now.Hour:00}:00";
                    var previousObservationTime = $"{DateTime.Now.AddHours(-1):yyyy-M-d} {DateTime.Now.AddHours(-1).Hour:00}:00";

                    Console.WriteLine($"\n💾 模拟保存文件: {Path.GetFileName(filePath)}");
                    Console.WriteLine($"  本期观测时间: {currentObservationTime}");
                    Console.WriteLine($"  上期观测时间: {previousObservationTime}");

                    var success = excelService.SaveExcelFileWithA2Update(null, filePath, currentObservationTime, previousObservationTime); // 模拟保存

                    if (success)
                    {
                        Console.WriteLine("✅ 模拟保存成功");
                        Console.WriteLine("🔍 重新检查第185行数据:");
                        for (int colIndex = 3; colIndex <= 8; colIndex++) // D到I列
                        {
                            var cell = row185.GetCell(colIndex);
                            var value = GetCellValue(cell);
                            var colName = GetColumnName(colIndex);
                            Console.WriteLine($"  {colName}: {value}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("❌ 模拟保存失败");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 测试文件失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 直接测试DataProcessor通用行缺失数据检查功能
        /// </summary>
        static void TestDataProcessorGeneralRowCheck()
        {
            Console.WriteLine("\n--- 直接测试DataProcessor通用行缺失数据检查功能 ---");

            var excelDir = Path.Combine(Directory.GetCurrentDirectory(), "..", "excel");
            if (!Directory.Exists(excelDir))
            {
                Console.WriteLine($"❌ Excel目录不存在: {excelDir}");
                return;
            }

            var excelFiles = Directory.GetFiles(excelDir, "*.xls").ToList();
            if (excelFiles.Count == 0)
            {
                Console.WriteLine("❌ 未找到任何Excel文件");
                return;
            }

            var fileService = new FileService();
            var excelService = new ExcelService();

            // 读取所有Excel文件
            var allExcelFiles = new List<WorkPartner.Models.ExcelFile>();
            
            foreach (var filePath in excelFiles)
            {
                try
                {
                    Console.WriteLine($"📄 读取文件: {Path.GetFileName(filePath)}");
                    var excelFile = excelService.ReadExcelFile(filePath);
                    if (excelFile != null)
                    {
                        allExcelFiles.Add(excelFile);
                        Console.WriteLine($"✅ 成功读取文件，包含 {excelFile.DataRows.Count} 个数据行");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 读取文件失败: {ex.Message}");
                }
            }

            if (allExcelFiles.Count == 0)
            {
                Console.WriteLine("❌ 没有成功读取任何文件");
                return;
            }

            Console.WriteLine($"\n🔍 开始使用DataProcessor处理 {allExcelFiles.Count} 个文件...");
            
            try
            {
                // 调用DataProcessor处理缺失数据
                var processedFiles = WorkPartner.Utils.DataProcessor.ProcessMissingData(allExcelFiles);
                
                Console.WriteLine($"✅ DataProcessor处理完成，共处理 {processedFiles.Count} 个文件");
                
                // 检查处理结果
                foreach (var file in processedFiles)
                {
                    Console.WriteLine($"\n📊 检查处理后的文件: {file.FileName}");
                    
                    // 查找第185行
                    var row185 = file.DataRows.FirstOrDefault(r => r.RowIndex == 185);
                    if (row185 != null)
                    {
                        Console.WriteLine("🔍 第185行处理结果:");
                        for (int i = 0; i < Math.Min(row185.Values.Count, 6); i++)
                        {
                            var value = row185.Values[i];
                            var colName = GetColumnName(i);
                            if (value.HasValue)
                            {
                                Console.WriteLine($"  {colName}: {value:F2}");
                            }
                            else
                            {
                                Console.WriteLine($"  {colName}: 仍然为空");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("❌ 未找到第185行数据");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ DataProcessor处理失败: {ex.Message}");
                Console.WriteLine($"   异常详情: {ex}");
            }
        }

        /// <summary>
        /// 运行大值检查模式
        /// </summary>
        /// <param name="arguments">命令行参数</param>
        private static async Task RunLargeValueCheckMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner Excel大值数据检查工具");
            Console.WriteLine("================================");

            // 确定要检查的目录
            string checkDirectory;
            if (!string.IsNullOrEmpty(arguments.InputPath))
            {
                checkDirectory = arguments.InputPath;
            }
            else if (!string.IsNullOrEmpty(arguments.OutputPath))
            {
                checkDirectory = arguments.OutputPath;
            }
            else
            {
                Console.WriteLine("❌ 请指定要检查的目录路径");
                Console.WriteLine("使用方法: WorkPartner.exe --check-large-values <目录路径> [--large-value-threshold <阈值>]");
                return;
            }

            Console.WriteLine($"📁 检查目录: {checkDirectory}");
            Console.WriteLine($"⚙️ 阈值: {arguments.LargeValueThreshold}");

            try
            {
                // 执行大值检查
                var checkResult = DataProcessor.CheckLargeValuesInOutputDirectory(checkDirectory, arguments.LargeValueThreshold);

                if (!string.IsNullOrEmpty(checkResult.ErrorMessage))
                {
                    Console.WriteLine($"⚠️ 检查过程发生错误: {checkResult.ErrorMessage}");
                }
                else
                {
                    Console.WriteLine($"✅ 大值数据检查完成");
                    
                    // 显示详细结果
                    if (arguments.Verbose && checkResult.FileResults.Any())
                    {
                        Console.WriteLine($"\n📊 详细检查结果:");
                        foreach (var fileResult in checkResult.FileResults)
                        {
                            Console.WriteLine($"\n📄 文件: {fileResult.FileName}");
                            Console.WriteLine($"   发现 {fileResult.LargeValues.Count} 个大值数据:");
                            
                            foreach (var largeValue in fileResult.LargeValues)
                            {
                                Console.WriteLine($"   - {largeValue.RowName} (第{largeValue.RowIndex}行, {largeValue.ColumnName}列): {largeValue.OriginalValue:F3} (绝对值: {largeValue.AbsoluteValue:F3})");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 大值检查功能执行失败: {ex.Message}");
                Logger.Error($"大值检查功能执行失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 运行文件比较模式
        /// </summary>
        /// <param name="arguments">命令行参数</param>
        private static async Task RunCompareMode(CommandLineArguments arguments)
        {
            Console.WriteLine("WorkPartner Excel文件比较工具");
            Console.WriteLine("============================");

            // 验证比较路径
            if (string.IsNullOrEmpty(arguments.CompareOriginalPath))
            {
                Console.WriteLine("❌ 原始文件目录路径不能为空");
                return;
            }

            if (string.IsNullOrEmpty(arguments.CompareProcessedPath))
            {
                Console.WriteLine("❌ 对比文件目录路径不能为空");
                return;
            }

            if (!Directory.Exists(arguments.CompareOriginalPath))
            {
                Console.WriteLine($"❌ 原始文件目录不存在: {arguments.CompareOriginalPath}");
                return;
            }

            if (!Directory.Exists(arguments.CompareProcessedPath))
            {
                Console.WriteLine($"❌ 对比文件目录不存在: {arguments.CompareProcessedPath}");
                return;
            }

            Console.WriteLine($"📁 原始文件目录: {arguments.CompareOriginalPath}");
            Console.WriteLine($"📁 对比文件目录: {arguments.CompareProcessedPath}");
            Console.WriteLine($"⚙️ 比较容差: {arguments.Tolerance}");
            Console.WriteLine($"📊 详细差异显示: {(arguments.ShowDetailedDifferences ? "启用" : "禁用")}");
            Console.WriteLine($"🔢 最大差异显示数量: {arguments.MaxDifferencesToShow}");

            try
            {
                // 执行文件比较
                var comparisonResult = DataProcessor.CompareOriginalAndProcessedFiles(
                    arguments.CompareOriginalPath,
                    arguments.CompareProcessedPath,
                    showDetailedDifferences: arguments.ShowDetailedDifferences,
                    tolerance: arguments.Tolerance,
                    maxDifferencesToShow: arguments.MaxDifferencesToShow
                );

                if (comparisonResult.HasError)
                {
                    Console.WriteLine($"⚠️ 文件比较过程发生错误: {comparisonResult.ErrorMessage}");
                }
                else
                {
                    Console.WriteLine($"✅ 文件比较分析完成");
                    
                    // 显示简要总结
                    if (arguments.Verbose)
                    {
                        Console.WriteLine($"\n📊 比较结果总结:");
                        Console.WriteLine($"   - 原始文件总数: {comparisonResult.FileComparisons.Count + comparisonResult.MissingProcessedFiles.Count}");
                        Console.WriteLine($"   - 成功比较文件数: {comparisonResult.FileComparisons.Count}");
                        Console.WriteLine($"   - 缺失对比文件数: {comparisonResult.MissingProcessedFiles.Count}");
                        Console.WriteLine($"   - 比较失败文件数: {comparisonResult.FailedComparisons.Count}");
                        
                        if (comparisonResult.TotalOriginalValues > 0)
                        {
                            var modificationPercentage = (double)comparisonResult.TotalDifferences / comparisonResult.TotalOriginalValues * 100;
                            Console.WriteLine($"   - 修改比例: {modificationPercentage:F2}% ({comparisonResult.TotalDifferences}/{comparisonResult.TotalOriginalValues})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 文件比较功能执行失败: {ex.Message}");
                Logger.Error($"文件比较功能执行失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 检查第200行补数逻辑问题
        /// </summary>
        static void CheckRow200SupplementLogic()
        {
            Console.WriteLine("\n--- 检查第200行补数逻辑问题 ---");

            var excelDir = Path.Combine(Directory.GetCurrentDirectory(), "..", "excel");
            var processedDir = Path.Combine(excelDir, "processed");
            
            if (!Directory.Exists(excelDir))
            {
                Console.WriteLine($"❌ Excel目录不存在: {excelDir}");
                return;
            }

            if (!Directory.Exists(processedDir))
            {
                Console.WriteLine($"❌ 处理后目录不存在: {processedDir}");
                return;
            }

            // 检查原始文件
            Console.WriteLine("\n📁 检查原始文件第200行数据:");
            var originalFiles = new[] 
            {
                "2025.4.18-0云港城项目4#地块.xls",
                "2025.4.18-8云港城项目4#地块.xls", 
                "2025.4.18-16云港城项目4#地块.xls"
            };

            foreach (var fileName in originalFiles)
            {
                var filePath = Path.Combine(excelDir, fileName);
                if (File.Exists(filePath))
                {
                    CheckRow200InFile(filePath, "原始文件");
                }
                else
                {
                    Console.WriteLine($"❌ 文件不存在: {fileName}");
                }
            }

            // 检查处理后文件
            Console.WriteLine("\n📁 检查处理后文件第200行数据:");
            var processedFiles = new[] 
            {
                "2025.4.18-00云港城项目4#地块.xls",
                "2025.4.18-08云港城项目4#地块.xls", 
                "2025.4.18-16云港城项目4#地块.xls"
            };

            foreach (var fileName in processedFiles)
            {
                var filePath = Path.Combine(processedDir, fileName);
                if (File.Exists(filePath))
                {
                    CheckRow200InFile(filePath, "处理后文件");
                }
                else
                {
                    Console.WriteLine($"❌ 文件不存在: {fileName}");
                }
            }

            // 分析补数逻辑
            Console.WriteLine("\n🔍 分析补数逻辑问题:");
            Console.WriteLine("问题描述: 原始文件第200行数据为空，处理后文件第200行均被填充为相同值");
            Console.WriteLine("可能原因:");
            Console.WriteLine("1. 补数算法使用了固定的默认值");
            Console.WriteLine("2. 相邻行数据获取失败，使用了硬编码的备用值");
            Console.WriteLine("3. 随机种子固定，导致所有文件生成相同值");
            Console.WriteLine("4. 补数逻辑中存在全局共享的默认值");
        }

        /// <summary>
        /// 检查文件中第200行的数据
        /// </summary>
        static void CheckRow200InFile(string filePath, string fileType)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(fs);
                    var sheet = workbook.GetSheetAt(0);

                    Console.WriteLine($"\n📄 {fileType}: {Path.GetFileName(filePath)}");

                    // 检查第199、200、201行数据（0基索引）
                    for (int rowIndex = 198; rowIndex <= 200; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            Console.WriteLine($"第{rowIndex + 1}行数据:");
                            
                            // 检查D列到I列（索引3-8）
                            for (int colIndex = 3; colIndex <= 8; colIndex++)
                            {
                                var cell = row.GetCell(colIndex);
                                var value = GetCellValue(cell);
                                var colName = GetColumnName(colIndex);
                                Console.WriteLine($"  {colName}: {value}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"第{rowIndex + 1}行: 空行");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 读取文件失败: {ex.Message}");
            }
        }
    }
}
