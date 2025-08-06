using WorkPartner.Models;
using WorkPartner.Utils;
using WorkPartner.Services;
using System.IO;

namespace WorkPartner
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("WorkPartner Excel数据处理工具 - 阶段3实现");
            Console.WriteLine("==========================================");

            // 初始化日志
            Logger.Initialize("logs/workpartner.log", LogLevel.Info);

            try
            {
                // 解析命令行参数
                var arguments = ParseCommandLineArguments(args);
                if (arguments == null)
                {
                    ShowUsage();
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
                
                // 3.1 数据补充算法
                Console.WriteLine("📊 处理缺失数据...");
                var processedFiles = DataProcessor.ProcessMissingData(filesWithData);
                
                // 保存处理后的数据到Excel文件
                Console.WriteLine("💾 保存处理后的数据...");
                await SaveProcessedFiles(processedFiles, arguments.OutputPath);
                
                // 3.2 数据完整性检查
                Console.WriteLine("🔍 检查数据完整性...");
                var completenessResult = DataProcessor.CheckCompleteness(processedFiles);
                
                // 生成补充文件列表
                var supplementFiles = DataProcessor.GenerateSupplementFiles(processedFiles);
                
                // 创建补充文件
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
                
                // 数据质量验证
                var qualityReport = DataProcessor.ValidateDataQuality(processedFiles);

                // 显示处理结果
                DisplayProcessingResults(processedFiles, completenessResult, supplementFiles, qualityReport);

                Console.WriteLine("\n✅ 阶段3数据处理逻辑完成！");
            }
            catch (Exception ex)
            {
                Logger.Error("程序执行过程中发生错误", ex);
                Console.WriteLine($"\n❌ 程序执行失败: {ex.Message}");
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
                    case "-h":
                    case "--help":
                        return null;
                    default:
                        // 如果没有指定参数，第一个参数作为输入路径
                        if (string.IsNullOrEmpty(arguments.InputPath))
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
            Console.WriteLine("");
            Console.WriteLine("参数:");
            Console.WriteLine("  <输入目录>              包含Excel文件的目录路径");
            Console.WriteLine("");
            Console.WriteLine("选项:");
            Console.WriteLine("  -o, --output <目录>     输出目录路径 (默认: <输入目录>/processed)");
            Console.WriteLine("  -v, --verbose           详细输出模式");
            Console.WriteLine("  -h, --help              显示此帮助信息");
            Console.WriteLine("");
            Console.WriteLine("支持的文件格式:");
            Console.WriteLine("  ✅ .xlsx (Excel 2007+)");
            Console.WriteLine("  ✅ .xls (Excel 97-2003)");
            Console.WriteLine("");
            Console.WriteLine("示例:");
            Console.WriteLine("  WorkPartner.exe C:\\excel\\");
            Console.WriteLine("  WorkPartner.exe ..\\excel\\");
            Console.WriteLine("  WorkPartner.exe C:\\excel\\ -o C:\\output\\ -v");
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
                Logger.Info($"开始扫描目录: {inputPath}");
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
                throw;
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
            var filesWithData = new List<ExcelFile>();
            var excelService = new ExcelService();

            for (int i = 0; i < files.Count; i++)
            {
                var file = files[i];
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
                    // 继续处理其他文件
                }
            }

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
            var excelService = new ExcelService();
            int savedCount = 0;
            int totalFiles = processedFiles.Count;

            foreach (var file in processedFiles)
            {
                try
                {
                    var outputFilePath = Path.Combine(outputPath, file.FileName);
                    var success = await excelService.SaveExcelFileAsync(file, outputFilePath);
                    
                    if (success)
                    {
                        savedCount++;
                        Console.WriteLine($"✅ 已保存: {file.FileName}");
                    }
                    else
                    {
                        Console.WriteLine($"❌ 保存失败: {file.FileName}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 保存文件失败: {file.FileName}");
                    Console.WriteLine($"   错误: {ex.Message}");
                    Logger.Error($"保存文件失败: {file.FileName}", ex);
                }
            }

            Console.WriteLine($"✅ 成功保存 {savedCount}/{totalFiles} 个处理后的文件");
        }
    }
}
