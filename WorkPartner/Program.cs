using WorkPartner.Models;
using WorkPartner.Utils;

namespace WorkPartner
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("WorkPartner Excel数据处理工具 - 阶段1测试");
            Console.WriteLine("==========================================");

            // 初始化日志
            Logger.Initialize("logs/test.log", LogLevel.Debug);

            try
            {
                // 测试文件名解析
                TestFileNameParser();

                // 测试数据模型
                TestDataModels();

                // 测试数据处理工具
                TestDataProcessor();

                // 测试日志功能
                TestLogger();

                Console.WriteLine("\n✅ 阶段1基础功能测试完成！");
            }
            catch (Exception ex)
            {
                Logger.Error("测试过程中发生错误", ex);
                Console.WriteLine($"\n❌ 测试失败: {ex.Message}");
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void TestFileNameParser()
        {
            Console.WriteLine("\n--- 测试文件名解析 ---");

            var testFiles = new[]
            {
                "2025.4.18-8云港城项目4#地块.xls",
                "2025.4.19-16云港城项目4#地块.xls",
                "invalid_file.txt",
                "2025.4.20-25云港城项目4#地块.xls" // 无效时间
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
                FileName = "test.xls",
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
                FileName = "2025.4.18-0云港城项目4#地块.xls",
                Date = new DateTime(2025, 4, 18),
                Hour = 0,
                ProjectName = "云港城项目4#地块.xls"
            };

            var file2 = new ExcelFile
            {
                FileName = "2025.4.18-16云港城项目4#地块.xls",
                Date = new DateTime(2025, 4, 18),
                Hour = 16,
                ProjectName = "云港城项目4#地块.xls"
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
    }
}
