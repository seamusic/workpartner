using WorkPartner.Models;
using WorkPartner.Services;
using WorkPartner.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 测试工具类 - 包含各种测试方法
    /// </summary>
    public static class TestTools
    {
        /// <summary>
        /// 测试文件名解析器
        /// </summary>
        public static void TestFileNameParser()
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

        /// <summary>
        /// 测试数据模型
        /// </summary>
        public static void TestDataModels()
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

        /// <summary>
        /// 测试数据处理工具
        /// </summary>
        public static void TestDataProcessor()
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

        /// <summary>
        /// 测试日志功能
        /// </summary>
        public static void TestLogger()
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
        /// 检查Excel文件第185行数据
        /// </summary>
        public static void CheckExcelRow185Data()
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
        private static void CheckSingleExcelFile(string filePath)
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
        private static string GetCellValue(ICell? cell)
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
        private static string GetColumnName(int colIndex)
        {
            return ((char)('A' + colIndex)).ToString();
        }

        /// <summary>
        /// 测试第185行补数逻辑
        /// </summary>
        public static void TestRow185SupplementLogic()
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
        public static void TestGeneralRowMissingDataCheck()
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
        public static void TestDataProcessorGeneralRowCheck()
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
        /// 检查第200行补数逻辑问题
        /// </summary>
        public static void CheckRow200SupplementLogic()
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
        private static void CheckRow200InFile(string filePath, string fileType)
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
