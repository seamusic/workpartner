using WorkPartner.Models;
using WorkPartner.Services;
using WorkPartner.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WorkPartner.Utils
{
    /// <summary>
    /// 文件处理器 - 负责文件扫描、解析、读取等操作
    /// </summary>
    public static class FileProcessor
    {
        /// <summary>
        /// 验证输入路径
        /// </summary>
        /// <param name="path">要验证的路径</param>
        /// <returns>路径是否有效</returns>
        public static bool ValidateInputPath(string path)
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

        /// <summary>
        /// 创建输出目录
        /// </summary>
        /// <param name="outputPath">输出目录路径</param>
        public static void CreateOutputDirectory(string outputPath)
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

        /// <summary>
        /// 扫描Excel文件
        /// </summary>
        /// <param name="inputPath">输入目录路径</param>
        /// <returns>Excel文件路径列表</returns>
        public static List<string> ScanExcelFiles(string inputPath)
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

        /// <summary>
        /// 解析文件名并排序
        /// </summary>
        /// <param name="filePaths">文件路径列表</param>
        /// <returns>解析后的Excel文件列表</returns>
        public static List<ExcelFile> ParseAndSortFiles(List<string> filePaths)
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

        /// <summary>
        /// 读取Excel数据
        /// </summary>
        /// <param name="files">Excel文件列表</param>
        /// <returns>包含数据的Excel文件列表</returns>
        public static List<ExcelFile> ReadExcelData(List<ExcelFile> files)
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

        /// <summary>
        /// 保存处理后的Excel文件
        /// </summary>
        /// <param name="processedFiles">处理后的文件列表</param>
        /// <param name="outputPath">输出目录</param>
        public static async Task SaveProcessedFiles(List<ExcelFile> processedFiles, string outputPath)
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
    }
}
