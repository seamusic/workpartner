using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// 数据质量检查服务
    /// </summary>
    public class DataQualityService
    {
        private readonly ILogger<DataQualityService> _logger;

        public DataQualityService(ILogger<DataQualityService> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 检查导出文件质量
        /// </summary>
        public async Task<DataQualityReport> CheckExportQualityAsync(List<ExportResult> exportResults, List<MergeResult> mergeResults)
        {
            try
            {
                _logger.LogInformation("开始检查导出数据质量...");

                var report = new DataQualityReport
                {
                    CheckTime = DateTime.Now,
                    TotalExportFiles = exportResults.Count,
                    TotalMergeFiles = mergeResults.Count
                };

                // 检查导出文件质量
                await CheckExportFileQualityAsync(exportResults, report);

                // 检查合并文件质量
                await CheckMergeFileQualityAsync(mergeResults, report);

                // 计算总体质量评分
                CalculateOverallQualityScore(report);

                _logger.LogInformation($"数据质量检查完成，总体评分: {report.OverallQualityScore:F1}/100");

                return report;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"数据质量检查失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 检查导出文件质量
        /// </summary>
        private async Task CheckExportFileQualityAsync(List<ExportResult> exportResults, DataQualityReport report)
        {
            foreach (var result in exportResults)
            {
                if (!result.Success || string.IsNullOrEmpty(result.FilePath))
                {
                    continue;
                }

                var fileQuality = new FileQualityInfo
                {
                    FilePath = result.FilePath,
                    FileName = result.FileName,
                    ProjectName = result.ProjectName,
                    DataName = result.DataName,
                    FileSize = GetFileSize(result.FilePath),
                    LastModified = GetFileLastModified(result.FilePath)
                };

                // 检查文件完整性
                await CheckFileIntegrityAsync(fileQuality);

                // 检查文件格式
                CheckFileFormat(fileQuality);

                // 检查文件大小合理性
                CheckFileSizeReasonableness(fileQuality);

                report.FileQualityIssues.Add(fileQuality);
            }
        }

        /// <summary>
        /// 检查合并文件质量
        /// </summary>
        private async Task CheckMergeFileQualityAsync(List<MergeResult> mergeResults, DataQualityReport report)
        {
            foreach (var result in mergeResults)
            {
                if (!result.Success || string.IsNullOrEmpty(result.MergedFilePath))
                {
                    continue;
                }

                var mergeQuality = new MergeQualityInfo
                {
                    MergedFilePath = result.MergedFilePath,
                    ProjectName = result.ProjectName,
                    DataName = result.DataName,
                    SourceFileCount = result.SourceFiles?.Count ?? 0,
                    FileSize = GetFileSize(result.MergedFilePath),
                    LastModified = GetFileLastModified(result.MergedFilePath)
                };

                // 检查合并文件完整性
                await CheckFileIntegrityAsync(mergeQuality);

                // 检查合并文件格式
                CheckFileFormat(mergeQuality);

                // 检查合并逻辑合理性
                CheckMergeLogicReasonableness(mergeQuality, result);

                report.MergeQualityIssues.Add(mergeQuality);
            }
        }

        /// <summary>
        /// 检查文件完整性
        /// </summary>
        private async Task CheckFileIntegrityAsync(FileQualityInfo fileQuality)
        {
            try
            {
                if (!File.Exists(fileQuality.FilePath))
                {
                    fileQuality.Issues.Add("文件不存在");
                    fileQuality.IntegrityScore = 0;
                    return;
                }

                // 检查文件是否可以正常打开
                using var stream = File.OpenRead(fileQuality.FilePath);
                var buffer = new byte[1024];
                var bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);

                if (bytesRead == 0)
                {
                    fileQuality.Issues.Add("文件为空");
                    fileQuality.IntegrityScore = 0;
                }
                else if (fileQuality.FileSize < 1024) // 小于1KB
                {
                    fileQuality.Issues.Add("文件过小，可能数据不完整");
                    fileQuality.IntegrityScore = 70;
                }
                else
                {
                    fileQuality.IntegrityScore = 100;
                }
            }
            catch (Exception ex)
            {
                fileQuality.Issues.Add($"文件完整性检查失败: {ex.Message}");
                fileQuality.IntegrityScore = 0;
            }
        }

        /// <summary>
        /// 检查文件格式
        /// </summary>
        private void CheckFileFormat(FileQualityInfo fileQuality)
        {
            try
            {
                var extension = Path.GetExtension(fileQuality.FilePath).ToLower();
                
                switch (extension)
                {
                    case ".xlsx":
                        fileQuality.FormatScore = 100;
                        fileQuality.FormatType = "Excel 2007+";
                        break;
                    case ".xls":
                        fileQuality.FormatScore = 80;
                        fileQuality.FormatType = "Excel 97-2003";
                        break;
                    case ".csv":
                        fileQuality.FormatScore = 90;
                        fileQuality.FormatType = "CSV";
                        break;
                    default:
                        fileQuality.FormatScore = 0;
                        fileQuality.FormatType = "未知格式";
                        fileQuality.Issues.Add($"不支持的文件格式: {extension}");
                        break;
                }
            }
            catch (Exception ex)
            {
                fileQuality.FormatScore = 0;
                fileQuality.FormatType = "检查失败";
                fileQuality.Issues.Add($"文件格式检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查文件大小合理性
        /// </summary>
        private void CheckFileSizeReasonableness(FileQualityInfo fileQuality)
        {
            try
            {
                if (fileQuality.FileSize < 100) // 小于100字节
                {
                    fileQuality.Issues.Add("文件过小，可能没有有效数据");
                    fileQuality.SizeScore = 0;
                }
                else if (fileQuality.FileSize < 1024) // 小于1KB
                {
                    fileQuality.Issues.Add("文件较小，数据可能不完整");
                    fileQuality.SizeScore = 60;
                }
                else if (fileQuality.FileSize > 100 * 1024 * 1024) // 大于100MB
                {
                    fileQuality.Issues.Add("文件过大，可能影响处理性能");
                    fileQuality.SizeScore = 80;
                }
                else
                {
                    fileQuality.SizeScore = 100;
                }
            }
            catch (Exception ex)
            {
                fileQuality.SizeScore = 0;
                fileQuality.Issues.Add($"文件大小检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查合并逻辑合理性
        /// </summary>
        private void CheckMergeLogicReasonableness(MergeQualityInfo mergeQuality, MergeResult mergeResult)
        {
            try
            {
                // 检查源文件数量
                if (mergeQuality.SourceFileCount == 0)
                {
                    mergeQuality.Issues.Add("没有源文件");
                    mergeQuality.LogicScore = 0;
                }
                else if (mergeQuality.SourceFileCount == 1)
                {
                    mergeQuality.Issues.Add("只有一个源文件，无需合并");
                    mergeQuality.LogicScore = 50;
                }
                else if (mergeQuality.SourceFileCount > 100)
                {
                    mergeQuality.Issues.Add("源文件过多，可能影响合并性能");
                    mergeQuality.LogicScore = 70;
                }
                else
                {
                    mergeQuality.LogicScore = 100;
                }

                // 检查合并后的文件大小是否合理
                if (mergeResult.SourceFiles != null && mergeResult.SourceFiles.Any())
                {
                    var totalSourceSize = mergeResult.SourceFiles.Sum(f => GetFileSize(f));
                    if (totalSourceSize > 0)
                    {
                        var compressionRatio = (double)mergeQuality.FileSize / totalSourceSize;
                        if (compressionRatio > 1.5)
                        {
                            mergeQuality.Issues.Add($"合并后文件过大，压缩比: {compressionRatio:F2}");
                            mergeQuality.LogicScore = Math.Max(mergeQuality.LogicScore - 20, 0);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                mergeQuality.LogicScore = 0;
                mergeQuality.Issues.Add($"合并逻辑检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 计算总体质量评分
        /// </summary>
        private void CalculateOverallQualityScore(DataQualityReport report)
        {
            try
            {
                var totalScore = 0.0;
                var totalWeight = 0.0;

                // 文件质量评分（权重40%）
                if (report.FileQualityIssues.Any())
                {
                    var fileScore = report.FileQualityIssues.Average(f => 
                        (f.IntegrityScore + f.FormatScore + f.SizeScore) / 3.0);
                    totalScore += fileScore * 0.4;
                    totalWeight += 0.4;
                }

                // 合并质量评分（权重30%）
                if (report.MergeQualityIssues.Any())
                {
                    var mergeScore = report.MergeQualityIssues.Average(m => 
                        (m.IntegrityScore + m.FormatScore + m.LogicScore) / 3.0);
                    totalScore += mergeScore * 0.3;
                    totalWeight += 0.3;
                }

                // 数据完整性评分（权重30%）
                var dataIntegrityScore = CalculateDataIntegrityScore(report);
                totalScore += dataIntegrityScore * 0.3;
                totalWeight += 0.3;

                report.OverallQualityScore = totalWeight > 0 ? totalScore / totalWeight : 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"计算总体质量评分失败: {ex.Message}");
                report.OverallQualityScore = 0;
            }
        }

        /// <summary>
        /// 计算数据完整性评分
        /// </summary>
        private double CalculateDataIntegrityScore(DataQualityReport report)
        {
            try
            {
                var totalFiles = report.TotalExportFiles + report.TotalMergeFiles;
                if (totalFiles == 0)
                {
                    return 0;
                }

                var validFiles = report.FileQualityIssues.Count(f => f.IntegrityScore >= 80) +
                                report.MergeQualityIssues.Count(m => m.IntegrityScore >= 80);

                return (double)validFiles / totalFiles * 100;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"计算数据完整性评分失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 获取文件大小
        /// </summary>
        private long GetFileSize(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    return fileInfo.Length;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 获取文件最后修改时间
        /// </summary>
        private DateTime GetFileLastModified(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    return fileInfo.LastWriteTime;
                }
                return DateTime.MinValue;
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// 生成质量检查报告
        /// </summary>
        public async Task<string> GenerateQualityReportAsync(DataQualityReport report, string outputPath)
        {
            try
            {
                var reportData = new
                {
                    report.CheckTime,
                    report.TotalExportFiles,
                    report.TotalMergeFiles,
                    report.OverallQualityScore,
                    FileQualitySummary = new
                    {
                        TotalFiles = report.FileQualityIssues.Count,
                        HighQualityFiles = report.FileQualityIssues.Count(f => f.IntegrityScore >= 80),
                        MediumQualityFiles = report.FileQualityIssues.Count(f => f.IntegrityScore >= 60 && f.IntegrityScore < 80),
                        LowQualityFiles = report.FileQualityIssues.Count(f => f.IntegrityScore < 60)
                    },
                    MergeQualitySummary = new
                    {
                        TotalFiles = report.MergeQualityIssues.Count,
                        HighQualityFiles = report.MergeQualityIssues.Count(m => m.LogicScore >= 80),
                        MediumQualityFiles = report.MergeQualityIssues.Count(m => m.LogicScore >= 60 && m.LogicScore < 80),
                        LowQualityFiles = report.MergeQualityIssues.Count(m => m.LogicScore < 60)
                    },
                    Issues = report.FileQualityIssues.Where(f => f.Issues.Any()).Select(f => new
                    {
                        f.FilePath,
                        f.ProjectName,
                        f.DataName,
                        f.Issues
                    }).ToList()
                };

                var json = JsonSerializer.Serialize(reportData, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                var fileName = $"data-quality-report-{DateTime.Now:yyyyMMdd-HHmmss}.json";
                var fullPath = Path.Combine(outputPath, fileName);

                await File.WriteAllTextAsync(fullPath, json);
                _logger.LogInformation($"生成数据质量报告: {fullPath}");

                return fullPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"生成数据质量报告失败: {ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// 数据质量报告
    /// </summary>
    public class DataQualityReport
    {
        /// <summary>
        /// 检查时间
        /// </summary>
        public DateTime CheckTime { get; set; }

        /// <summary>
        /// 总导出文件数
        /// </summary>
        public int TotalExportFiles { get; set; }

        /// <summary>
        /// 总合并文件数
        /// </summary>
        public int TotalMergeFiles { get; set; }

        /// <summary>
        /// 总体质量评分
        /// </summary>
        public double OverallQualityScore { get; set; }

        /// <summary>
        /// 文件质量问题列表
        /// </summary>
        public List<FileQualityInfo> FileQualityIssues { get; set; } = new();

        /// <summary>
        /// 合并质量问题列表
        /// </summary>
        public List<MergeQualityInfo> MergeQualityIssues { get; set; } = new();
    }

    /// <summary>
    /// 文件质量信息
    /// </summary>
    public class FileQualityInfo
    {
        /// <summary>
        /// 文件路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 数据类型名称
        /// </summary>
        public string DataName { get; set; } = string.Empty;

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// 最后修改时间
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// 完整性评分
        /// </summary>
        public double IntegrityScore { get; set; }

        /// <summary>
        /// 格式评分
        /// </summary>
        public double FormatScore { get; set; }

        /// <summary>
        /// 大小合理性评分
        /// </summary>
        public double SizeScore { get; set; }

        /// <summary>
        /// 文件格式类型
        /// </summary>
        public string FormatType { get; set; } = string.Empty;

        /// <summary>
        /// 质量问题列表
        /// </summary>
        public List<string> Issues { get; set; } = new();

        /// <summary>
        /// 综合质量评分
        /// </summary>
        public double OverallScore => (IntegrityScore + FormatScore + SizeScore) / 3.0;
    }

    /// <summary>
    /// 合并质量信息
    /// </summary>
    public class MergeQualityInfo : FileQualityInfo
    {
        /// <summary>
        /// 源文件数量
        /// </summary>
        public int SourceFileCount { get; set; }

        /// <summary>
        /// 合并逻辑评分
        /// </summary>
        public double LogicScore { get; set; }

        /// <summary>
        /// 综合质量评分（重写）
        /// </summary>
        public new double OverallScore => (IntegrityScore + FormatScore + LogicScore) / 3.0;
    }
}
