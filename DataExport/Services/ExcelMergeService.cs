using DataExport.Models;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace DataExport.Services
{
    /// <summary>
    /// Excel文件合并服务
    /// </summary>
    public class ExcelMergeService
    {
        private readonly ILogger<ExcelMergeService> _logger;
        private readonly ExportConfig _config;

        public ExcelMergeService(ILogger<ExcelMergeService> logger, ExportConfig config)
        {
            _logger = logger;
            _config = config;
            
            // 设置EPPlus许可证上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// 合并导出结果中的Excel文件
        /// </summary>
        /// <param name="exportResults">导出结果列表</param>
        /// <returns>合并结果</returns>
        public async Task<List<MergeResult>> MergeExcelFilesAsync(List<ExportResult> exportResults)
        {
            _logger.LogInformation("开始Excel文件合并");
            
            var mergeResults = new List<MergeResult>();
            
            try
            {
                // 按项目名称和数据名称分组
                var groupedResults = exportResults
                    .Where(r => r.Success && !string.IsNullOrEmpty(r.FilePath) && File.Exists(r.FilePath))
                    .GroupBy(r => new { r.ProjectName, r.DataName })
                    .ToList();

                _logger.LogInformation($"找到 {groupedResults.Count} 个合并组");

                var totalGroups = groupedResults.Count;
                var currentGroup = 0;

                foreach (var group in groupedResults)
                {
                    currentGroup++;
                    var progressPercent = (double)currentGroup / totalGroups * 100;
                    
                    _logger.LogInformation($"合并进度: {currentGroup}/{totalGroups} ({progressPercent:F1}%) - 处理: {group.Key.ProjectName} - {group.Key.DataName}");

                    try
                    {
                        var mergeResult = await MergeSingleGroupAsync(group.Key.ProjectName, group.Key.DataName, group.ToList());
                        mergeResults.Add(mergeResult);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "合并组失败: {ProjectName} - {DataName}", group.Key.ProjectName, group.Key.DataName);
                        
                        mergeResults.Add(new MergeResult
                        {
                            Success = false,
                            ProjectName = group.Key.ProjectName,
                            DataName = group.Key.DataName,
                            ErrorMessage = ex.Message,
                            SourceFiles = group.Select(g => g.FilePath).Where(f => !string.IsNullOrEmpty(f)).ToList()
                        });
                    }
                }

                var successCount = mergeResults.Count(r => r.Success);
                _logger.LogInformation($"Excel合并完成: {successCount}/{mergeResults.Count} 个组合并成功");

                return mergeResults;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel合并过程中发生错误");
                throw;
            }
        }

        /// <summary>
        /// 合并单个分组的Excel文件
        /// </summary>
        /// <param name="projectName">项目名称</param>
        /// <param name="dataName">数据名称</param>
        /// <param name="exportResults">该分组的导出结果</param>
        /// <returns>合并结果</returns>
        private async Task<MergeResult> MergeSingleGroupAsync(string projectName, string dataName, List<ExportResult> exportResults)
        {
                                var sourceFiles = exportResults.Select(r => r.FilePath).Where(f => !string.IsNullOrEmpty(f)).ToList();
            _logger.LogInformation($"开始合并: {projectName} - {dataName}, 源文件数量: {sourceFiles.Count}");

            // 生成合并后的文件名
            var mergedFileName = GenerateMergedFileName(projectName, dataName);
            var mergedFilePath = Path.Combine(_config.ExportSettings.OutputDirectory, mergedFileName);

            try
            {
                using var mergedPackage = new ExcelPackage();
                var mergedWorksheet = mergedPackage.Workbook.Worksheets.Add("合并数据");
                var currentRow = 1;
                var isFirstFile = true;

                foreach (var sourceFile in sourceFiles)
                {
                    _logger.LogDebug($"处理源文件: {sourceFile}");

                    if (!File.Exists(sourceFile))
                    {
                        _logger.LogWarning($"源文件不存在: {sourceFile}");
                        continue;
                    }

                    try
                    {
                        using var sourcePackage = new ExcelPackage(new FileInfo(sourceFile));
                        var sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();
                        
                        if (sourceWorksheet == null)
                        {
                            _logger.LogWarning($"源文件没有工作表: {sourceFile}");
                            continue;
                        }

                        var rowCount = sourceWorksheet.Dimension?.Rows ?? 0;
                        var colCount = sourceWorksheet.Dimension?.Columns ?? 0;

                        if (rowCount == 0 || colCount == 0)
                        {
                            _logger.LogWarning($"源文件数据为空: {sourceFile}");
                            continue;
                        }

                        // 第一个文件复制所有行包括标题
                        // 后续文件跳过标题行
                        var startRow = isFirstFile ? 1 : 2;
                        var copyRowCount = isFirstFile ? rowCount : rowCount - 1;

                        if (startRow <= rowCount)
                        {
                            // 复制数据
                            for (int row = startRow; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cellValue = sourceWorksheet.Cells[row, col].Value;
                                    mergedWorksheet.Cells[currentRow, col].Value = cellValue;
                                }
                                currentRow++;
                            }
                        }

                        isFirstFile = false;
                        _logger.LogDebug($"成功处理文件: {sourceFile}, 复制了 {copyRowCount} 行数据");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"处理源文件失败: {sourceFile}");
                        // 继续处理其他文件
                    }
                }

                // 自动调整列宽
                if (mergedWorksheet.Dimension != null)
                {
                    mergedWorksheet.Cells[mergedWorksheet.Dimension.Address].AutoFitColumns();
                }

                // 保存合并后的文件
                var mergedFileInfo = new FileInfo(mergedFilePath);
                await mergedPackage.SaveAsAsync(mergedFileInfo);

                _logger.LogInformation($"合并完成: {mergedFilePath}");

                // 验证合并结果
                var validation = await ValidateMergedFileAsync(mergedFilePath, sourceFiles.Count);

                return new MergeResult
                {
                    Success = true,
                    ProjectName = projectName,
                    DataName = dataName,
                    MergedFilePath = mergedFilePath,
                    MergedFileName = mergedFileName,
                    SourceFiles = sourceFiles,
                    SourceFileCount = sourceFiles.Count,
                    DataRowCount = currentRow - 1,
                    ValidationResult = validation
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"合并文件失败: {projectName} - {dataName}");
                
                return new MergeResult
                {
                    Success = false,
                    ProjectName = projectName,
                    DataName = dataName,
                    ErrorMessage = ex.Message,
                    SourceFiles = sourceFiles
                };
            }
        }

        /// <summary>
        /// 生成合并后的文件名
        /// </summary>
        /// <param name="projectName">项目名称</param>
        /// <param name="dataName">数据名称</param>
        /// <returns>文件名</returns>
        private static string GenerateMergedFileName(string projectName, string dataName)
        {
            // 清理文件名中的非法字符
            var cleanProjectName = Regex.Replace(projectName, @"[<>:""/\\|?*]", "_");
            var cleanDataName = Regex.Replace(dataName, @"[<>:""/\\|?*]", "_");
            
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return $"{cleanProjectName}_{cleanDataName}_合并_{timestamp}.xlsx";
        }

        /// <summary>
        /// 验证合并后的文件
        /// </summary>
        /// <param name="mergedFilePath">合并文件路径</param>
        /// <param name="expectedSourceCount">预期源文件数量</param>
        /// <returns>验证结果</returns>
        private Task<MergeValidationResult> ValidateMergedFileAsync(string mergedFilePath, int expectedSourceCount)
        {
            try
            {
                if (!File.Exists(mergedFilePath))
                {
                    return Task.FromResult(new MergeValidationResult
                    {
                        IsValid = false,
                        Message = "合并文件不存在"
                    });
                }

                var fileInfo = new FileInfo(mergedFilePath);
                if (fileInfo.Length == 0)
                {
                    return Task.FromResult(new MergeValidationResult
                    {
                        IsValid = false,
                        Message = "合并文件为空"
                    });
                }

                using var package = new ExcelPackage(fileInfo);
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                {
                    return Task.FromResult(new MergeValidationResult
                    {
                        IsValid = false,
                        Message = "合并文件没有工作表"
                    });
                }

                var rowCount = worksheet.Dimension?.Rows ?? 0;
                var colCount = worksheet.Dimension?.Columns ?? 0;

                return Task.FromResult(new MergeValidationResult
                {
                    IsValid = rowCount > 1 && colCount > 0, // 至少有标题行和一行数据
                    Message = $"验证通过: {rowCount} 行, {colCount} 列",
                    RowCount = rowCount,
                    ColumnCount = colCount,
                    FileSizeBytes = fileInfo.Length
                });
            }
            catch (Exception ex)
            {
                return Task.FromResult(new MergeValidationResult
                {
                    IsValid = false,
                    Message = $"验证失败: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// 清理临时文件（可选）
        /// </summary>
        /// <param name="exportResults">导出结果</param>
        /// <param name="keepOriginalFiles">是否保留原文件</param>
        public Task CleanupFilesAsync(List<ExportResult> exportResults, bool keepOriginalFiles = true)
        {
            if (keepOriginalFiles)
            {
                _logger.LogInformation("保留原始导出文件");
                return Task.CompletedTask;
            }

            _logger.LogInformation("开始清理临时文件");
            
            foreach (var result in exportResults.Where(r => r.Success && !string.IsNullOrEmpty(r.FilePath)))
            {
                try
                {
                    if (File.Exists(result.FilePath))
                    {
                        File.Delete(result.FilePath);
                        _logger.LogDebug($"删除文件: {result.FilePath}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"删除文件失败: {result.FilePath}");
                }
            }
            
            return Task.CompletedTask;
        }
    }
}
