using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// 导出文件管理服务
    /// </summary>
    public class ExportFileManager
    {
        private readonly ILogger<ExportFileManager> _logger;
        private readonly string _baseDirectory;
        private readonly ExportFileManagerConfig _config;

        public ExportFileManager(ILogger<ExportFileManager> logger, string baseDirectory = "./exports", ExportFileManagerConfig? config = null)
        {
            _logger = logger;
            _baseDirectory = baseDirectory;
            _config = config ?? new ExportFileManagerConfig();
            
            // 确保基础目录存在
            if (!Directory.Exists(_baseDirectory))
            {
                Directory.CreateDirectory(_baseDirectory);
            }
        }

        /// <summary>
        /// 获取文件存储信息
        /// </summary>
        public async Task<FileStorageInfo> GetStorageInfoAsync()
        {
            try
            {
                var info = new FileStorageInfo
                {
                    BaseDirectory = _baseDirectory,
                    CheckTime = DateTime.Now
                };

                await CalculateStorageInfoAsync(info);
                _logger.LogInformation($"获取存储信息完成，总文件数: {info.TotalFiles}, 总大小: {FormatFileSize(info.TotalSize)}");

                return info;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"获取存储信息失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 整理导出文件
        /// </summary>
        public async Task<FileOrganizationResult> OrganizeFilesAsync()
        {
            try
            {
                _logger.LogInformation("开始整理导出文件...");

                var result = new FileOrganizationResult
                {
                    StartTime = DateTime.Now
                };

                // 按项目整理文件
                await OrganizeByProjectAsync(result);

                // 按数据类型整理文件
                await OrganizeByDataTypeAsync(result);

                // 按时间整理文件
                await OrganizeByTimeAsync(result);

                // 清理重复文件
                await CleanupDuplicateFilesAsync(result);

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"文件整理完成，处理文件数: {result.ProcessedFiles}, 移动文件数: {result.MovedFiles}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"整理导出文件失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 清理过期文件
        /// </summary>
        public async Task<FileCleanupResult> CleanupExpiredFilesAsync(TimeSpan retentionPeriod)
        {
            try
            {
                _logger.LogInformation($"开始清理过期文件，保留期限: {retentionPeriod.TotalDays} 天");

                var result = new FileCleanupResult
                {
                    StartTime = DateTime.Now,
                    RetentionPeriod = retentionPeriod
                };

                var cutoffDate = DateTime.Now.Subtract(retentionPeriod);
                var expiredFiles = new List<string>();

                // 查找过期文件
                await FindExpiredFilesAsync(_baseDirectory, cutoffDate, expiredFiles);

                // 删除过期文件
                foreach (var filePath in expiredFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(filePath);
                        result.TotalSize += fileInfo.Length;
                        
                        File.Delete(filePath);
                        result.DeletedFiles.Add(filePath);
                        result.DeletedCount++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"删除过期文件失败: {filePath}");
                        result.FailedFiles.Add(filePath);
                        result.FailedCount++;
                    }
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"过期文件清理完成，删除文件数: {result.DeletedCount}, 释放空间: {FormatFileSize(result.TotalSize)}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"清理过期文件失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 归档导出文件
        /// </summary>
        public async Task<FileArchiveResult> ArchiveFilesAsync(string archiveName, List<string> filePaths, string archiveFormat = "zip")
        {
            try
            {
                _logger.LogInformation($"开始归档文件: {archiveName}, 文件数: {filePaths.Count}");

                var result = new FileArchiveResult
                {
                    ArchiveName = archiveName,
                    StartTime = DateTime.Now,
                    SourceFileCount = filePaths.Count
                };

                var archivePath = await CreateArchiveAsync(archiveName, filePaths, archiveFormat);
                result.ArchivePath = archivePath;

                // 计算归档文件大小
                if (File.Exists(archivePath))
                {
                    var archiveInfo = new FileInfo(archivePath);
                    result.ArchiveSize = archiveInfo.Length;
                }

                // 如果配置了归档后删除源文件
                if (_config.DeleteSourceAfterArchive)
                {
                    await DeleteSourceFilesAfterArchiveAsync(filePaths, result);
                }

                result.EndTime = DateTime.Now;
                _logger.LogInformation($"文件归档完成: {archivePath}, 大小: {FormatFileSize(result.ArchiveSize)}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"归档文件失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 查找文件
        /// </summary>
        public async Task<List<FileInfo>> FindFilesAsync(FileSearchCriteria criteria)
        {
            try
            {
                var files = new List<FileInfo>();
                await SearchFilesAsync(_baseDirectory, criteria, files);

                _logger.LogInformation($"文件搜索完成，找到 {files.Count} 个文件");
                return files;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"查找文件失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 计算存储信息
        /// </summary>
        private async Task CalculateStorageInfoAsync(FileStorageInfo info)
        {
            var files = Directory.GetFiles(_baseDirectory, "*.*", SearchOption.AllDirectories);
            
            foreach (var file in files)
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    info.TotalFiles++;
                    info.TotalSize += fileInfo.Length;

                    // 按扩展名统计
                    var extension = fileInfo.Extension.ToLower();
                    if (!info.FilesByExtension.ContainsKey(extension))
                    {
                        info.FilesByExtension[extension] = 0;
                    }
                    info.FilesByExtension[extension]++;

                    // 按大小分类
                    if (fileInfo.Length < 1024 * 1024) // 小于1MB
                    {
                        info.SmallFiles++;
                    }
                    else if (fileInfo.Length < 10 * 1024 * 1024) // 小于10MB
                    {
                        info.MediumFiles++;
                    }
                    else
                    {
                        info.LargeFiles++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"计算文件信息失败: {file}");
                }
            }
        }

        /// <summary>
        /// 按项目整理文件
        /// </summary>
        private async Task OrganizeByProjectAsync(FileOrganizationResult result)
        {
            var projectDirs = Directory.GetDirectories(_baseDirectory);
            
            foreach (var projectDir in projectDirs)
            {
                var projectName = Path.GetFileName(projectDir);
                var projectFiles = Directory.GetFiles(projectDir, "*.*", SearchOption.TopDirectoryOnly);

                foreach (var file in projectFiles)
                {
                    try
                    {
                        var fileName = Path.GetFileName(file);
                        if (fileName.Contains(projectName))
                        {
                            result.ProcessedFiles++;
                            // 文件已经在正确的项目目录中，无需移动
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"处理项目文件失败: {file}");
                    }
                }
            }
        }

        /// <summary>
        /// 按数据类型整理文件
        /// </summary>
        private async Task OrganizeByDataTypeAsync(FileOrganizationResult result)
        {
            var allFiles = Directory.GetFiles(_baseDirectory, "*.*", SearchOption.AllDirectories);
            
            foreach (var file in allFiles)
            {
                try
                {
                    var fileName = Path.GetFileName(file);
                    var dataType = ExtractDataTypeFromFileName(fileName);
                    
                    if (!string.IsNullOrEmpty(dataType))
                    {
                        var dataTypeDir = Path.Combine(_baseDirectory, "DataTypes", dataType);
                        if (!Directory.Exists(dataTypeDir))
                        {
                            Directory.CreateDirectory(dataTypeDir);
                        }

                        var targetPath = Path.Combine(dataTypeDir, Path.GetFileName(file));
                        if (file != targetPath && !File.Exists(targetPath))
                        {
                            File.Move(file, targetPath);
                            result.MovedFiles++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"按数据类型整理文件失败: {file}");
                }
            }
        }

        /// <summary>
        /// 按时间整理文件
        /// </summary>
        private async Task OrganizeByTimeAsync(FileOrganizationResult result)
        {
            var allFiles = Directory.GetFiles(_baseDirectory, "*.*", SearchOption.AllDirectories);
            
            foreach (var file in allFiles)
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    var yearMonth = fileInfo.LastWriteTime.ToString("yyyy-MM");
                    var timeDir = Path.Combine(_baseDirectory, "TimeBased", yearMonth);
                    
                    if (!Directory.Exists(timeDir))
                    {
                        Directory.CreateDirectory(timeDir);
                    }

                    var targetPath = Path.Combine(timeDir, Path.GetFileName(file));
                    if (file != targetPath && !File.Exists(targetPath))
                    {
                        File.Move(file, targetPath);
                        result.MovedFiles++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"按时间整理文件失败: {file}");
                }
            }
        }

        /// <summary>
        /// 清理重复文件
        /// </summary>
        private async Task CleanupDuplicateFilesAsync(FileOrganizationResult result)
        {
            var allFiles = Directory.GetFiles(_baseDirectory, "*.*", SearchOption.AllDirectories);
            var fileGroups = allFiles.GroupBy(f => Path.GetFileName(f)).Where(g => g.Count() > 1);

            foreach (var group in fileGroups)
            {
                var files = group.ToList();
                var primaryFile = files.First();
                
                // 保留第一个文件，删除其他重复文件
                for (int i = 1; i < files.Count; i++)
                {
                    try
                    {
                        File.Delete(files[i]);
                        result.DeletedCount++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"删除重复文件失败: {files[i]}");
                    }
                }
            }
        }

        /// <summary>
        /// 查找过期文件
        /// </summary>
        private async Task FindExpiredFilesAsync(string directory, DateTime cutoffDate, List<string> expiredFiles)
        {
            try
            {
                var files = Directory.GetFiles(directory);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.LastWriteTime < cutoffDate)
                    {
                        expiredFiles.Add(file);
                    }
                }

                var subdirs = Directory.GetDirectories(directory);
                foreach (var subdir in subdirs)
                {
                    await FindExpiredFilesAsync(subdir, cutoffDate, expiredFiles);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"查找过期文件失败: {directory}");
            }
        }

        /// <summary>
        /// 创建归档文件
        /// </summary>
        private async Task<string> CreateArchiveAsync(string archiveName, List<string> filePaths, string archiveFormat)
        {
            var archiveDir = Path.Combine(_baseDirectory, "Archives");
            if (!Directory.Exists(archiveDir))
            {
                Directory.CreateDirectory(archiveDir);
            }

            var archivePath = Path.Combine(archiveDir, $"{archiveName}.{archiveFormat}");
            
            // 这里应该实现具体的归档逻辑
            // 由于.NET Core没有内置的压缩功能，这里只是创建了一个占位符
            // 实际使用时可以集成第三方压缩库如SharpCompress或使用系统命令
            
            await File.WriteAllTextAsync(archivePath, $"Archive: {archiveName}\nFiles: {string.Join("\n", filePaths)}");
            
            return archivePath;
        }

        /// <summary>
        /// 归档后删除源文件
        /// </summary>
        private async Task DeleteSourceFilesAfterArchiveAsync(List<string> filePaths, FileArchiveResult result)
        {
            foreach (var filePath in filePaths)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        var fileInfo = new FileInfo(filePath);
                        result.SourceFilesSize += fileInfo.Length;
                        
                        File.Delete(filePath);
                        result.DeletedSourceFiles.Add(filePath);
                        result.DeletedSourceCount++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"归档后删除源文件失败: {filePath}");
                }
            }
        }

        /// <summary>
        /// 搜索文件
        /// </summary>
        private async Task SearchFilesAsync(string directory, FileSearchCriteria criteria, List<FileInfo> results)
        {
            try
            {
                var files = Directory.GetFiles(directory);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    
                    // 应用搜索条件
                    if (criteria.MinSize.HasValue && fileInfo.Length < criteria.MinSize.Value)
                        continue;
                    
                    if (criteria.MaxSize.HasValue && fileInfo.Length > criteria.MaxSize.Value)
                        continue;
                    
                    if (criteria.MinDate.HasValue && fileInfo.LastWriteTime < criteria.MinDate.Value)
                        continue;
                    
                    if (criteria.MaxDate.HasValue && fileInfo.LastWriteTime > criteria.MaxDate.Value)
                        continue;
                    
                    if (!string.IsNullOrEmpty(criteria.FileNamePattern) && 
                        !fileInfo.Name.Contains(criteria.FileNamePattern, StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    if (!string.IsNullOrEmpty(criteria.Extension) && 
                        !fileInfo.Extension.Equals(criteria.Extension, StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    results.Add(fileInfo);
                }

                var subdirs = Directory.GetDirectories(directory);
                foreach (var subdir in subdirs)
                {
                    await SearchFilesAsync(subdir, criteria, results);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"搜索文件失败: {directory}");
            }
        }

        /// <summary>
        /// 从文件名提取数据类型
        /// </summary>
        private string ExtractDataTypeFromFileName(string fileName)
        {
            // 简单的数据类型提取逻辑，可以根据实际命名规则调整
            if (fileName.Contains("Angle"))
                return "Angle";
            if (fileName.Contains("Rainfall"))
                return "Rainfall";
            if (fileName.Contains("ThreeTran"))
                return "ThreeTran";
            if (fileName.Contains("Inclinometer"))
                return "Inclinometer";
            
            return string.Empty;
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;
            
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            
            return $"{len:0.##} {sizes[order]}";
        }
    }

    /// <summary>
    /// 文件管理器配置
    /// </summary>
    public class ExportFileManagerConfig
    {
        /// <summary>
        /// 归档后是否删除源文件
        /// </summary>
        public bool DeleteSourceAfterArchive { get; set; } = false;

        /// <summary>
        /// 是否启用自动整理
        /// </summary>
        public bool EnableAutoOrganization { get; set; } = true;

        /// <summary>
        /// 是否启用重复文件检测
        /// </summary>
        public bool EnableDuplicateDetection { get; set; } = true;

        /// <summary>
        /// 最大文件大小限制（字节）
        /// </summary>
        public long MaxFileSize { get; set; } = 100 * 1024 * 1024; // 100MB

        /// <summary>
        /// 支持的文件扩展名
        /// </summary>
        public List<string> SupportedExtensions { get; set; } = new() { ".xlsx", ".xls", ".csv" };
    }

    /// <summary>
    /// 文件存储信息
    /// </summary>
    public class FileStorageInfo
    {
        /// <summary>
        /// 基础目录
        /// </summary>
        public string BaseDirectory { get; set; } = string.Empty;

        /// <summary>
        /// 检查时间
        /// </summary>
        public DateTime CheckTime { get; set; }

        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 总大小（字节）
        /// </summary>
        public long TotalSize { get; set; }

        /// <summary>
        /// 小文件数量（<1MB）
        /// </summary>
        public int SmallFiles { get; set; }

        /// <summary>
        /// 中等文件数量（1MB-10MB）
        /// </summary>
        public int MediumFiles { get; set; }

        /// <summary>
        /// 大文件数量（>10MB）
        /// </summary>
        public int LargeFiles { get; set; }

        /// <summary>
        /// 按扩展名统计的文件数量
        /// </summary>
        public Dictionary<string, int> FilesByExtension { get; set; } = new();
    }

    /// <summary>
    /// 文件整理结果
    /// </summary>
    public class FileOrganizationResult
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 处理的文件数量
        /// </summary>
        public int ProcessedFiles { get; set; }

        /// <summary>
        /// 移动的文件数量
        /// </summary>
        public int MovedFiles { get; set; }

        /// <summary>
        /// 删除的文件数量
        /// </summary>
        public int DeletedCount { get; set; }

        /// <summary>
        /// 删除的文件列表
        /// </summary>
        public List<string> DeletedFiles { get; set; } = new();
    }

    /// <summary>
    /// 文件清理结果
    /// </summary>
    public class FileCleanupResult
    {
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 保留期限
        /// </summary>
        public TimeSpan RetentionPeriod { get; set; }

        /// <summary>
        /// 删除的文件数量
        /// </summary>
        public int DeletedCount { get; set; }

        /// <summary>
        /// 失败的文件数量
        /// </summary>
        public int FailedCount { get; set; }

        /// <summary>
        /// 总释放空间（字节）
        /// </summary>
        public long TotalSize { get; set; }

        /// <summary>
        /// 删除的文件列表
        /// </summary>
        public List<string> DeletedFiles { get; set; } = new();

        /// <summary>
        /// 失败的文件列表
        /// </summary>
        public List<string> FailedFiles { get; set; } = new();
    }

    /// <summary>
    /// 文件归档结果
    /// </summary>
    public class FileArchiveResult
    {
        /// <summary>
        /// 归档名称
        /// </summary>
        public string ArchiveName { get; set; } = string.Empty;

        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 源文件数量
        /// </summary>
        public int SourceFileCount { get; set; }

        /// <summary>
        /// 归档文件路径
        /// </summary>
        public string ArchivePath { get; set; } = string.Empty;

        /// <summary>
        /// 归档文件大小（字节）
        /// </summary>
        public long ArchiveSize { get; set; }

        /// <summary>
        /// 源文件总大小（字节）
        /// </summary>
        public long SourceFilesSize { get; set; }

        /// <summary>
        /// 删除的源文件数量
        /// </summary>
        public int DeletedSourceCount { get; set; }

        /// <summary>
        /// 删除的源文件列表
        /// </summary>
        public List<string> DeletedSourceFiles { get; set; } = new();
    }

    /// <summary>
    /// 文件搜索条件
    /// </summary>
    public class FileSearchCriteria
    {
        /// <summary>
        /// 最小文件大小（字节）
        /// </summary>
        public long? MinSize { get; set; }

        /// <summary>
        /// 最大文件大小（字节）
        /// </summary>
        public long? MaxSize { get; set; }

        /// <summary>
        /// 最小修改日期
        /// </summary>
        public DateTime? MinDate { get; set; }

        /// <summary>
        /// 最大修改日期
        /// </summary>
        public DateTime? MaxDate { get; set; }

        /// <summary>
        /// 文件名模式（模糊匹配）
        /// </summary>
        public string? FileNamePattern { get; set; }

        /// <summary>
        /// 文件扩展名
        /// </summary>
        public string? Extension { get; set; }
    }
}
