using System.IO;

namespace WorkPartner.Services
{
    public class FileService : IFileService
    {
        public List<string> ScanExcelFiles(string folderPath)
        {
            var excelFiles = new List<string>();
            var supportedExtensions = new[] { ".xlsx", ".xls" };

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    throw new DirectoryNotFoundException($"目录不存在: {folderPath}");
                }

                foreach (var file in Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
                {
                    var extension = Path.GetExtension(file).ToLower();
                    if (supportedExtensions.Contains(extension))
                    {
                        excelFiles.Add(file);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"获取Excel文件列表失败: {folderPath}", ex);
            }

            return excelFiles;
        }

        public async Task<List<string>> ScanExcelFilesAsync(string folderPath)
        {
            return await Task.Run(() => ScanExcelFiles(folderPath));
        }

        public bool CreateDirectory(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"创建目录失败: {path}", ex);
            }
        }

        public async Task<bool> CreateDirectoryAsync(string path)
        {
            return await Task.Run(() => CreateDirectory(path));
        }

        public bool DirectoryExists(string path)
        {
            return Directory.Exists(path);
        }

        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        public bool DeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"删除文件失败: {path}", ex);
            }
        }

        public async Task<bool> DeleteFileAsync(string path)
        {
            return await Task.Run(() => DeleteFile(path));
        }

        public bool CopyFile(string sourcePath, string targetPath, bool overwrite = false)
        {
            try
            {
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetDirectory) && !Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                File.Copy(sourcePath, targetPath, overwrite);
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"复制文件失败: {sourcePath} -> {targetPath}", ex);
            }
        }

        public async Task<bool> CopyFileAsync(string sourcePath, string targetPath, bool overwrite = false)
        {
            return await Task.Run(() => CopyFile(sourcePath, targetPath, overwrite));
        }

        public bool MoveFile(string sourcePath, string targetPath, bool overwrite = false)
        {
            try
            {
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetDirectory) && !Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                if (overwrite && File.Exists(targetPath))
                {
                    File.Delete(targetPath);
                }

                File.Move(sourcePath, targetPath);
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"移动文件失败: {sourcePath} -> {targetPath}", ex);
            }
        }

        public async Task<bool> MoveFileAsync(string sourcePath, string targetPath, bool overwrite = false)
        {
            return await Task.Run(() => MoveFile(sourcePath, targetPath, overwrite));
        }

        public long GetFileSize(string path)
        {
            try
            {
                var fileInfo = new FileInfo(path);
                return fileInfo.Exists ? fileInfo.Length : 0;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"获取文件大小失败: {path}", ex);
            }
        }

        public FileInfo? GetFileInfo(string path)
        {
            try
            {
                var fileInfo = new FileInfo(path);
                return fileInfo.Exists ? fileInfo : null;
            }
            catch
            {
                return null;
            }
        }

        public DirectoryInfo? GetDirectoryInfo(string path)
        {
            try
            {
                var dirInfo = new DirectoryInfo(path);
                return dirInfo.Exists ? dirInfo : null;
            }
            catch
            {
                return null;
            }
        }

        public bool CheckFileAccess(string path, FileAccess access)
        {
            try
            {
                using var stream = File.Open(path, FileMode.Open, access, FileShare.None);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public long GetAvailableDiskSpace(string path)
        {
            try
            {
                var driveInfo = new DriveInfo(Path.GetPathRoot(path));
                return driveInfo.AvailableFreeSpace;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"获取磁盘空间失败: {path}", ex);
            }
        }

        public int CleanupTempFiles(string folderPath, string pattern = "*.tmp", TimeSpan? olderThan = null)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                    return 0;

                var files = Directory.GetFiles(folderPath, pattern);
                var deletedCount = 0;

                foreach (var file in files)
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        var shouldDelete = true;

                        if (olderThan.HasValue)
                        {
                            var cutoffTime = DateTime.Now - olderThan.Value;
                            shouldDelete = fileInfo.LastWriteTime < cutoffTime;
                        }

                        if (shouldDelete)
                        {
                            File.Delete(file);
                            deletedCount++;
                        }
                    }
                    catch
                    {
                        // 忽略单个文件的删除错误
                    }
                }

                return deletedCount;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"清理临时文件失败: {folderPath}", ex);
            }
        }
    }
}