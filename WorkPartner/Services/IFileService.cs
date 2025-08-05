using WorkPartner.Models;

namespace WorkPartner.Services
{
    /// <summary>
    /// 文件服务接口
    /// </summary>
    public interface IFileService
    {
        /// <summary>
        /// 扫描指定目录下的Excel文件
        /// </summary>
        /// <param name="folderPath">目录路径</param>
        /// <returns>Excel文件列表</returns>
        List<string> ScanExcelFiles(string folderPath);

        /// <summary>
        /// 扫描指定目录下的Excel文件（异步版本）
        /// </summary>
        /// <param name="folderPath">目录路径</param>
        /// <returns>Excel文件列表</returns>
        Task<List<string>> ScanExcelFilesAsync(string folderPath);

        /// <summary>
        /// 创建目录
        /// </summary>
        /// <param name="path">目录路径</param>
        /// <returns>是否创建成功</returns>
        bool CreateDirectory(string path);

        /// <summary>
        /// 创建目录（异步版本）
        /// </summary>
        /// <param name="path">目录路径</param>
        /// <returns>是否创建成功</returns>
        Task<bool> CreateDirectoryAsync(string path);

        /// <summary>
        /// 检查目录是否存在
        /// </summary>
        /// <param name="path">目录路径</param>
        /// <returns>是否存在</returns>
        bool DirectoryExists(string path);

        /// <summary>
        /// 检查文件是否存在
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>是否存在</returns>
        bool FileExists(string path);

        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>是否删除成功</returns>
        bool DeleteFile(string path);

        /// <summary>
        /// 删除文件（异步版本）
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>是否删除成功</returns>
        Task<bool> DeleteFileAsync(string path);

        /// <summary>
        /// 复制文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="overwrite">是否覆盖</param>
        /// <returns>是否复制成功</returns>
        bool CopyFile(string sourcePath, string targetPath, bool overwrite = false);

        /// <summary>
        /// 复制文件（异步版本）
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="overwrite">是否覆盖</param>
        /// <returns>是否复制成功</returns>
        Task<bool> CopyFileAsync(string sourcePath, string targetPath, bool overwrite = false);

        /// <summary>
        /// 移动文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="overwrite">是否覆盖</param>
        /// <returns>是否移动成功</returns>
        bool MoveFile(string sourcePath, string targetPath, bool overwrite = false);

        /// <summary>
        /// 移动文件（异步版本）
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <param name="overwrite">是否覆盖</param>
        /// <returns>是否移动成功</returns>
        Task<bool> MoveFileAsync(string sourcePath, string targetPath, bool overwrite = false);

        /// <summary>
        /// 获取文件大小
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>文件大小（字节）</returns>
        long GetFileSize(string path);

        /// <summary>
        /// 获取文件信息
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>文件信息</returns>
        FileInfo? GetFileInfo(string path);

        /// <summary>
        /// 获取目录信息
        /// </summary>
        /// <param name="path">目录路径</param>
        /// <returns>目录信息</returns>
        DirectoryInfo? GetDirectoryInfo(string path);

        /// <summary>
        /// 检查文件权限
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="access">访问权限</param>
        /// <returns>是否有权限</returns>
        bool CheckFileAccess(string path, FileAccess access);

        /// <summary>
        /// 获取可用磁盘空间
        /// </summary>
        /// <param name="path">路径</param>
        /// <returns>可用空间（字节）</returns>
        long GetAvailableDiskSpace(string path);

        /// <summary>
        /// 清理临时文件
        /// </summary>
        /// <param name="folderPath">目录路径</param>
        /// <param name="pattern">文件模式</param>
        /// <param name="olderThan">删除早于此时间的文件</param>
        /// <returns>删除的文件数量</returns>
        int CleanupTempFiles(string folderPath, string pattern = "*.tmp", TimeSpan? olderThan = null);
    }
} 