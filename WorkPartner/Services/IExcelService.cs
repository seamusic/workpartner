using WorkPartner.Models;

namespace WorkPartner.Services
{
    /// <summary>
    /// Excel服务接口
    /// </summary>
    public interface IExcelService
    {
        /// <summary>
        /// 读取Excel文件数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Excel文件数据</returns>
        Task<ExcelFile> ReadExcelFileAsync(string filePath);

        /// <summary>
        /// 读取Excel文件数据（同步版本）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Excel文件数据</returns>
        ExcelFile ReadExcelFile(string filePath);

        /// <summary>
        /// 保存Excel文件
        /// </summary>
        /// <param name="excelFile">Excel文件数据</param>
        /// <param name="outputPath">输出路径</param>
        /// <returns>是否保存成功</returns>
        Task<bool> SaveExcelFileAsync(ExcelFile excelFile, string outputPath);

        /// <summary>
        /// 保存Excel文件（同步版本）
        /// </summary>
        /// <param name="excelFile">Excel文件数据</param>
        /// <param name="outputPath">输出路径</param>
        /// <returns>是否保存成功</returns>
        bool SaveExcelFile(ExcelFile excelFile, string outputPath);

        /// <summary>
        /// 验证Excel文件格式
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否格式正确</returns>
        bool ValidateExcelFile(string filePath);

        /// <summary>
        /// 获取Excel文件信息
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>文件信息</returns>
        ExcelFile GetExcelFileInfo(string filePath);

        /// <summary>
        /// 复制Excel文件
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>是否复制成功</returns>
        bool CopyExcelFile(string sourcePath, string targetPath);

        /// <summary>
        /// 检查Excel文件是否被占用
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否被占用</returns>
        bool IsFileLocked(string filePath);
    }
} 