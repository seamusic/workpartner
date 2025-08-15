using WorkPartner.Models;

namespace WorkPartner.Utils.Interfaces
{
    /// <summary>
    /// 数据处理接口
    /// </summary>
    public interface IDataProcessor
    {
        /// <summary>
        /// 处理缺失数据
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <param name="config">配置参数</param>
        /// <returns>处理后的文件列表</returns>
        List<ExcelFile> ProcessMissingData(List<ExcelFile> files, DataProcessorConfig? config = null);

        /// <summary>
        /// 处理缺失数据（保持向后兼容）
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>处理后的文件列表</returns>
        List<ExcelFile> ProcessMissingData(List<ExcelFile> files);
    }
}
