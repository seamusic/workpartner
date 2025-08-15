using WorkPartner.Models;

namespace WorkPartner.Utils.Interfaces
{
    /// <summary>
    /// 数据验证处理接口
    /// </summary>
    public interface IValidationProcessor
    {
        /// <summary>
        /// 验证并重新计算第4、5、6列的值
        /// </summary>
        /// <param name="files">已处理的文件列表</param>
        /// <param name="config">配置参数</param>
        /// <returns>处理后的文件列表</returns>
        List<ExcelFile> ValidateAndRecalculateColumns456(List<ExcelFile> files, DataProcessorConfig? config = null);

        /// <summary>
        /// 验证数据完整性
        /// </summary>
        /// <param name="files">文件列表</param>
        void ValidateDataIntegrity(List<ExcelFile> files);

        /// <summary>
        /// 检查数据完整性
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>完整性检查结果</returns>
        CompletenessCheckResult CheckCompleteness(List<ExcelFile> files);

        /// <summary>
        /// 验证数据质量
        /// </summary>
        /// <param name="files">文件列表</param>
        /// <returns>数据质量报告</returns>
        DataQualityReport ValidateDataQuality(List<ExcelFile> files);
    }
}
