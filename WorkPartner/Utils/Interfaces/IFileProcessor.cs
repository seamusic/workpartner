using WorkPartner.Models;

namespace WorkPartner.Utils.Interfaces
{
    /// <summary>
    /// 文件处理接口
    /// </summary>
    public interface IFileProcessor
    {
        /// <summary>
        /// 比较原始文件和已处理文件的数值差异
        /// </summary>
        /// <param name="originalDirectory">原始文件目录路径</param>
        /// <param name="processedDirectory">已处理文件目录路径</param>
        /// <param name="config">配置参数</param>
        /// <param name="showDetailedDifferences">是否显示详细的差异信息</param>
        /// <param name="tolerance">数值比较容差</param>
        /// <param name="maxDifferencesToShow">最大显示差异数量</param>
        /// <returns>比较结果统计</returns>
        ComparisonResult CompareOriginalAndProcessedFiles(
            string originalDirectory, 
            string processedDirectory, 
            DataProcessorConfig? config = null,
            bool showDetailedDifferences = false,
            double? tolerance = null,
            int maxDifferencesToShow = -1);

        /// <summary>
        /// 检查输出结果目录中第7、8、9列绝对值超过阈值的数据
        /// </summary>
        /// <param name="outputDirectory">输出结果目录路径</param>
        /// <param name="threshold">阈值，默认为4</param>
        /// <returns>检查结果</returns>
        LargeValueCheckResult CheckLargeValuesInOutputDirectory(string outputDirectory, double threshold = 4.0);
    }
}
