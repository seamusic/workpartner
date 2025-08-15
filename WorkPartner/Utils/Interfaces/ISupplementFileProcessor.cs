using WorkPartner.Models;

namespace WorkPartner.Utils.Interfaces
{
    /// <summary>
    /// 补充文件处理接口
    /// </summary>
    public interface ISupplementFileProcessor
    {
        /// <summary>
        /// 生成补充文件列表
        /// </summary>
        /// <param name="files">现有文件列表</param>
        /// <returns>需要补充的文件列表</returns>
        List<SupplementFileInfo> GenerateSupplementFiles(List<ExcelFile> files);

        /// <summary>
        /// 创建补充文件
        /// </summary>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>创建成功的文件数量</returns>
        int CreateSupplementFiles(List<SupplementFileInfo> supplementFiles, string outputDirectory);

        /// <summary>
        /// 创建补充文件并修改A2列数据内容
        /// </summary>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="allFiles">所有文件列表（用于确定上期观测时间）</param>
        /// <returns>创建的文件数量</returns>
        int CreateSupplementFilesWithA2Update(List<SupplementFileInfo> supplementFiles, string outputDirectory, List<ExcelFile> allFiles);

        /// <summary>
        /// 获取所有需要处理的文件（包括原始文件和补充文件）
        /// </summary>
        /// <param name="originalFiles">原始文件列表</param>
        /// <param name="supplementFiles">补充文件信息列表</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <returns>所有需要处理的文件列表</returns>
        List<ExcelFile> GetAllFilesForProcessing(List<ExcelFile> originalFiles, List<SupplementFileInfo> supplementFiles, string outputDirectory);
		string GetPreviousObservationTime(SupplementFileInfo supplementFile, List<ExcelFile> allFiles);
    }
}
