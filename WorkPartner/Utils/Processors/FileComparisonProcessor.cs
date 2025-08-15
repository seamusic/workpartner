using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;

namespace WorkPartner.Utils.Processors
{
	/// <summary>
	/// 文件比较处理器（阶段1：包装调用现有实现）
	/// </summary>
	public class FileComparisonProcessor : IFileProcessor
	{
		public ComparisonResult CompareOriginalAndProcessedFiles(
			string originalDirectory,
			string processedDirectory,
			DataProcessorConfig? config = null,
			bool showDetailedDifferences = false,
			double? tolerance = null,
			int maxDifferencesToShow = -1)
		{
			return DataProcessor.CompareOriginalAndProcessedFiles(originalDirectory, processedDirectory, config, showDetailedDifferences, tolerance, maxDifferencesToShow);
		}

		public LargeValueCheckResult CheckLargeValuesInOutputDirectory(string outputDirectory, double threshold = 4.0)
		{
			return DataProcessor.CheckLargeValuesInOutputDirectory(outputDirectory, threshold);
		}
	}
}
