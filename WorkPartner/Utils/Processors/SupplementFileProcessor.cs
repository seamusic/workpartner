using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;

namespace WorkPartner.Utils.Processors
{
	public class SupplementFileProcessor : ISupplementFileProcessor
	{
		public List<SupplementFileInfo> GenerateSupplementFiles(List<ExcelFile> files)
		{
			return DataProcessor.GenerateSupplementFiles(files);
		}

		public int CreateSupplementFiles(List<SupplementFileInfo> supplementFiles, string outputDirectory)
		{
			return DataProcessor.CreateSupplementFiles(supplementFiles, outputDirectory);
		}

		public int CreateSupplementFilesWithA2Update(List<SupplementFileInfo> supplementFiles, string outputDirectory, List<ExcelFile> allFiles)
		{
			return DataProcessor.CreateSupplementFilesWithA2Update(supplementFiles, outputDirectory, allFiles);
		}

		public List<ExcelFile> GetAllFilesForProcessing(List<ExcelFile> originalFiles, List<SupplementFileInfo> supplementFiles, string outputDirectory)
		{
			return DataProcessor.GetAllFilesForProcessing(originalFiles, supplementFiles, outputDirectory);
		}

		public string GetPreviousObservationTime(SupplementFileInfo supplementFile, List<ExcelFile> allFiles)
		{
			return DataProcessor.GetPreviousObservationTime(supplementFile, allFiles);
		}
	}
}
