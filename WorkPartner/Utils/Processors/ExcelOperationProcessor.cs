using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;

namespace WorkPartner.Utils.Processors
{
	public class ExcelOperationProcessor : IExcelOperationProcessor
	{
		public int UpdateA2ColumnForAllFiles(System.Collections.Generic.List<ExcelFile> files, string outputDirectory)
		{
			return DataProcessor.UpdateA2ColumnForAllFiles(files, outputDirectory);
		}
	}
}
