namespace WorkPartner.Utils.Interfaces
{
	public interface IExcelOperationProcessor
	{
		int UpdateA2ColumnForAllFiles(System.Collections.Generic.List<WorkPartner.Models.ExcelFile> files, string outputDirectory);
	}
}
