using WorkPartner.Utils.Interfaces;
using WorkPartner.Utils.Processors;

namespace WorkPartner.Utils.DependencyInjection
{
	public static class ServiceCollectionExtensions
	{
		public static (IDataProcessor dataProcessor,
			IValidationProcessor validationProcessor,
			IFileProcessor fileProcessor,
			ISupplementFileProcessor supplementFileProcessor,
			IExcelOperationProcessor excelOperationProcessor,
			IDataCorrectionProcessor dataCorrectionProcessor)
			CreateDefaultProcessors()
		{
			return (
				dataProcessor: new MissingDataProcessor(),
				validationProcessor: new DataValidationProcessor(),
				fileProcessor: new FileComparisonProcessor(),
				supplementFileProcessor: new SupplementFileProcessor(),
				excelOperationProcessor: new ExcelOperationProcessor(),
				dataCorrectionProcessor: new DataCorrectionProcessor()
			);
		}
	}
}
