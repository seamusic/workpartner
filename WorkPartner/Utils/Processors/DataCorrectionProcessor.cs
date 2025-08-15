using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;

namespace WorkPartner.Utils.Processors
{
	public class DataCorrectionProcessor : IDataCorrectionProcessor
	{
		public DataCorrectionResult ProcessDataCorrection(string originalDirectory, string processedDirectory, DataProcessorConfig? config = null)
		{
			return DataProcessor.ProcessDataCorrection(originalDirectory, processedDirectory, config);
		}
	}
}
