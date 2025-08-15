using WorkPartner.Models;

namespace WorkPartner.Utils.Interfaces
{
	public interface IDataCorrectionProcessor
	{
		DataCorrectionResult ProcessDataCorrection(string originalDirectory, string processedDirectory, DataProcessorConfig? config = null);
	}
}
