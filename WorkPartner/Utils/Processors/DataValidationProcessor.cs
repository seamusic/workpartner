using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;

namespace WorkPartner.Utils.Processors
{
	/// <summary>
	/// 数据验证处理器（阶段1：包装调用现有实现）
	/// </summary>
	public class DataValidationProcessor : IValidationProcessor
	{
		public List<ExcelFile> ValidateAndRecalculateColumns456(List<ExcelFile> files, DataProcessorConfig? config = null)
		{
			return DataProcessor.ValidateAndRecalculateColumns456(files, config);
		}

		public void ValidateDataIntegrity(List<ExcelFile> files)
		{
			if (files == null) return;
			foreach (var file in files)
			{
				foreach (var dataRow in file.DataRows)
				{
					for (int i = 0; i < dataRow.Values.Count; i++)
					{
						if (dataRow.Values[i].HasValue)
						{
							var value = dataRow.Values[i].Value;
							if (double.IsNaN(value) || double.IsInfinity(value))
							{
								throw new InvalidOperationException($"Invalid data detected in file {file.FileName}, row {dataRow.Name}, column {i}: {value}");
							}
						}
					}
				}
			}
		}

		public CompletenessCheckResult CheckCompleteness(List<ExcelFile> files)
		{
			return DataProcessor.CheckCompleteness(files);
		}

		public DataQualityReport ValidateDataQuality(List<ExcelFile> files)
		{
			return DataProcessor.ValidateDataQuality(files);
		}
	}
}
