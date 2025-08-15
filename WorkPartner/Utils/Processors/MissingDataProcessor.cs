using WorkPartner.Models;
using WorkPartner.Utils.Interfaces;
using System.Collections.Concurrent;

namespace WorkPartner.Utils.Processors
{
	/// <summary>
	/// 缺失数据处理器（阶段2：承载核心实现，DataProcessor提供薄适配层）
	/// </summary>
	public class MissingDataProcessor : IDataProcessor
	{
		public List<ExcelFile> ProcessMissingData(List<ExcelFile> files, DataProcessorConfig? config = null)
		{
			config ??= DataProcessorConfig.Default;
			return DataProcessor.ProcessMissingData(files, config);
		}

		public List<ExcelFile> ProcessMissingData(List<ExcelFile> files)
		{
			return ProcessMissingData(files, DataProcessorConfig.Default);
		}

		// 暴露核心阶段方法，供适配层委托（后续可逐步完全迁移实现）
		public List<MissingPeriod> IdentifyMissingPeriods(List<ExcelFile> files, DataProcessorConfig config)
		{
			return WorkPartner.Utils.DataProcessor.IdentifyMissingPeriodsPure(files, config);
		}

		public MissingProcessingStats ProcessMissingPeriods(
			List<ExcelFile> files, List<MissingPeriod> periods, DataProcessorConfig config, DataCache? cache, PerformanceMetrics metrics)
		{
			// 复用 DataProcessor 的并行/串行实现以保持一致；后续可完全内聚迁移
			// 这里直接调用总入口，确保统计与并行策略一致
			return WorkPartner.Utils.DataProcessor.ProcessConsecutiveMissingDataOptimized(files, config, cache, metrics);
		}

		// ===== 下沉的核心实现（供 DataProcessor 代理调用）=====
		internal static Dictionary<DateTime, int> BuildTimeIndexMap(List<ExcelFile> files)
		{
			var map = new Dictionary<DateTime, int>();
			for (int i = 0; i < files.Count; i++)
			{
				var key = files[i].Date.Date.AddHours(files[i].Hour);
				if (!map.ContainsKey(key))
				{
					map[key] = i;
				}
			}
			return map;
		}

		internal static Dictionary<string, Dictionary<int, List<int>>> BuildValidIndicesMap(List<ExcelFile> files)
		{
			var result = new Dictionary<string, Dictionary<int, List<int>>>();
			for (int i = 0; i < files.Count; i++)
			{
				var file = files[i];
				foreach (var dataRow in file.DataRows)
				{
					if (!result.TryGetValue(dataRow.Name, out var colMap))
					{
						colMap = new Dictionary<int, List<int>>();
						result[dataRow.Name] = colMap;
					}

					for (int valueIndex = 0; valueIndex < dataRow.Values.Count; valueIndex++)
					{
						if (!colMap.TryGetValue(valueIndex, out var list))
						{
							list = new List<int>();
							colMap[valueIndex] = list;
						}
						if (dataRow.Values[valueIndex].HasValue)
						{
							list.Add(i);
						}
					}
				}
			}
			return result;
		}

		internal static MissingProcessingStats ProcessMissingPeriodOptimized(
			MissingPeriod period,
			List<ExcelFile> files,
			Dictionary<DateTime, int> timeToIndex,
			Dictionary<string, Dictionary<int, List<int>>> validIndexMap,
			DataProcessorConfig config,
			DataCache? cache,
			Action<string>? log = null)
		{
			var logger = log ?? Console.WriteLine;
			var stats = new MissingProcessingStats();
			var totalTimePoints = period.MissingTimes.Count;
			var totalDataRows = period.MissingDataRows.Count;
			var lastProgressTime = DateTime.Now;

			// 为所有文件预构建名称到数据行的映射，避免在时间点循环内重复构建
			var rowsByFile = new List<Dictionary<string, DataRow>>(files.Count);
			for (int fi = 0; fi < files.Count; fi++)
			{
				rowsByFile.Add(files[fi].DataRows.ToDictionary(r => r.Name));
			}

			logger($"  📅 处理时间段: {period.StartTime:yyyy-MM-dd HH:mm} 到 {period.EndTime:yyyy-MM-dd HH:mm}, 共 {totalTimePoints} 个时间点, {totalDataRows} 个数据行");

			for (int timeIndex = 0; timeIndex < period.MissingTimes.Count; timeIndex++)
			{
				var missingTime = period.MissingTimes[timeIndex];
				var missingTimeKey = missingTime.ToString("yyyyMMddHH");
				var currentTime = DateTime.Now;

				if ((timeIndex + 1) % 5 == 0 || (currentTime - lastProgressTime).TotalSeconds >= 20)
				{
					var timeProgress = (double)(timeIndex + 1) / totalTimePoints * 100;
					logger($"    ⏰ 时间点进度: {timeIndex + 1}/{totalTimePoints} ({timeProgress:F1}%) - 当前处理: {missingTime:yyyy-MM-dd HH:mm}");
					lastProgressTime = currentTime;
				}

				if (!timeToIndex.TryGetValue(missingTime, out var targetIdx))
				{
					continue;
				}
				var targetFile = files[targetIdx];
				var rowByName = rowsByFile[targetIdx];

				foreach (var dataRowName in period.MissingDataRows)
				{
					rowByName.TryGetValue(dataRowName, out var dataRow);
					if (dataRow == null) continue;

					int changeColumns = (config.ChangeColumnsPerRow ?? (dataRow.Values.Count / 2));
					var cacheKeyRowTimePrefix = string.Concat("missing_period_", dataRowName, "_", missingTimeKey, "_");

					if (config.EnableParallelValueColumns && changeColumns > 1)
					{
						Parallel.For(0, changeColumns, valueIndex =>
						{
							if (dataRow.Values[valueIndex].HasValue) return;

							var cacheKey = cacheKeyRowTimePrefix + valueIndex.ToString();
							if (cache != null)
							{
								var cachedValue = cache.Get<double?>(cacheKey);
								if (cachedValue.HasValue)
								{
									dataRow.Values[valueIndex] = cachedValue.Value;
									stats.CacheHits++;
									return;
								}
							}

							var (previousValue, nextValue) = WorkPartner.Utils.DataProcessor.GetNearestValuesUsingIndex(
								files, timeToIndex, validIndexMap, rowsByFile, dataRowName, missingTime, valueIndex);

							var cumulativeColumnIndex = valueIndex + changeColumns;

							if (previousValue.HasValue && nextValue.HasValue)
							{
								var missingPoint = new MissingDataPoint
								{
									DataRowName = dataRowName,
									ValueIndex = valueIndex,
									TimePoint = missingTime,
									PreviousValue = previousValue,
									NextValue = nextValue,
									BaseValue = (previousValue.Value + nextValue.Value) / 2
								};

								var adjustedValue = WorkPartner.Utils.DataProcessor.CalculateAdjustedValueForMissingPoint(missingPoint, period, config);
								dataRow.Values[valueIndex] = adjustedValue;

								if (cumulativeColumnIndex < dataRow.Values.Count && !dataRow.Values[cumulativeColumnIndex].HasValue)
								{
									var previousPeriodValue = WorkPartner.Utils.DataProcessor.GetPreviousPeriodValueUsingIndex(
										files, timeToIndex, validIndexMap, rowsByFile, dataRowName, missingTime, cumulativeColumnIndex);
									if (previousPeriodValue.HasValue)
									{
										var newCumulativeValue = previousPeriodValue.Value + adjustedValue;
										dataRow.Values[cumulativeColumnIndex] = newCumulativeValue;
									}
								}

								cache?.Set(cacheKey, adjustedValue);
								stats.Processings++;
							}
							else
							{
								stats.CacheMisses++;
							}
						});
					}
					else
					{
						for (int valueIndex = 0; valueIndex < changeColumns; valueIndex++)
						{
							if (dataRow.Values[valueIndex].HasValue) continue;

							var cacheKey = cacheKeyRowTimePrefix + valueIndex.ToString();
							if (cache != null)
							{
								var cachedValue = cache.Get<double?>(cacheKey);
								if (cachedValue.HasValue)
								{
									dataRow.Values[valueIndex] = cachedValue.Value;
									stats.CacheHits++;
									continue;
								}
							}

							var (previousValue, nextValue) = WorkPartner.Utils.DataProcessor.GetNearestValuesUsingIndex(
								files, timeToIndex, validIndexMap, rowsByFile, dataRowName, missingTime, valueIndex);

							var cumulativeColumnIndex = valueIndex + changeColumns;

							if (previousValue.HasValue && nextValue.HasValue)
							{
								var missingPoint = new MissingDataPoint
								{
									DataRowName = dataRowName,
									ValueIndex = valueIndex,
									TimePoint = missingTime,
									PreviousValue = previousValue,
									NextValue = nextValue,
									BaseValue = (previousValue.Value + nextValue.Value) / 2
								};

								var adjustedValue = WorkPartner.Utils.DataProcessor.CalculateAdjustedValueForMissingPoint(missingPoint, period, config);
								dataRow.Values[valueIndex] = adjustedValue;

								if (cumulativeColumnIndex < dataRow.Values.Count && !dataRow.Values[cumulativeColumnIndex].HasValue)
								{
									var previousPeriodValue = WorkPartner.Utils.DataProcessor.GetPreviousPeriodValueUsingIndex(
										files, timeToIndex, validIndexMap, rowsByFile, dataRowName, missingTime, cumulativeColumnIndex);
									if (previousPeriodValue.HasValue)
									{
										var newCumulativeValue = previousPeriodValue.Value + adjustedValue;
										dataRow.Values[cumulativeColumnIndex] = newCumulativeValue;
									}
								}

								cache?.Set(cacheKey, adjustedValue);
								stats.Processings++;
							}
							else
							{
								stats.CacheMisses++;
							}
						}
					}
				}
			}

			logger($"    ✅ 时间段处理完成: 补充了 {stats.Processings} 个缺失值, 缓存命中 {stats.CacheHits} 次, 缓存未命中 {stats.CacheMisses} 次");
			return stats;
		}
	}
}
