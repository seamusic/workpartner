using System;
using System.Collections.Generic;
using System.Linq;
using DataFixter.Models;
using Serilog;

namespace DataFixter.Excel
{
    /// <summary>
    /// 数据标准化器，负责将Excel读取结果转换为标准的数据模型
    /// </summary>
    public class DataNormalizer
    {
        private readonly ILogger _logger;
        private readonly DataNormalizationOptions _options;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="options">标准化选项</param>
        public DataNormalizer(ILogger logger, DataNormalizationOptions? options = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _options = options ?? new DataNormalizationOptions();
        }

        /// <summary>
        /// 标准化Excel读取结果
        /// </summary>
        /// <param name="readResults">Excel读取结果列表</param>
        /// <returns>标准化后的数据列表</returns>
        public List<PeriodData> NormalizeData(List<ExcelReadResult> readResults)
        {
            var normalizedData = new List<PeriodData>();
            var totalRows = 0;
            var processedRows = 0;
            var skippedRows = 0;

            try
            {
                foreach (var readResult in readResults)
                {
                    if (!readResult.IsSuccess) continue;

                    totalRows += readResult.DataRows.Count;
                    foreach (var excelRow in readResult.DataRows)
                    {
                        try
                        {
                            var normalizedRow = NormalizeRow(excelRow);
                            if (normalizedRow != null)
                            {
                                normalizedData.Add(normalizedRow);
                                processedRows++;
                            }
                            else
                            {
                                skippedRows++;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "标准化行数据失败: 文件 {FileName}, 行号 {RowNumber}", 
                                System.IO.Path.GetFileName(readResult.FilePath), excelRow.RowNumber);
                            skippedRows++;
                        }
                    }
                }

                _logger.Information("数据标准化完成: 总计 {TotalRows} 行, 处理 {ProcessedRows} 行, 跳过 {SkippedRows} 行", 
                    totalRows, processedRows, skippedRows);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "数据标准化过程中发生异常");
            }

            return normalizedData;
        }

        /// <summary>
        /// 标准化单行数据
        /// </summary>
        /// <param name="excelRow">Excel数据行</param>
        /// <returns>标准化后的PeriodData</returns>
        private PeriodData? NormalizeRow(ExcelDataRow excelRow)
        {
            try
            {
                // 验证点名
                if (!ValidatePointName(excelRow.PointName))
                {
                    return null;
                }

                // 验证里程
                if (!ValidateMileage(excelRow.Mileage))
                {
                    return null;
                }

                // 标准化数值数据
                var normalizedData = new PeriodData
                {
                    FileInfo = excelRow.FileInfo,
                    RowNumber = excelRow.RowNumber,
                    PointName = excelRow.PointName?.Trim(),
                    Mileage = NormalizeMileage(excelRow.Mileage),
                    CurrentPeriodX = NormalizeValue(excelRow.CurrentPeriodX, "本期变化量X"),
                    CurrentPeriodY = NormalizeValue(excelRow.CurrentPeriodY, "本期变化量Y"),
                    CurrentPeriodZ = NormalizeValue(excelRow.CurrentPeriodZ, "本期变化量Z"),
                    CumulativeX = NormalizeValue(excelRow.CumulativeX, "累计变化量X"),
                    CumulativeY = NormalizeValue(excelRow.CumulativeY, "累计变化量Y"),
                    CumulativeZ = NormalizeValue(excelRow.CumulativeZ, "累计变化量Z"),
                    DailyX = NormalizeValue(excelRow.DailyX, "日变化量X"),
                    DailyY = NormalizeValue(excelRow.DailyY, "日变化量Y"),
                    DailyZ = NormalizeValue(excelRow.DailyZ, "日变化量Z")
                };

                // 验证数据完整性
                if (!ValidateDataIntegrity(normalizedData))
                {
                    return null;
                }

                return normalizedData;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "标准化行数据失败: 行号 {RowNumber}", excelRow.RowNumber);
                return null;
            }
        }

        /// <summary>
        /// 验证点名
        /// </summary>
        /// <param name="pointName">点名</param>
        /// <returns>是否有效</returns>
        private bool ValidatePointName(string? pointName)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return false;
            }

            var trimmedName = pointName.Trim();
            
            // 检查点名长度
            if (trimmedName.Length < _options.MinPointNameLength || 
                trimmedName.Length > _options.MaxPointNameLength)
            {
                return false;
            }

            // 检查点名格式（可以根据实际需求调整）
            if (_options.ValidatePointNameFormat && 
                !System.Text.RegularExpressions.Regex.IsMatch(trimmedName, _options.PointNamePattern))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 验证里程
        /// </summary>
        /// <param name="mileage">里程</param>
        /// <returns>是否有效</returns>
        private bool ValidateMileage(double mileage)
        {
            return mileage >= _options.MinMileage && mileage <= _options.MaxMileage;
        }

        /// <summary>
        /// 标准化里程
        /// </summary>
        /// <param name="mileage">里程</param>
        /// <returns>标准化后的里程</returns>
        private double NormalizeMileage(double mileage)
        {
            // 保留指定的小数位数
            return Math.Round(mileage, _options.MileageDecimalPlaces);
        }

        /// <summary>
        /// 标准化数值
        /// </summary>
        /// <param name="value">原始值</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns>标准化后的值</returns>
        private double NormalizeValue(double value, string fieldName)
        {
            try
            {
                // 处理无穷大和NaN
                if (double.IsInfinity(value) || double.IsNaN(value))
                {
                    _logger.Debug("检测到无效数值: {FieldName} = {Value}, 设置为0", fieldName, value);
                    return 0.0;
                }

                // 处理超出范围的值
                if (Math.Abs(value) > _options.MaxValueThreshold)
                {
                    _logger.Debug("检测到超出阈值的数值: {FieldName} = {Value}, 限制为阈值", fieldName, value);
                    return value > 0 ? _options.MaxValueThreshold : -_options.MaxValueThreshold;
                }

                // 保留指定的小数位数
                var normalizedValue = Math.Round(value, _options.ValueDecimalPlaces);

                // 处理过小的值（接近0）
                if (Math.Abs(normalizedValue) < _options.MinValueThreshold)
                {
                    normalizedValue = 0.0;
                }

                return normalizedValue;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "标准化数值失败: {FieldName} = {Value}", fieldName, value);
                return 0.0;
            }
        }

        /// <summary>
        /// 验证数据完整性
        /// </summary>
        /// <param name="data">数据</param>
        /// <returns>是否完整</returns>
        private bool ValidateDataIntegrity(PeriodData data)
        {
            try
            {
                // 检查点名是否为空
                if (string.IsNullOrWhiteSpace(data.PointName))
                {
                    return false;
                }

                // 检查里程是否在有效范围内
                if (data.Mileage < _options.MinMileage || data.Mileage > _options.MaxMileage)
                {
                    return false;
                }

                // 检查是否有至少一个方向的数据不为0
                var hasValidData = Math.Abs(data.CurrentPeriodX) > _options.MinValueThreshold ||
                                 Math.Abs(data.CurrentPeriodY) > _options.MinValueThreshold ||
                                 Math.Abs(data.CurrentPeriodZ) > _options.MinValueThreshold ||
                                 Math.Abs(data.CumulativeX) > _options.MinValueThreshold ||
                                 Math.Abs(data.CumulativeY) > _options.MinValueThreshold ||
                                 Math.Abs(data.CumulativeZ) > _options.MinValueThreshold ||
                                 Math.Abs(data.DailyX) > _options.MinValueThreshold ||
                                 Math.Abs(data.DailyY) > _options.MinValueThreshold ||
                                 Math.Abs(data.DailyZ) > _options.MinValueThreshold;

                if (!hasValidData && _options.RequireAtLeastOneValidValue)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "验证数据完整性失败: 点名 {PointName}", data.PointName);
                return false;
            }
        }

        /// <summary>
        /// 获取标准化统计信息
        /// </summary>
        /// <param name="originalData">原始数据</param>
        /// <param name="normalizedData">标准化后的数据</param>
        /// <returns>统计信息</returns>
        public DataNormalizationStatistics GetStatistics(List<ExcelDataRow> originalData, List<PeriodData> normalizedData)
        {
            var statistics = new DataNormalizationStatistics
            {
                TotalOriginalRows = originalData.Count,
                TotalNormalizedRows = normalizedData.Count,
                SkippedRows = originalData.Count - normalizedData.Count,
                SuccessRate = originalData.Count > 0 ? (double)normalizedData.Count / originalData.Count * 100 : 0
            };

            // 统计各字段的异常值
            if (originalData.Count > 0)
            {
                statistics.InvalidPointNames = originalData.Count(r => !ValidatePointName(r.PointName));
                statistics.InvalidMileages = originalData.Count(r => !ValidateMileage(r.Mileage));
                statistics.ExtremeValues = originalData.Count(r => 
                    Math.Abs(r.CurrentPeriodX) > _options.MaxValueThreshold ||
                    Math.Abs(r.CurrentPeriodY) > _options.MaxValueThreshold ||
                    Math.Abs(r.CurrentPeriodZ) > _options.MaxValueThreshold ||
                    Math.Abs(r.CumulativeX) > _options.MaxValueThreshold ||
                    Math.Abs(r.CumulativeY) > _options.MaxValueThreshold ||
                    Math.Abs(r.CumulativeZ) > _options.MaxValueThreshold ||
                    Math.Abs(r.DailyX) > _options.MaxValueThreshold ||
                    Math.Abs(r.DailyY) > _options.MaxValueThreshold ||
                    Math.Abs(r.DailyZ) > _options.MaxValueThreshold);
            }

            return statistics;
        }
    }

    /// <summary>
    /// 数据标准化选项
    /// </summary>
    public class DataNormalizationOptions
    {
        /// <summary>
        /// 点名最小长度
        /// </summary>
        public int MinPointNameLength { get; set; } = 1;

        /// <summary>
        /// 点名最大长度
        /// </summary>
        public int MaxPointNameLength { get; set; } = 100;

        /// <summary>
        /// 是否验证点名格式
        /// </summary>
        public bool ValidatePointNameFormat { get; set; } = false;

        /// <summary>
        /// 点名格式正则表达式
        /// </summary>
        public string PointNamePattern { get; set; } = @"^[\w\u4e00-\u9fa5]+$";

        /// <summary>
        /// 最小里程
        /// </summary>
        public double MinMileage { get; set; } = 0.0;

        /// <summary>
        /// 最大里程
        /// </summary>
        public double MaxMileage { get; set; } = 100000.0;

        /// <summary>
        /// 里程小数位数
        /// </summary>
        public int MileageDecimalPlaces { get; set; } = 2;

        /// <summary>
        /// 数值小数位数
        /// </summary>
        public int ValueDecimalPlaces { get; set; } = 3;

        /// <summary>
        /// 最小数值阈值
        /// </summary>
        public double MinValueThreshold { get; set; } = 0.001;

        /// <summary>
        /// 最大数值阈值
        /// </summary>
        public double MaxValueThreshold { get; set; } = 1000.0;

        /// <summary>
        /// 是否要求至少有一个有效值
        /// </summary>
        public bool RequireAtLeastOneValidValue { get; set; } = true;
    }

    /// <summary>
    /// 数据标准化统计信息
    /// </summary>
    public class DataNormalizationStatistics
    {
        /// <summary>
        /// 原始数据总行数
        /// </summary>
        public int TotalOriginalRows { get; set; }

        /// <summary>
        /// 标准化后数据总行数
        /// </summary>
        public int TotalNormalizedRows { get; set; }

        /// <summary>
        /// 跳过的行数
        /// </summary>
        public int SkippedRows { get; set; }

        /// <summary>
        /// 成功率（百分比）
        /// </summary>
        public double SuccessRate { get; set; }

        /// <summary>
        /// 无效点名的数量
        /// </summary>
        public int InvalidPointNames { get; set; }

        /// <summary>
        /// 无效里程的数量
        /// </summary>
        public int InvalidMileages { get; set; }

        /// <summary>
        /// 极端值的数量
        /// </summary>
        public int ExtremeValues { get; set; }

        /// <summary>
        /// 获取统计摘要
        /// </summary>
        /// <returns>统计摘要字符串</returns>
        public string GetSummary()
        {
            return $"数据标准化统计: 原始{TotalOriginalRows}行, 成功{TotalNormalizedRows}行, 跳过{SkippedRows}行, 成功率{SuccessRate:F1}%";
        }

        /// <summary>
        /// 获取详细统计信息
        /// </summary>
        /// <returns>详细统计信息字符串</returns>
        public string GetDetailedInfo()
        {
            var info = $"数据标准化详细统计:\n";
            info += $"原始数据总行数: {TotalOriginalRows}\n";
            info += $"标准化后数据总行数: {TotalNormalizedRows}\n";
            info += $"跳过的行数: {SkippedRows}\n";
            info += $"成功率: {SuccessRate:F2}%\n";
            info += $"无效点名数量: {InvalidPointNames}\n";
            info += $"无效里程数量: {InvalidMileages}\n";
            info += $"极端值数量: {ExtremeValues}\n";
            return info;
        }
    }
}
