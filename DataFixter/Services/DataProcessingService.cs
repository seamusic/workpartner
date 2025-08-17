using DataFixter.Excel;
using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 数据处理服务，负责Excel数据的验证、修正和统计
    /// </summary>
    public class DataProcessingService
    {
        private readonly ExcelProcessor _excelProcessor;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="excelProcessor">Excel处理器实例</param>
        public DataProcessingService(ExcelProcessor excelProcessor)
        {
            _excelProcessor = excelProcessor ?? throw new ArgumentNullException(nameof(excelProcessor));
        }

        /// <summary>
        /// 验证数据完整性
        /// </summary>
        /// <param name="startRow">开始行索引（从0开始）</param>
        /// <param name="endRow">结束行索引（包含）</param>
        /// <param name="requiredColumns">必需列索引数组</param>
        /// <returns>验证结果</returns>
        public DataValidationResult ValidateData(int startRow, int endRow, int[] requiredColumns)
        {
            var result = new DataValidationResult
            {
                StartRow = startRow,
                EndRow = endRow,
                TotalRows = endRow - startRow + 1,
                ValidRows = 0,
                InvalidRows = 0,
                MissingDataRows = new List<int>(),
                InvalidDataRows = new List<int>()
            };

            try
            {
                Log.Information("开始验证数据完整性: 行{StartRow}到{EndRow}", startRow, endRow);

                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    var rowValidation = ValidateRow(rowIndex, requiredColumns);
                    
                    if (rowValidation.IsValid)
                    {
                        result.ValidRows++;
                    }
                    else
                    {
                        result.InvalidRows++;
                        
                        if (rowValidation.HasMissingData)
                        {
                            result.MissingDataRows.Add(rowIndex);
                        }
                        
                        if (rowValidation.HasInvalidData)
                        {
                            result.InvalidDataRows.Add(rowIndex);
                        }
                    }
                }

                Log.Information("数据验证完成: 总行数{Total}, 有效行{Valid}, 无效行{Invalid}", 
                    result.TotalRows, result.ValidRows, result.InvalidRows);

                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "数据验证过程中发生错误");
                throw;
            }
        }

        /// <summary>
        /// 验证单行数据
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="requiredColumns">必需列索引数组</param>
        /// <returns>行验证结果</returns>
        private RowValidationResult ValidateRow(int rowIndex, int[] requiredColumns)
        {
            var result = new RowValidationResult
            {
                RowIndex = rowIndex,
                IsValid = true,
                HasMissingData = false,
                HasInvalidData = false,
                MissingColumns = new List<int>(),
                InvalidColumns = new List<int>()
            };

            foreach (var columnIndex in requiredColumns)
            {
                var cellValue = _excelProcessor.GetCellValue(rowIndex, columnIndex);
                
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    result.HasMissingData = true;
                    result.MissingColumns.Add(columnIndex);
                    result.IsValid = false;
                }
                else if (!IsValidData(cellValue, columnIndex))
                {
                    result.HasInvalidData = true;
                    result.InvalidColumns.Add(columnIndex);
                    result.IsValid = false;
                }
            }

            return result;
        }

        /// <summary>
        /// 验证数据有效性
        /// </summary>
        /// <param name="value">数据值</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>是否有效</returns>
        private bool IsValidData(string value, int columnIndex)
        {
            // 这里可以根据具体业务需求添加数据验证逻辑
            // 例如：日期格式、数字范围、字符串长度等
            
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            // 示例：第0列（A列）应该是非空字符串
            if (columnIndex == 0)
            {
                return value.Trim().Length > 0;
            }

            // 示例：第1列（B列）应该是数字
            if (columnIndex == 1)
            {
                return double.TryParse(value, out _);
            }

            // 默认情况下，非空数据认为是有效的
            return true;
        }

        /// <summary>
        /// 修正缺失数据
        /// </summary>
        /// <param name="missingRows">缺失数据的行索引列表</param>
        /// <param name="missingColumns">缺失数据的列索引列表</param>
        /// <param name="defaultValues">默认值字典，键为列索引，值为默认值</param>
        /// <returns>修正结果</returns>
        public DataCorrectionResult CorrectMissingData(List<int> missingRows, List<int> missingColumns, Dictionary<int, string> defaultValues)
        {
            var result = new DataCorrectionResult
            {
                TotalCorrections = 0,
                SuccessfulCorrections = 0,
                FailedCorrections = 0,
                CorrectionDetails = new List<CorrectionDetail>()
            };

            try
            {
                Log.Information("开始修正缺失数据: {RowCount}行, {ColumnCount}列", missingRows.Count, missingColumns.Count);

                foreach (var rowIndex in missingRows)
                {
                    foreach (var columnIndex in missingColumns)
                    {
                        var detail = new CorrectionDetail
                        {
                            RowIndex = rowIndex,
                            ColumnIndex = columnIndex,
                            OriginalValue = _excelProcessor.GetCellValue(rowIndex, columnIndex),
                            NewValue = defaultValues.GetValueOrDefault(columnIndex, string.Empty)
                        };

                        if (_excelProcessor.SetCellValue(rowIndex, columnIndex, detail.NewValue))
                        {
                            detail.IsSuccessful = true;
                            result.SuccessfulCorrections++;
                        }
                        else
                        {
                            detail.IsSuccessful = false;
                            result.FailedCorrections++;
                        }

                        result.CorrectionDetails.Add(detail);
                        result.TotalCorrections++;
                    }
                }

                Log.Information("缺失数据修正完成: 总计{Total}, 成功{Success}, 失败{Failed}", 
                    result.TotalCorrections, result.SuccessfulCorrections, result.FailedCorrections);

                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "修正缺失数据过程中发生错误");
                throw;
            }
        }

        /// <summary>
        /// 生成数据统计报告
        /// </summary>
        /// <param name="startRow">开始行索引</param>
        /// <param name="endRow">结束行索引</param>
        /// <returns>统计报告</returns>
        public DataStatisticsReport GenerateStatistics(int startRow, int endRow)
        {
            var report = new DataStatisticsReport
            {
                StartRow = startRow,
                EndRow = endRow,
                TotalRows = endRow - startRow + 1,
                ColumnStatistics = new Dictionary<int, ColumnStats>()
            };

            try
            {
                Log.Information("开始生成数据统计报告: 行{StartRow}到{EndRow}", startRow, endRow);

                // 获取列数
                var columnCount = _excelProcessor.GetColumnCount(startRow);
                
                // 初始化每列的统计信息
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    report.ColumnStatistics[colIndex] = new ColumnStats
                    {
                        ColumnIndex = colIndex,
                        NonEmptyCount = 0,
                        EmptyCount = 0,
                        UniqueValues = new HashSet<string>(),
                        MinLength = int.MaxValue,
                        MaxLength = 0
                    };
                }

                // 统计每列的数据
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < columnCount; colIndex++)
                    {
                        var cellValue = _excelProcessor.GetCellValue(rowIndex, colIndex);
                        var stats = report.ColumnStatistics[colIndex];

                        if (string.IsNullOrWhiteSpace(cellValue))
                        {
                            stats.EmptyCount++;
                        }
                        else
                        {
                            stats.NonEmptyCount++;
                            stats.UniqueValues.Add(cellValue.Trim());
                            
                            var length = cellValue.Length;
                            stats.MinLength = Math.Min(stats.MinLength, length);
                            stats.MaxLength = Math.Max(stats.MaxLength, length);
                        }
                    }
                }

                // 处理最小长度（如果没有非空数据，设为0）
                foreach (var stats in report.ColumnStatistics.Values)
                {
                    if (stats.MinLength == int.MaxValue)
                    {
                        stats.MinLength = 0;
                    }
                }

                Log.Information("数据统计报告生成完成: 总行数{Total}, 总列数{Columns}", 
                    report.TotalRows, columnCount);

                return report;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "生成数据统计报告过程中发生错误");
                throw;
            }
        }
    }

    #region 数据模型

    /// <summary>
    /// 数据验证结果
    /// </summary>
    public class DataValidationResult
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int TotalRows { get; set; }
        public int ValidRows { get; set; }
        public int InvalidRows { get; set; }
        public List<int> MissingDataRows { get; set; } = new();
        public List<int> InvalidDataRows { get; set; } = new();
    }

    /// <summary>
    /// 行验证结果
    /// </summary>
    public class RowValidationResult
    {
        public int RowIndex { get; set; }
        public bool IsValid { get; set; }
        public bool HasMissingData { get; set; }
        public bool HasInvalidData { get; set; }
        public List<int> MissingColumns { get; set; } = new();
        public List<int> InvalidColumns { get; set; } = new();
    }

    /// <summary>
    /// 数据修正结果
    /// </summary>
    public class DataCorrectionResult
    {
        public int TotalCorrections { get; set; }
        public int SuccessfulCorrections { get; set; }
        public int FailedCorrections { get; set; }
        public List<CorrectionDetail> CorrectionDetails { get; set; } = new();
    }

    /// <summary>
    /// 修正详情
    /// </summary>
    public class CorrectionDetail
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string? OriginalValue { get; set; }
        public string NewValue { get; set; } = string.Empty;
        public bool IsSuccessful { get; set; }
    }

    /// <summary>
    /// 数据统计报告
    /// </summary>
    public class DataStatisticsReport
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int TotalRows { get; set; }
        public Dictionary<int, ColumnStats> ColumnStatistics { get; set; } = new();
    }

    /// <summary>
    /// 列统计信息
    /// </summary>
    public class ColumnStats
    {
        public int ColumnIndex { get; set; }
        public int NonEmptyCount { get; set; }
        public int EmptyCount { get; set; }
        public HashSet<string> UniqueValues { get; set; } = new();
        public int MinLength { get; set; }
        public int MaxLength { get; set; }
        public int UniqueValueCount => UniqueValues.Count;
    }

    #endregion
}
