using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataFixter.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.Extensions.Logging;

namespace DataFixter.Excel
{
    /// <summary>
    /// Excel文件批量读取器
    /// 支持读取指定目录下的所有.xls文件，实现从第5行到第364行的数据提取
    /// </summary>
    public class ExcelBatchReader
    {
        private readonly ILogger<ExcelBatchReader> _logger;
        private readonly string _directoryPath;
        private readonly string _fileExtension;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="directoryPath">目录路径</param>
        /// <param name="logger">日志记录器</param>
        /// <param name="fileExtension">文件扩展名，默认为.xls</param>
        public ExcelBatchReader(string directoryPath, ILogger<ExcelBatchReader> logger, string fileExtension = "*.xls")
        {
            _directoryPath = directoryPath ?? throw new ArgumentNullException(nameof(directoryPath));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _fileExtension = fileExtension;
        }

        /// <summary>
        /// 读取目录下的所有Excel文件
        /// </summary>
        /// <returns>所有文件的读取结果</returns>
        public List<ExcelReadResult> ReadAllFiles()
        {
            var results = new List<ExcelReadResult>();
            
            try
            {
                if (!Directory.Exists(_directoryPath))
                {
                    _logger.LogError("目录不存在: {DirectoryPath}", _directoryPath);
                    return results;
                }

                var files = Directory.GetFiles(_directoryPath, _fileExtension, SearchOption.TopDirectoryOnly);
                _logger.LogInformation("找到 {FileCount} 个Excel文件", files.Length);

                foreach (var filePath in files)
                {
                    try
                    {
                        var result = ReadSingleFile(filePath);
                        results.Add(result);
                        
                        if (result.IsSuccess)
                        {
                            _logger.LogInformation("成功读取文件: {FileName}, 数据行数: {RowCount}", 
                                Path.GetFileName(filePath), result.DataRows.Count);
                        }
                        else
                        {
                            _logger.LogWarning("读取文件失败: {FileName}, 错误: {Error}", 
                                Path.GetFileName(filePath), result.ErrorMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "读取文件时发生异常: {FilePath}", filePath);
                        results.Add(new ExcelReadResult
                        {
                            FilePath = filePath,
                            IsSuccess = false,
                            ErrorMessage = ex.Message
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "批量读取Excel文件时发生异常");
            }

            return results;
        }

        /// <summary>
        /// 读取单个Excel文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>读取结果</returns>
        public ExcelReadResult ReadSingleFile(string filePath)
        {
            var result = new ExcelReadResult
            {
                FilePath = filePath,
                IsSuccess = false
            };

            try
            {
                if (!File.Exists(filePath))
                {
                    result.ErrorMessage = "文件不存在";
                    return result;
                }

                var systemFileInfo = new System.IO.FileInfo(filePath);
                if (systemFileInfo.Length == 0)
                {
                    result.ErrorMessage = "文件为空";
                    return result;
                }

                using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                var workbook = new HSSFWorkbook(fileStream);
                
                if (workbook.NumberOfSheets == 0)
                {
                    result.ErrorMessage = "工作簿中没有工作表";
                    return result;
                }

                var sheet = workbook.GetSheetAt(0); // 读取第一个工作表
                var dataRows = ExtractDataRows(sheet, filePath);
                
                result.DataRows = dataRows;
                result.IsSuccess = true;
                result.RowCount = dataRows.Count;
                result.SheetName = sheet.SheetName;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "读取Excel文件失败: {FilePath}", filePath);
            }

            return result;
        }

        /// <summary>
        /// 提取数据行（第5行到第364行）
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="filePath">文件路径</param>
        /// <returns>数据行列表</returns>
        private List<ExcelDataRow> ExtractDataRows(ISheet sheet, string filePath)
        {
            var dataRows = new List<ExcelDataRow>();
            var fileInfo = GetFileInfo(filePath);

            // 从第5行开始读取（索引为4）
            for (int rowIndex = 4; rowIndex <= 363 && rowIndex < sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                var dataRow = ExtractRowData(row, rowIndex + 1, fileInfo);
                if (dataRow != null)
                {
                    dataRows.Add(dataRow);
                }
            }

            return dataRows;
        }

        /// <summary>
        /// 提取行数据
        /// </summary>
        /// <param name="row">Excel行</param>
        /// <param name="rowNumber">行号（从1开始）</param>
        /// <param name="fileInfo">文件信息</param>
        /// <returns>行数据</returns>
        private ExcelDataRow? ExtractRowData(IRow row, int rowNumber, Models.FileInfo fileInfo)
        {
            try
            {
                // 检查点名列（第1列，索引为0）
                var pointNameCell = row.GetCell(0);
                if (pointNameCell == null || string.IsNullOrWhiteSpace(GetCellValue(pointNameCell)))
                {
                    return null; // 跳过空行
                }

                var dataRow = new ExcelDataRow
                {
                    RowNumber = rowNumber,
                    FileInfo = fileInfo,
                    PointName = GetCellValue(pointNameCell)?.Trim(),
                    Mileage = GetNumericCellValue(row.GetCell(1)), // 里程列
                    CurrentPeriodX = GetNumericCellValue(row.GetCell(2)), // 本期变化量X
                    CurrentPeriodY = GetNumericCellValue(row.GetCell(3)), // 本期变化量Y
                    CurrentPeriodZ = GetNumericCellValue(row.GetCell(4)), // 本期变化量Z
                    CumulativeX = GetNumericCellValue(row.GetCell(5)), // 累计变化量X
                    CumulativeY = GetNumericCellValue(row.GetCell(6)), // 累计变化量Y
                    CumulativeZ = GetNumericCellValue(row.GetCell(7)), // 累计变化量Z
                    DailyX = GetNumericCellValue(row.GetCell(8)), // 日变化量X
                    DailyY = GetNumericCellValue(row.GetCell(9)), // 日变化量Y
                    DailyZ = GetNumericCellValue(row.GetCell(10)) // 日变化量Z
                };

                // 验证点名不为空
                if (string.IsNullOrWhiteSpace(dataRow.PointName))
                {
                    return null;
                }

                return dataRow;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "提取行数据失败: 行号 {RowNumber}", rowNumber);
                return null;
            }
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>单元格值</returns>
        private string? GetCellValue(ICell? cell)
        {
            if (cell == null) return null;

            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Formula:
                        return cell.CellFormula;
                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取数值单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>数值</returns>
        private double GetNumericCellValue(ICell? cell)
        {
            if (cell == null) return 0.0;

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue;
                    case CellType.String:
                        if (double.TryParse(cell.StringCellValue, out var result))
                            return result;
                        return 0.0;
                    case CellType.Formula:
                        try
                        {
                            return cell.NumericCellValue;
                        }
                        catch
                        {
                            return 0.0;
                        }
                    default:
                        return 0.0;
                }
            }
            catch
            {
                return 0.0;
            }
        }

        /// <summary>
        /// 获取文件信息
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>文件信息</returns>
        private Models.FileInfo GetFileInfo(string filePath)
        {
            var systemFileInfo = new System.IO.FileInfo(filePath);
            return new Models.FileInfo(filePath, systemFileInfo.Length, systemFileInfo.LastWriteTime);
        }
    }

    /// <summary>
    /// Excel读取结果
    /// </summary>
    public class ExcelReadResult
    {
        /// <summary>
        /// 文件路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 是否读取成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string? SheetName { get; set; }

        /// <summary>
        /// 数据行列表
        /// </summary>
        public List<ExcelDataRow> DataRows { get; set; } = new List<ExcelDataRow>();

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; set; }
    }

    /// <summary>
    /// Excel数据行
    /// </summary>
    public class ExcelDataRow
    {
        /// <summary>
        /// 行号（从1开始）
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 文件信息
        /// </summary>
        public Models.FileInfo FileInfo { get; set; } = null!;

        /// <summary>
        /// 点名
        /// </summary>
        public string? PointName { get; set; }

        /// <summary>
        /// 里程
        /// </summary>
        public double Mileage { get; set; }

        /// <summary>
        /// 本期变化量（X方向）
        /// </summary>
        public double CurrentPeriodX { get; set; }

        /// <summary>
        /// 本期变化量（Y方向）
        /// </summary>
        public double CurrentPeriodY { get; set; }

        /// <summary>
        /// 本期变化量（Z方向）
        /// </summary>
        public double CurrentPeriodZ { get; set; }

        /// <summary>
        /// 累计变化量（X方向）
        /// </summary>
        public double CumulativeX { get; set; }

        /// <summary>
        /// 累计变化量（Y方向）
        /// </summary>
        public double CumulativeY { get; set; }

        /// <summary>
        /// 累计变化量（Z方向）
        /// </summary>
        public double CumulativeZ { get; set; }

        /// <summary>
        /// 日变化量（X方向）
        /// </summary>
        public double DailyX { get; set; }

        /// <summary>
        /// 日变化量（Y方向）
        /// </summary>
        public double DailyY { get; set; }

        /// <summary>
        /// 日变化量（Z方向）
        /// </summary>
        public double DailyZ { get; set; }

        /// <summary>
        /// 转换为PeriodData对象
        /// </summary>
        /// <returns>PeriodData对象</returns>
        public PeriodData ToPeriodData()
        {
            return new PeriodData
            {
                FileInfo = FileInfo,
                RowNumber = RowNumber,
                PointName = PointName,
                Mileage = Mileage,
                CurrentPeriodX = CurrentPeriodX,
                CurrentPeriodY = CurrentPeriodY,
                CurrentPeriodZ = CurrentPeriodZ,
                CumulativeX = CumulativeX,
                CumulativeY = CumulativeY,
                CumulativeZ = CumulativeZ,
                DailyX = DailyX,
                DailyY = DailyY,
                DailyZ = DailyZ
            };
        }
    }
}
