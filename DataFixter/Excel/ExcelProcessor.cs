using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Serilog;

namespace DataFixter.Excel
{
    /// <summary>
    /// Excel文件处理器，负责Excel文件的读取、写入和数据处理
    /// </summary>
    public class ExcelProcessor : IDisposable
    {
        private IWorkbook? _workbook;
        private ISheet? _currentSheet;
        private readonly string _filePath;
        private bool _disposed = false;

        /// <summary>
        /// 当前工作簿
        /// </summary>
        public IWorkbook? Workbook => _workbook;

        /// <summary>
        /// 当前工作表
        /// </summary>
        public ISheet? CurrentSheet => _currentSheet;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        public ExcelProcessor(string filePath)
        {
            _filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <returns>是否成功打开</returns>
        public bool OpenFile()
        {
            try
            {
                if (!File.Exists(_filePath))
                {
                    Log.Error("Excel文件不存在: {FilePath}", _filePath);
                    return false;
                }

                using var fileStream = new FileStream(_filePath, FileMode.Open, FileAccess.Read);
                
                // 根据文件扩展名选择合适的工作簿类型
                var extension = Path.GetExtension(_filePath).ToLower();
                switch (extension)
                {
                    case ".xls":
                        _workbook = new HSSFWorkbook(fileStream);
                        break;
                    case ".xlsx":
                        _workbook = new XSSFWorkbook(fileStream);
                        break;
                    default:
                        Log.Error("不支持的Excel文件格式: {Extension}", extension);
                        return false;
                }
                
                // 默认选择第一个工作表
                if (_workbook.NumberOfSheets > 0)
                {
                    _currentSheet = _workbook.GetSheetAt(0);
                }

                Log.Information("成功打开Excel文件: {FilePath}, 工作表数量: {SheetCount}", 
                    _filePath, _workbook.NumberOfSheets);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "打开Excel文件失败: {FilePath}", _filePath);
                return false;
            }
        }

        /// <summary>
        /// 选择工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <returns>是否成功选择</returns>
        public bool SelectSheet(int sheetIndex)
        {
            try
            {
                if (_workbook == null)
                {
                    Log.Warning("工作簿未打开，无法选择工作表");
                    return false;
                }

                if (sheetIndex < 0 || sheetIndex >= _workbook.NumberOfSheets)
                {
                    Log.Warning("工作表索引超出范围: {Index}, 总工作表数: {Total}", 
                        sheetIndex, _workbook.NumberOfSheets);
                    return false;
                }

                _currentSheet = _workbook.GetSheetAt(sheetIndex);
                Log.Information("已选择工作表: {Index}, 名称: {Name}", 
                    sheetIndex, _currentSheet.SheetName);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "选择工作表失败: {Index}", sheetIndex);
                return false;
            }
        }

        /// <summary>
        /// 选择工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>是否成功选择</returns>
        public bool SelectSheet(string sheetName)
        {
            try
            {
                if (_workbook == null)
                {
                    Log.Warning("工作簿未打开，无法选择工作表");
                    return false;
                }

                var sheet = _workbook.GetSheet(sheetName);
                if (sheet == null)
                {
                    Log.Warning("未找到工作表: {Name}", sheetName);
                    return false;
                }

                _currentSheet = sheet;
                Log.Information("已选择工作表: {Name}", sheetName);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "选择工作表失败: {Name}", sheetName);
                return false;
            }
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>单元格值</returns>
        public string? GetCellValue(int rowIndex, int columnIndex)
        {
            try
            {
                if (_currentSheet == null)
                {
                    Log.Warning("未选择工作表，无法获取单元格值");
                    return null;
                }

                var row = _currentSheet.GetRow(rowIndex);
                if (row == null)
                {
                    return null;
                }

                var cell = row.GetCell(columnIndex);
                if (cell == null)
                {
                    return null;
                }

                return GetCellValueAsString(cell);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取单元格值失败: 行{Row}, 列{Column}", rowIndex, columnIndex);
                return null;
            }
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="value">要设置的值</param>
        /// <returns>是否成功设置</returns>
        public bool SetCellValue(int rowIndex, int columnIndex, string value)
        {
            try
            {
                if (_currentSheet == null)
                {
                    Log.Warning("未选择工作表，无法设置单元格值");
                    return false;
                }

                var row = _currentSheet.GetRow(rowIndex) ?? _currentSheet.CreateRow(rowIndex);
                var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
                
                cell.SetCellValue(value);
                
                Log.Debug("已设置单元格值: 行{Row}, 列{Column}, 值: {Value}", rowIndex, columnIndex, value);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "设置单元格值失败: 行{Row}, 列{Column}, 值: {Value}", rowIndex, columnIndex, value);
                return false;
            }
        }

        /// <summary>
        /// 获取行数
        /// </summary>
        /// <returns>行数</returns>
        public int GetRowCount()
        {
            if (_currentSheet == null)
            {
                return 0;
            }

            return _currentSheet.LastRowNum + 1;
        }

        /// <summary>
        /// 获取列数
        /// </summary>
        /// <param name="rowIndex">行索引，默认为0</param>
        /// <returns>列数</returns>
        public int GetColumnCount(int rowIndex = 0)
        {
            if (_currentSheet == null)
            {
                return 0;
            }

            var row = _currentSheet.GetRow(rowIndex);
            return row?.LastCellNum ?? 0;
        }

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="outputPath">输出路径，如果为null则覆盖原文件</param>
        /// <returns>是否成功保存</returns>
        public bool SaveFile(string? outputPath = null)
        {
            try
            {
                if (_workbook == null)
                {
                    Log.Warning("工作簿未打开，无法保存文件");
                    return false;
                }

                var savePath = outputPath ?? _filePath;
                
                // 确保输出目录存在
                var outputDir = Path.GetDirectoryName(savePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                using var fileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write);
                _workbook.Write(fileStream);
                
                Log.Information("文件保存成功: {Path}", savePath);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "保存文件失败: {Path}", outputPath ?? _filePath);
                return false;
            }
        }

        /// <summary>
        /// 将单元格值转换为字符串
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>字符串值</returns>
        private static string GetCellValueAsString(ICell cell)
        {
            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue ?? string.Empty,
                CellType.Numeric => cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Formula => cell.CellFormula ?? string.Empty,
                CellType.Blank => string.Empty,
                _ => string.Empty
            };
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="disposing">是否正在释放托管资源</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _workbook?.Close();
                _workbook = null;
                _currentSheet = null;
                _disposed = true;
            }
        }
    }
}
