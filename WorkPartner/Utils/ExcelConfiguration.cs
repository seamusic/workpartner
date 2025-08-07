using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Configuration;
using WorkPartner.Utils;

namespace WorkPartner.Utils
{
    /// <summary>
    /// Excel配置管理类
    /// 负责管理Excel文件读取的各种配置参数
    /// </summary>
    public class ExcelConfiguration
    {
        private static ExcelConfiguration _instance;
        private static readonly object _lock = new object();
        private readonly IConfiguration _configuration;
        
        // 默认配置值
        private const int DEFAULT_START_ROW = 5;
        private const int DEFAULT_END_ROW = 368;
        private const int DEFAULT_START_COL = 4; // D列
        private const int DEFAULT_END_COL = 9;   // I列
        private const int DEFAULT_NAME_COL = 2;  // B列
        
        // 配置属性
        public int StartRow { get; set; } = DEFAULT_START_ROW;
        public int EndRow { get; set; } = DEFAULT_END_ROW;
        public int StartCol { get; set; } = DEFAULT_START_COL;
        public int EndCol { get; set; } = DEFAULT_END_COL;
        public int NameCol { get; set; } = DEFAULT_NAME_COL;
        
        // 计算属性
        public int TotalRows => EndRow - StartRow + 1;
        public int TotalCols => EndCol - StartCol + 1;
        
        // NPOI索引转换（0基索引）
        public int NpoiStartRow => StartRow - 1;
        public int NpoiEndRow => EndRow - 1;
        public int NpoiStartCol => StartCol - 1;
        public int NpoiEndCol => EndCol - 1;
        public int NpoiNameCol => NameCol - 1;
        
        private ExcelConfiguration()
        {
            // 构建配置
            _configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();
            
            LoadConfiguration();
        }
        
        /// <summary>
        /// 单例模式获取配置实例
        /// </summary>
        public static ExcelConfiguration Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ExcelConfiguration();
                        }
                    }
                }
                return _instance;
            }
        }
        
        /// <summary>
        /// 从配置文件加载配置
        /// </summary>
        public void LoadConfiguration()
        {
            try
            {
                // 从配置节读取Excel配置
                var excelSection = _configuration.GetSection("Excel");
                
                if (excelSection.Exists())
                {
                    StartRow = excelSection.GetValue<int>("StartRow", DEFAULT_START_ROW);
                    EndRow = excelSection.GetValue<int>("EndRow", DEFAULT_END_ROW);
                    StartCol = excelSection.GetValue<int>("StartCol", DEFAULT_START_COL);
                    EndCol = excelSection.GetValue<int>("EndCol", DEFAULT_END_COL);
                    NameCol = excelSection.GetValue<int>("NameCol", DEFAULT_NAME_COL);
                    
                    Logger.Info("从appsettings.json加载Excel配置");
                    Logger.Info($"配置参数: StartRow={StartRow}, EndRow={EndRow}, StartCol={StartCol}, EndCol={EndCol}, NameCol={NameCol}");
                }
                else
                {
                    Logger.Info("未找到Excel配置节，使用默认配置");
                    ResetToDefault();
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"加载Excel配置失败，使用默认配置: {ex.Message}");
                ResetToDefault();
            }
        }
        
        /// <summary>
        /// 从Excel文件动态读取配置
        /// 通过分析Excel文件的结构来确定配置参数
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>是否成功读取配置</returns>
        public bool LoadConfigurationFromExcel(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return LoadConfigurationFromXlsx(filePath);
                }
                else if (extension == ".xls")
                {
                    return LoadConfigurationFromXls(filePath);
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Logger.Warning($"从Excel文件读取配置失败: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 从XLSX文件读取配置
        /// </summary>
        private bool LoadConfigurationFromXlsx(string filePath)
        {
            using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null) return false;
            
            // 分析工作表结构来确定配置
            var config = AnalyzeWorksheetStructure(worksheet);
            
            if (config != null)
            {
                StartRow = config.StartRow;
                EndRow = config.EndRow;
                StartCol = config.StartCol;
                EndCol = config.EndCol;
                NameCol = config.NameCol;
                
                Logger.Info($"从XLSX文件动态读取配置: StartRow={StartRow}, EndRow={EndRow}, StartCol={StartCol}, EndCol={EndCol}, NameCol={NameCol}");
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// 从XLS文件读取配置
        /// </summary>
        private bool LoadConfigurationFromXls(string filePath)
        {
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fileStream);
            var sheet = workbook.GetSheetAt(0);
            
            if (sheet == null) return false;
            
            // 分析工作表结构来确定配置
            var config = AnalyzeWorksheetStructure(sheet);
            
            if (config != null)
            {
                StartRow = config.StartRow;
                EndRow = config.EndRow;
                StartCol = config.StartCol;
                EndCol = config.EndCol;
                NameCol = config.NameCol;
                
                Logger.Info($"从XLS文件动态读取配置: StartRow={StartRow}, EndRow={EndRow}, StartCol={StartCol}, EndCol={EndCol}, NameCol={NameCol}");
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// 分析EPPlus工作表结构
        /// </summary>
        private ExcelConfiguration AnalyzeWorksheetStructure(OfficeOpenXml.ExcelWorksheet worksheet)
        {
            var config = new ExcelConfiguration();
            
            // 查找数据区域的开始和结束
            int startRow = 1, endRow = 1, startCol = 1, endCol = 1;
            bool foundData = false;
            
            // 扫描工作表找到数据区域
            var maxRow = worksheet.Dimension?.End?.Row ?? 1000;
            var maxCol = worksheet.Dimension?.End?.Column ?? 10;
            
            for (int row = 1; row <= maxRow; row++)
            {
                for (int col = 1; col <= maxCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    if (cell?.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        if (!foundData)
                        {
                            startRow = row;
                            startCol = col;
                            foundData = true;
                        }
                        endRow = Math.Max(endRow, row);
                        endCol = Math.Max(endCol, col);
                    }
                }
            }
            
            if (foundData)
            {
                // 根据数据区域确定配置
                config.StartRow = startRow;
                config.EndRow = endRow;
                config.StartCol = startCol;
                config.EndCol = endCol;
                config.NameCol = startCol; // 假设第一列是名称列
                
                return config;
            }
            
            return null;
        }
        
        /// <summary>
        /// 分析NPOI工作表结构
        /// </summary>
        private ExcelConfiguration AnalyzeWorksheetStructure(NPOI.SS.UserModel.ISheet sheet)
        {
            var config = new ExcelConfiguration();
            
            // 查找数据区域的开始和结束
            int startRow = 1, endRow = 1, startCol = 1, endCol = 1;
            bool foundData = false;
            
            // 扫描工作表找到数据区域
            for (int row = 0; row < sheet.LastRowNum; row++)
            {
                var sheetRow = sheet.GetRow(row);
                if (sheetRow != null)
                {
                    for (int col = 0; col < sheetRow.LastCellNum; col++)
                    {
                        var cell = sheetRow.GetCell(col);
                        if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                        {
                            if (!foundData)
                            {
                                startRow = row + 1; // 转换为1基索引
                                startCol = col + 1;
                                foundData = true;
                            }
                            endRow = Math.Max(endRow, row + 1);
                            endCol = Math.Max(endCol, col + 1);
                        }
                    }
                }
            }
            
            if (foundData)
            {
                // 根据数据区域确定配置
                config.StartRow = startRow;
                config.EndRow = endRow;
                config.StartCol = startCol;
                config.EndCol = endCol;
                config.NameCol = startCol; // 假设第一列是名称列
                
                return config;
            }
            
            return null;
        }
        
        /// <summary>
        /// 重置为默认配置
        /// </summary>
        public void ResetToDefault()
        {
            StartRow = DEFAULT_START_ROW;
            EndRow = DEFAULT_END_ROW;
            StartCol = DEFAULT_START_COL;
            EndCol = DEFAULT_END_COL;
            NameCol = DEFAULT_NAME_COL;
            
            Logger.Info("重置为默认Excel配置");
        }
        
        /// <summary>
        /// 验证配置的有效性
        /// </summary>
        /// <returns>配置是否有效</returns>
        public bool ValidateConfiguration()
        {
            var isValid = StartRow > 0 && EndRow >= StartRow && 
                         StartCol > 0 && EndCol >= StartCol && 
                         NameCol > 0;
            
            if (!isValid)
            {
                Logger.Warning("Excel配置无效，重置为默认配置");
                ResetToDefault();
            }
            
            return isValid;
        }
        
        /// <summary>
        /// 获取配置摘要信息
        /// </summary>
        /// <returns>配置摘要字符串</returns>
        public string GetConfigurationSummary()
        {
            return $"Excel配置: 数据行范围[{StartRow}-{EndRow}], 数据列范围[{StartCol}-{EndCol}], 名称列[{NameCol}], 总行数[{TotalRows}], 总列数[{TotalCols}]";
        }
    }
} 