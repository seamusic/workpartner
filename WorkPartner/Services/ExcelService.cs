using WorkPartner.Models;
using WorkPartner.Utils;
using OfficeOpenXml;
using NPOI.HSSF.UserModel; // for .xls files
using NPOI.SS.UserModel;
using System.IO;
using System.Threading;

namespace WorkPartner.Services
{
    public class ExcelService : IExcelService
    {
        public ExcelService()
        {
            // 设置EPPlus许可证模式
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<ExcelFile> ReadExcelFileAsync(string filePath)
        {
            return await Task.Run(() => ReadExcelFile(filePath));
        }

        public ExcelFile ReadExcelFile(string filePath)
        {
            return ExceptionHandler.HandleFileRead(() =>
            {
                using var operation = Logger.StartOperation("读取Excel文件", Path.GetFileName(filePath));
                Logger.StartFileProcessing(Path.GetFileName(filePath), "读取");

                var fileInfo = new FileInfo(filePath);
                var excelFile = new ExcelFile
                {
                    FilePath = filePath,
                    FileName = fileInfo.Name,
                    FileSize = fileInfo.Length,
                    LastModified = fileInfo.LastWriteTime,
                    IsLocked = IsFileLocked(filePath),
                    IsValid = ValidateExcelFile(filePath)
                };

                Logger.Validation("文件格式", excelFile.IsValid, excelFile.IsValid ? "Excel格式有效" : "Excel格式无效");
                Logger.Validation("文件可访问性", !excelFile.IsLocked, excelFile.IsLocked ? "文件被占用" : "文件可访问");

                if (!excelFile.IsValid)
                {
                    throw new WorkPartnerException("InvalidFileFormat", $"Excel文件格式无效: {Path.GetFileName(filePath)}", filePath);
                }

                if (excelFile.IsLocked)
                {
                    throw new WorkPartnerException("FileLocked", $"Excel文件被占用: {Path.GetFileName(filePath)}", filePath);
                }

                // 读取Excel数据
                var dataRows = new List<DataRow>();
                var extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".xlsx")
                {
                    using var package = new ExcelPackage(fileInfo);
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet == null)
                    {
                        throw new InvalidOperationException("Excel文件中没有找到工作表");
                    }

                    // 获取Excel配置
                    var config = ExcelConfiguration.Instance;
                    
                    // 读取配置指定行的数据名称
                    for (int row = config.StartRow; row <= config.EndRow; row++)
                    {
                        var nameCell = worksheet.Cells[row, config.NameCol]; // 配置的名称列
                        var name = nameCell?.Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(name))
                        {
                            var dataRow = new DataRow
                            {
                                Name = name,
                                RowIndex = row
                            };

                            // 读取配置指定列的数据值
                            for (int col = config.StartCol; col <= config.EndCol; col++) // 配置的数据列范围
                            {
                                var valueCell = worksheet.Cells[row, col];
                                var value = valueCell?.Value;

                                if (value != null && double.TryParse(value.ToString(), out double numericValue))
                                {
                                    dataRow.AddValue(numericValue);
                                }
                                else
                                {
                                    dataRow.AddValue(null); // 空值或非数值
                                }
                            }

                            dataRows.Add(dataRow);
                        }
                    }
                }
                else if (extension == ".xls")
                {
                    using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    var workbook = new HSSFWorkbook(fileStream);
                    var sheet = workbook.GetSheetAt(0);

                    if (sheet == null)
                    {
                        throw new InvalidOperationException("Excel文件中没有找到工作表");
                    }

                    // 获取Excel配置
                    var config = ExcelConfiguration.Instance;
                    
                    // 读取配置指定行的数据名称（NPOI使用0基索引）
                    for (int row = config.NpoiStartRow; row <= config.NpoiEndRow; row++)
                    {
                        var nameRow = sheet.GetRow(row);
                        if (nameRow == null) continue;

                        var nameCell = nameRow.GetCell(config.NpoiNameCol); // 配置的名称列
                        var name = nameCell?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(name))
                        {
                            var dataRow = new DataRow
                            {
                                Name = name,
                                RowIndex = row + 1 // 转换为1基索引显示
                            };

                            // 读取配置指定列的数据值（NPOI使用0基索引）
                            for (int col = config.NpoiStartCol; col <= config.NpoiEndCol; col++) // 配置的数据列范围
                            {
                                var valueCell = nameRow.GetCell(col);
                                var value = valueCell?.ToString();

                                if (value != null && double.TryParse(value, out double numericValue))
                                {
                                    dataRow.AddValue(numericValue);
                                }
                                else
                                {
                                    dataRow.AddValue(null); // 空值或非数值
                                }
                            }

                            dataRows.Add(dataRow);
                        }
                    }
                }
                else
                {
                    throw new InvalidOperationException($"不支持的文件格式: {extension}");
                }

                // 检查数据行数长短不一致的情况
                if (dataRows.Count > 0)
                {
                    var config = ExcelConfiguration.Instance;
                    var expectedRowCount = config.TotalRows; // 使用配置计算预期行数
                    if (dataRows.Count != expectedRowCount)
                    {
                        var message = $"警告：文件 {Path.GetFileName(filePath)} 的数据行数不一致。\n" +
                                     $"预期行数：{expectedRowCount}，实际读取行数：{dataRows.Count}\n" +
                                     $"是否继续处理？(Y/N): ";
                        
                        Console.Write(message);
                        var response = Console.ReadLine()?.Trim().ToUpper();
                        
                        if (response != "Y" && response != "YES")
                        {
                            throw new WorkPartnerException("UserCancelled", "用户取消了处理操作", filePath);
                        }
                        
                        Logger.Warning($"数据行数不一致：预期{expectedRowCount}行，实际{dataRows.Count}行，用户选择继续处理");
                    }
                }

                excelFile.DataRows = dataRows;
                Logger.CompleteFileProcessing(Path.GetFileName(filePath), "读取", fileInfo.Length, dataRows.Count);
                return excelFile;
            }, filePath);
        }

        public async Task<bool> SaveExcelFileAsync(ExcelFile excelFile, string outputPath)
        {
            return await Task.Run(() => SaveExcelFile(excelFile, outputPath));
        }

        public bool SaveExcelFile(ExcelFile excelFile, string outputPath)
        {
            return ExceptionHandler.HandleDataFormat(() =>
            {
                using var operation = Logger.StartOperation("保存Excel文件", Path.GetFileName(outputPath));
                Logger.StartFileProcessing(Path.GetFileName(outputPath), "保存");

                // 检查目标文件是否被锁定
                if (File.Exists(outputPath) && IsFileLocked(outputPath))
                {
                    // 等待一段时间后重试
                    Thread.Sleep(1000);
                    if (IsFileLocked(outputPath))
                    {
                        throw new InvalidOperationException($"目标文件被锁定，无法保存: {Path.GetFileName(outputPath)}");
                    }
                }

                var extension = Path.GetExtension(excelFile.FilePath).ToLower();
                
                bool result;
                if (extension == ".xls")
                {
                    result = SaveAsXlsFile(excelFile, outputPath);
                }
                else if (extension == ".xlsx")
                {
                    result = SaveAsXlsxFile(excelFile, outputPath);
                }
                else
                {
                    throw new WorkPartnerException("UnsupportedFormat", $"不支持的文件格式: {extension}", outputPath);
                }

                if (result)
                {
                    var fileInfo = new FileInfo(outputPath);
                    Logger.CompleteFileProcessing(Path.GetFileName(outputPath), "保存", fileInfo.Length);
                }
                
                return result;
            }, $"Excel文件保存 - {Path.GetFileName(outputPath)}", outputPath);
        }

        /// <summary>
        /// 保存Excel文件并更新A2列内容
        /// </summary>
        /// <param name="excelFile">Excel文件对象</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="currentObservationTime">本期观测时间</param>
        /// <param name="previousObservationTime">上期观测时间</param>
        /// <returns>是否保存成功</returns>
        public bool SaveExcelFileWithA2Update(ExcelFile excelFile, string outputPath, string currentObservationTime, string previousObservationTime)
        {
            return ExceptionHandler.HandleDataFormat(() =>
            {
                using var operation = Logger.StartOperation("保存Excel文件并更新A2列", Path.GetFileName(outputPath));
                Logger.StartFileProcessing(Path.GetFileName(outputPath), "保存");

                // 检查目标文件是否被锁定
                if (File.Exists(outputPath) && IsFileLocked(outputPath))
                {
                    // 等待一段时间后重试
                    Thread.Sleep(1000);
                    if (IsFileLocked(outputPath))
                    {
                        throw new InvalidOperationException($"目标文件被锁定，无法保存: {Path.GetFileName(outputPath)}");
                    }
                }

                var extension = Path.GetExtension(excelFile.FilePath).ToLower();
                
                bool result;
                if (extension == ".xls")
                {
                    result = SaveAsXlsFileWithA2Update(excelFile, outputPath, currentObservationTime, previousObservationTime);
                }
                else if (extension == ".xlsx")
                {
                    result = SaveAsXlsxFileWithA2Update(excelFile, outputPath, currentObservationTime, previousObservationTime);
                }
                else
                {
                    throw new WorkPartnerException("UnsupportedFormat", $"不支持的文件格式: {extension}", outputPath);
                }

                if (result)
                {
                    var fileInfo = new FileInfo(outputPath);
                    Logger.CompleteFileProcessing(Path.GetFileName(outputPath), "保存", fileInfo.Length);
                }
                
                return result;
            }, $"Excel文件保存并更新A2列 - {Path.GetFileName(outputPath)}", outputPath);
        }

        private bool SaveAsXlsFile(ExcelFile excelFile, string outputPath)
        {
            try
            {
                // 使用临时文件来避免文件锁定问题
                var tempPath = Path.Combine(Path.GetDirectoryName(outputPath), 
                    $"temp_{Guid.NewGuid()}_{Path.GetFileName(outputPath)}");
                
                try
                {
                    // 复制原文件到临时位置
                    File.Copy(excelFile.FilePath, tempPath, true);
                    
                    // 打开临时文件，更新数据
                    HSSFWorkbook workbook = null;
                    try
                    {
                        using (var fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }
                        
                        var sheet = workbook.GetSheetAt(0);

                        // 获取配置
                        var config = ExcelConfiguration.Instance;
                        
                        // 更新数据行
                        foreach (var dataRow in excelFile.DataRows)
                        {
                            var rowIndex = dataRow.RowIndex - 1; // 转换为0基索引
                            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

                            // 更新名称列
                            var nameCell = row.GetCell(config.NpoiNameCol) ?? row.CreateCell(config.NpoiNameCol);
                            nameCell.SetCellValue(dataRow.Name);

                            // 更新数据列
                            for (int j = 0; j < dataRow.Values.Count && j < config.TotalCols; j++)
                            {
                                var valueCell = row.GetCell(config.NpoiStartCol + j) ?? row.CreateCell(config.NpoiStartCol + j);
                                var value = dataRow.Values[j];
                                
                                if (value.HasValue)
                                {
                                    valueCell.SetCellValue(value.Value);
                                }
                                else
                                {
                                    valueCell.SetCellType(CellType.Blank);
                                }
                            }
                        }

                        // 保存到临时文件
                        using (var outputStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(outputStream);
                        }
                    }
                    finally
                    {
                        // 确保workbook被正确释放
                        if (workbook != null)
                        {
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }

                    // 等待一小段时间确保文件写入完成
                    Thread.Sleep(100);

                    // 如果目标文件存在，先删除它
                    if (File.Exists(outputPath))
                    {
                        try
                        {
                            File.Delete(outputPath);
                            Thread.Sleep(100); // 等待文件系统释放
                        }
                        catch (IOException)
                        {
                            // 如果删除失败，等待更长时间
                            Thread.Sleep(1000);
                            if (File.Exists(outputPath))
                            {
                                File.Delete(outputPath);
                            }
                        }
                    }

                    // 将临时文件移动到目标位置
                    File.Move(tempPath, outputPath);
                    
                    return true;
                }
                catch (Exception)
                {
                    // 如果出现任何错误，清理临时文件
                    if (File.Exists(tempPath))
                    {
                        try
                        {
                            File.Delete(tempPath);
                        }
                        catch
                        {
                            // 忽略清理错误
                        }
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLS文件失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        /// <summary>
        /// 保存XLS文件并更新A2列内容
        /// </summary>
        private bool SaveAsXlsFileWithA2Update(ExcelFile excelFile, string outputPath, string currentObservationTime, string previousObservationTime)
        {
            try
            {
                // 使用临时文件来避免文件锁定问题
                var tempPath = Path.Combine(Path.GetDirectoryName(outputPath), 
                    $"temp_{Guid.NewGuid()}_{Path.GetFileName(outputPath)}");
                
                try
                {
                    // 复制原文件到临时位置
                    File.Copy(excelFile.FilePath, tempPath, true);
                    
                    // 打开临时文件，更新数据
                    HSSFWorkbook workbook = null;
                    try
                    {
                        using (var fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }
                        
                        var sheet = workbook.GetSheetAt(0);

                        // 更新A2列内容
                        var a2Content = $"本期观测：{currentObservationTime} 上期观测：{previousObservationTime}";
                        var a2Row = sheet.GetRow(1) ?? sheet.CreateRow(1);
                        var a2Cell = a2Row.GetCell(0) ?? a2Row.CreateCell(0);
                        a2Cell.SetCellValue(a2Content);

                        // 获取配置
                        var config = ExcelConfiguration.Instance;
                        
                        // 更新数据行
                        foreach (var dataRow in excelFile.DataRows)
                        {
                            var rowIndex = dataRow.RowIndex - 1; // 转换为0基索引
                            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

                            // 更新名称列
                            var nameCell = row.GetCell(config.NpoiNameCol) ?? row.CreateCell(config.NpoiNameCol);
                            nameCell.SetCellValue(dataRow.Name);

                            // 更新数据列
                            for (int j = 0; j < dataRow.Values.Count && j < config.TotalCols; j++)
                            {
                                var valueCell = row.GetCell(config.NpoiStartCol + j) ?? row.CreateCell(config.NpoiStartCol + j);
                                var value = dataRow.Values[j];
                                
                                if (value.HasValue)
                                {
                                    valueCell.SetCellValue(value.Value);
                                }
                                else
                                {
                                    valueCell.SetCellType(CellType.Blank);
                                }
                            }
                        }

                        // 保存到临时文件
                        using (var outputStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(outputStream);
                        }
                    }
                    finally
                    {
                        // 确保workbook被正确释放
                        if (workbook != null)
                        {
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }

                    // 等待一小段时间确保文件写入完成
                    Thread.Sleep(100);

                    // 如果目标文件存在，先删除它
                    if (File.Exists(outputPath))
                    {
                        try
                        {
                            File.Delete(outputPath);
                            Thread.Sleep(100); // 等待文件系统释放
                        }
                        catch (IOException)
                        {
                            // 如果删除失败，等待更长时间
                            Thread.Sleep(1000);
                            if (File.Exists(outputPath))
                            {
                                File.Delete(outputPath);
                            }
                        }
                    }

                    // 将临时文件移动到目标位置
                    File.Move(tempPath, outputPath);
                    
                    return true;
                }
                catch (Exception)
                {
                    // 如果出现任何错误，清理临时文件
                    if (File.Exists(tempPath))
                    {
                        try
                        {
                            File.Delete(tempPath);
                        }
                        catch
                        {
                            // 忽略清理错误
                        }
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLS文件并更新A2列失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        private bool SaveAsXlsxFile(ExcelFile excelFile, string outputPath)
        {
            try
            {
                // 使用临时文件来避免文件锁定问题
                var tempPath = Path.Combine(Path.GetDirectoryName(outputPath), 
                    $"temp_{Guid.NewGuid()}_{Path.GetFileName(outputPath)}");
                
                try
                {
                    // 复制原文件到临时位置
                    File.Copy(excelFile.FilePath, tempPath, true);
                    
                    // 打开临时文件，更新数据
                    using var package = new ExcelPackage(new FileInfo(tempPath));
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    
                    if (worksheet == null)
                    {
                        throw new InvalidOperationException("Excel文件中没有找到工作表");
                    }

                    // 获取配置
                    var config = ExcelConfiguration.Instance;
                    
                    // 更新数据行
                    foreach (var dataRow in excelFile.DataRows)
                    {
                        var rowIndex = dataRow.RowIndex; // EPPlus使用1基索引

                        // 更新名称列
                        worksheet.Cells[rowIndex, config.NameCol].Value = dataRow.Name;

                        // 更新数据列
                        for (int j = 0; j < dataRow.Values.Count && j < config.TotalCols; j++)
                        {
                            var value = dataRow.Values[j];
                            
                            if (value.HasValue)
                            {
                                worksheet.Cells[rowIndex, config.StartCol + j].Value = value.Value;
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, config.StartCol + j].Value = null;
                            }
                        }
                    }

                    package.Save();
                    
                    // 等待一小段时间确保文件写入完成
                    Thread.Sleep(100);

                    // 如果目标文件存在，先删除它
                    if (File.Exists(outputPath))
                    {
                        try
                        {
                            File.Delete(outputPath);
                            Thread.Sleep(100); // 等待文件系统释放
                        }
                        catch (IOException)
                        {
                            // 如果删除失败，等待更长时间
                            Thread.Sleep(1000);
                            if (File.Exists(outputPath))
                            {
                                File.Delete(outputPath);
                            }
                        }
                    }

                    // 将临时文件移动到目标位置
                    File.Move(tempPath, outputPath);
                    
                    return true;
                }
                catch (Exception)
                {
                    // 如果出现任何错误，清理临时文件
                    if (File.Exists(tempPath))
                    {
                        try
                        {
                            File.Delete(tempPath);
                        }
                        catch
                        {
                            // 忽略清理错误
                        }
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLSX文件失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        /// <summary>
        /// 保存XLSX文件并更新A2列内容
        /// </summary>
        private bool SaveAsXlsxFileWithA2Update(ExcelFile excelFile, string outputPath, string currentObservationTime, string previousObservationTime)
        {
            try
            {
                // 使用临时文件来避免文件锁定问题
                var tempPath = Path.Combine(Path.GetDirectoryName(outputPath), 
                    $"temp_{Guid.NewGuid()}_{Path.GetFileName(outputPath)}");
                
                try
                {
                    // 复制原文件到临时位置
                    File.Copy(excelFile.FilePath, tempPath, true);
                    
                    // 打开临时文件，更新数据
                    using var package = new ExcelPackage(new FileInfo(tempPath));
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    
                    if (worksheet == null)
                    {
                        throw new InvalidOperationException("Excel文件中没有找到工作表");
                    }

                    // 更新A2列内容
                    var a2Content = $"本期观测：{currentObservationTime} 上期观测：{previousObservationTime}";
                    worksheet.Cells["A2"].Value = a2Content;

                    // 获取配置
                    var config = ExcelConfiguration.Instance;
                    
                    // 更新数据行
                    foreach (var dataRow in excelFile.DataRows)
                    {
                        var rowIndex = dataRow.RowIndex; // EPPlus使用1基索引

                        // 更新名称列
                        worksheet.Cells[rowIndex, config.NameCol].Value = dataRow.Name;

                        // 更新数据列
                        for (int j = 0; j < dataRow.Values.Count && j < config.TotalCols; j++)
                        {
                            var value = dataRow.Values[j];
                            
                            if (value.HasValue)
                            {
                                worksheet.Cells[rowIndex, config.StartCol + j].Value = value.Value;
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, config.StartCol + j].Value = null;
                            }
                        }
                    }

                    package.Save();
                    
                    // 等待一小段时间确保文件写入完成
                    Thread.Sleep(100);

                    // 如果目标文件存在，先删除它
                    if (File.Exists(outputPath))
                    {
                        try
                        {
                            File.Delete(outputPath);
                            Thread.Sleep(100); // 等待文件系统释放
                        }
                        catch (IOException)
                        {
                            // 如果删除失败，等待更长时间
                            Thread.Sleep(1000);
                            if (File.Exists(outputPath))
                            {
                                File.Delete(outputPath);
                            }
                        }
                    }

                    // 将临时文件移动到目标位置
                    File.Move(tempPath, outputPath);
                    
                    return true;
                }
                catch (Exception)
                {
                    // 如果出现任何错误，清理临时文件
                    if (File.Exists(tempPath))
                    {
                        try
                        {
                            File.Delete(tempPath);
                        }
                        catch
                        {
                            // 忽略清理错误
                        }
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLSX文件并更新A2列失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        public bool ValidateExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            var extension = Path.GetExtension(filePath).ToLower();
            if (extension != ".xlsx" && extension != ".xls")
                return false;

            try
            {
                if (extension == ".xlsx")
                {
                    using var package = new ExcelPackage(new FileInfo(filePath));
                    return package.Workbook.Worksheets.Count > 0;
                }
                else if (extension == ".xls")
                {
                    using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    var workbook = new HSSFWorkbook(fileStream);
                    return workbook.NumberOfSheets > 0;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public ExcelFile GetExcelFileInfo(string filePath)
        {
            var fileInfo = new FileInfo(filePath);
            
            return new ExcelFile
            {
                FilePath = filePath,
                FileName = fileInfo.Name,
                FileSize = fileInfo.Length,
                LastModified = fileInfo.LastWriteTime,
                IsLocked = IsFileLocked(filePath),
                IsValid = ValidateExcelFile(filePath)
            };
        }

        public bool CopyExcelFile(string sourcePath, string targetPath)
        {
            try
            {
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetDirectory) && !Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                File.Copy(sourcePath, targetPath, true);
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"复制Excel文件失败: {sourcePath} -> {targetPath}", ex);
            }
        }

        public bool IsFileLocked(string filePath)
        {
            try
            {
                using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                return false; // 文件未被锁定
            }
            catch (IOException)
            {
                return true; // 文件被锁定
            }
            catch (Exception)
            {
                return true; // 其他错误也认为文件不可访问
            }
        }
    }
} 