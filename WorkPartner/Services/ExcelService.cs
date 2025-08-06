using WorkPartner.Models;
using OfficeOpenXml;
using NPOI.HSSF.UserModel; // for .xls files
using NPOI.SS.UserModel;
using System.IO;

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
            try
            {
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

                if (!excelFile.IsValid)
                {
                    throw new InvalidOperationException($"Excel文件格式无效: {Path.GetFileName(filePath)}");
                }

                if (excelFile.IsLocked)
                {
                    throw new InvalidOperationException($"Excel文件被占用: {Path.GetFileName(filePath)}");
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

                    // 读取B5-B368行的数据名称
                    for (int row = 5; row <= 368; row++)
                    {
                        var nameCell = worksheet.Cells[row, 2]; // B列
                        var name = nameCell?.Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(name))
                        {
                            var dataRow = new DataRow
                            {
                                Name = name,
                                RowIndex = row
                            };

                            // 读取D5-I5列的数据值
                            for (int col = 4; col <= 9; col++) // D-I列
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

                    // 读取B5-B368行的数据名称
                    for (int row = 4; row <= 367; row++) // NPOI使用0基索引，所以B5对应row=4, col=1
                    {
                        var nameRow = sheet.GetRow(row);
                        if (nameRow == null) continue;

                        var nameCell = nameRow.GetCell(1); // B列（索引为1）
                        var name = nameCell?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(name))
                        {
                            var dataRow = new DataRow
                            {
                                Name = name,
                                RowIndex = row + 1 // 转换为1基索引显示
                            };

                            // 读取D5-I5列的数据值
                            for (int col = 3; col <= 8; col++) // D-I列（索引3-8）
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

                excelFile.DataRows = dataRows;
                return excelFile;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"读取Excel文件失败: {Path.GetFileName(filePath)}", ex);
            }
        }

        public async Task<bool> SaveExcelFileAsync(ExcelFile excelFile, string outputPath)
        {
            return await Task.Run(() => SaveExcelFile(excelFile, outputPath));
        }

        public bool SaveExcelFile(ExcelFile excelFile, string outputPath)
        {
            try
            {
                var extension = Path.GetExtension(excelFile.FilePath).ToLower();
                
                if (extension == ".xls")
                {
                    return SaveAsXlsFile(excelFile, outputPath);
                }
                else if (extension == ".xlsx")
                {
                    return SaveAsXlsxFile(excelFile, outputPath);
                }
                else
                {
                    throw new InvalidOperationException($"不支持的文件格式: {extension}");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存Excel文件失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        private bool SaveAsXlsFile(ExcelFile excelFile, string outputPath)
        {
            try
            {
                // 复制原文件保持格式
                File.Copy(excelFile.FilePath, outputPath, true);
                
                // 打开复制的文件，更新数据
                HSSFWorkbook workbook;
                using (var fileStream = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                
                var sheet = workbook.GetSheetAt(0);

                // 更新数据行
                foreach (var dataRow in excelFile.DataRows)
                {
                    var rowIndex = dataRow.RowIndex - 1; // 转换为0基索引
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

                    // 更新B列的名称
                    var nameCell = row.GetCell(1) ?? row.CreateCell(1);
                    nameCell.SetCellValue(dataRow.Name);

                    // 更新D-I列的数据值
                    for (int j = 0; j < dataRow.Values.Count && j < 6; j++)
                    {
                        var valueCell = row.GetCell(j + 3) ?? row.CreateCell(j + 3);
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

                // 保存回文件
                using (var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outputStream);
                }
                
                workbook.Close();
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLS文件失败: {Path.GetFileName(outputPath)}", ex);
            }
        }

        private bool SaveAsXlsxFile(ExcelFile excelFile, string outputPath)
        {
            try
            {
                // 复制原文件保持格式
                File.Copy(excelFile.FilePath, outputPath, true);
                
                // 打开复制的文件，更新数据
                using var package = new ExcelPackage(new FileInfo(outputPath));
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                {
                    throw new InvalidOperationException("Excel文件中没有找到工作表");
                }

                // 更新数据行
                foreach (var dataRow in excelFile.DataRows)
                {
                    var rowIndex = dataRow.RowIndex; // EPPlus使用1基索引

                    // 更新B列的名称
                    worksheet.Cells[rowIndex, 2].Value = dataRow.Name;

                    // 更新D-I列的数据值
                    for (int j = 0; j < dataRow.Values.Count && j < 6; j++)
                    {
                        var value = dataRow.Values[j];
                        
                        if (value.HasValue)
                        {
                            worksheet.Cells[rowIndex, j + 4].Value = value.Value;
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, j + 4].Value = null;
                        }
                    }
                }

                package.Save();
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"保存XLSX文件失败: {Path.GetFileName(outputPath)}", ex);
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