using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using System.Text;

namespace DataExport.Services
{
    public class DataExportService
    {
        private readonly ILogger<DataExportService> _logger;
        private readonly HttpClient _httpClient;
        private readonly ExportConfig _config;

        public DataExportService(ILogger<DataExportService> logger, HttpClient httpClient, ExportConfig config)
        {
            _logger = logger;
            _httpClient = httpClient;
            _config = config;
        }

        /// <summary>
        /// 导出数据
        /// </summary>
        public async Task<ExportResult> ExportDataAsync(ExportParameters parameters)
        {
            _logger.LogInformation("开始导出数据: {ProjectName} - {DataName} ({StartTime} 至 {EndTime})", 
                parameters.ProjectName, parameters.DataName, parameters.StartTime, parameters.EndTime);

            try
            {
                // 构建请求URL
                var url = $"{_config.ApiSettings.BaseUrl}{_config.ApiSettings.Endpoint}";
                
                // 构建查询参数
                var queryParams = new Dictionary<string, string>
                {
                    ["projectId"] = parameters.ProjectId,
                    ["projectCode"] = "", // 暂时留空
                    ["ProjectName"] = parameters.ProjectName,
                    ["DataCode"] = parameters.DataCode,
                    ["DataName"] = parameters.DataName,
                    ["StartTime"] = parameters.StartTime,
                    ["EndTime"] = parameters.EndTime,
                    ["PointCodes"] = parameters.PointCodes,
                    ["WithDetail"] = parameters.WithDetail.ToString()
                };

                var queryString = string.Join("&", queryParams.Select(kvp => $"{kvp.Key}={Uri.EscapeDataString(kvp.Value)}"));
                var fullUrl = $"{url}?{queryString}";

                _logger.LogDebug("请求URL: {Url}", fullUrl);

                // 创建HTTP请求
                var request = new HttpRequestMessage(HttpMethod.Get, fullUrl);
                
                // 添加请求头
                request.Headers.Add("User-Agent", _config.ApiSettings.UserAgent);
                request.Headers.Add("Referer", _config.ApiSettings.Referer);
                request.Headers.Add("Cookie", _config.ApiSettings.Cookie);

                // 发送请求
                var response = await _httpClient.SendAsync(request);
                
                if (!response.IsSuccessStatusCode)
                {
                    var errorMessage = $"HTTP请求失败: {response.StatusCode} - {response.ReasonPhrase}";
                    _logger.LogError(errorMessage);
                    return new ExportResult
                    {
                        Success = false,
                        ErrorMessage = errorMessage
                    };
                }

                // 检查响应内容类型
                var contentType = response.Content.Headers.ContentType?.ToString() ?? "";
                _logger.LogDebug("响应内容类型: {ContentType}", contentType);

                // 读取响应内容
                var content = await response.Content.ReadAsByteArrayAsync();
                if (content.Length == 0)
                {
                    var errorMessage = "响应内容为空";
                    _logger.LogError(errorMessage);
                    return new ExportResult
                    {
                        Success = false,
                        ErrorMessage = errorMessage
                    };
                }

                _logger.LogInformation("接收到响应数据，大小: {Size} 字节", content.Length);

                // 检测文件格式
                var actualFormat = DetectFileFormat(content);
                _logger.LogInformation("检测到的实际文件格式: {Format}", actualFormat);

                // 生成文件名
                var fileName = GenerateFileName(parameters, actualFormat);
                var filePath = Path.Combine(_config.ExportSettings.OutputDirectory, fileName);

                // 确保输出目录存在
                Directory.CreateDirectory(_config.ExportSettings.OutputDirectory);

                // 保存文件
                await File.WriteAllBytesAsync(filePath, content);
                _logger.LogInformation("文件已保存: {FilePath}", filePath);

                // 验证文件格式
                var formatValidation = ValidateFileFormat(filePath, actualFormat);
                if (!formatValidation.IsValid)
                {
                    _logger.LogWarning("文件格式验证失败: {Message}", formatValidation.Message);
                }

                // 如果扩展名与格式不匹配，自动重命名
                var finalFilePath = AutoRenameFileIfNeeded(filePath, actualFormat);

                return new ExportResult
                {
                    Success = true,
                    ProjectName = parameters.ProjectName,
                    DataName = parameters.DataName,
                    StartTime = parameters.StartTime,
                    EndTime = parameters.EndTime,
                    FileName = Path.GetFileName(finalFilePath),
                    FilePath = finalFilePath
                };
            }
            catch (Exception ex)
            {
                var errorMessage = $"导出过程中发生异常: {ex.Message}";
                _logger.LogError(ex, errorMessage);
                return new ExportResult
                {
                    Success = false,
                    ErrorMessage = errorMessage
                };
            }
        }

        /// <summary>
        /// 检测文件格式
        /// </summary>
        private string DetectFileFormat(byte[] content)
        {
            if (content.Length < 4) return "未知";

            // 检查文件头
            if (content[0] == 0xD0 && content[1] == 0xCF && content[2] == 0x11 && content[3] == 0xE0)
            {
                return ".xls"; // Microsoft Office 97-2003
            }
            else if (content[0] == 0x50 && content[1] == 0x4B && content[2] == 0x03 && content[3] == 0x04)
            {
                return ".xlsx"; // Excel 2007+
            }

            return "未知";
        }

        /// <summary>
        /// 生成文件名
        /// </summary>
        private string GenerateFileName(ExportParameters parameters, string format)
        {
            var startDate = DateTime.Parse(parameters.StartTime).ToString("yyyyMMdd");
            var endDate = DateTime.Parse(parameters.EndTime).ToString("yyyyMMdd");
            
            return $"{parameters.ProjectName}_{parameters.DataName}_{startDate}_{endDate}{format}";
        }

        /// <summary>
        /// 验证文件格式
        /// </summary>
        private FormatValidationResult ValidateFileFormat(string filePath, string expectedFormat)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == expectedFormat)
                {
                    return new FormatValidationResult { IsValid = true };
                }

                return new FormatValidationResult 
                { 
                    IsValid = false, 
                    Message = $"文件扩展名({extension})与格式({expectedFormat})不匹配" 
                };
            }
            catch (Exception ex)
            {
                return new FormatValidationResult 
                { 
                    IsValid = false, 
                    Message = $"格式验证异常: {ex.Message}" 
                };
            }
        }

        /// <summary>
        /// 自动重命名文件（如果需要）
        /// </summary>
        private string AutoRenameFileIfNeeded(string filePath, string actualFormat)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == actualFormat)
                {
                    return filePath; // 格式匹配，无需重命名
                }

                var newFileName = Path.ChangeExtension(Path.GetFileName(filePath), actualFormat);
                var newFilePath = Path.Combine(Path.GetDirectoryName(filePath)!, newFileName);

                if (File.Exists(newFilePath))
                {
                    _logger.LogWarning("目标文件已存在，无法重命名: {FileName}", newFileName);
                    return filePath;
                }

                File.Move(filePath, newFilePath);
                _logger.LogInformation("文件已自动重命名为: {FileName}", newFileName);
                
                return newFilePath;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("自动重命名失败: {Message}", ex.Message);
                return filePath;
            }
        }

        /// <summary>
        /// 测试连接
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                var url = $"{_config.ApiSettings.BaseUrl}{_config.ApiSettings.Endpoint}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                
                request.Headers.Add("User-Agent", _config.ApiSettings.UserAgent);
                request.Headers.Add("Referer", _config.ApiSettings.Referer);
                request.Headers.Add("Cookie", _config.ApiSettings.Cookie);

                var response = await _httpClient.SendAsync(request);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "连接测试失败");
                return false;
            }
        }
    }

    /// <summary>
    /// 导出结果
    /// </summary>
    public class ExportResult
    {
        public bool Success { get; set; }
        public string ProjectName { get; set; } = string.Empty;
        public string DataName { get; set; } = string.Empty;
        public string MonthName { get; set; } = string.Empty;
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public string? FileName { get; set; }
        public string? FilePath { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// 格式验证结果
    /// </summary>
    public class FormatValidationResult
    {
        public bool IsValid { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
