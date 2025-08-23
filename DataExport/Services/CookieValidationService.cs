using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// Cookie验证服务
    /// </summary>
    public class CookieValidationService
    {
        private readonly ILogger<CookieValidationService> _logger;
        private readonly HttpClient _httpClient;
        private readonly ExportConfig _config;

        public CookieValidationService(ILogger<CookieValidationService> logger, HttpClient httpClient, ExportConfig config)
        {
            _logger = logger;
            _httpClient = httpClient;
            _config = config;
        }

        /// <summary>
        /// 验证Cookie是否有效
        /// </summary>
        /// <returns>验证结果</returns>
        public async Task<CookieValidationResult> ValidateCookieAsync()
        {
            _logger.LogInformation("开始验证Cookie有效性...");
            
            try
            {
                // 检查Cookie配置
                if (string.IsNullOrWhiteSpace(_config.ApiSettings.Cookie))
                {
                    return new CookieValidationResult
                    {
                        IsValid = false,
                        Message = "Cookie配置为空",
                        Details = "请在appsettings.json中配置有效的Cookie值"
                    };
                }

                // 检查Cookie格式
                var cookieValidation = ValidateCookieFormat(_config.ApiSettings.Cookie);
                if (!cookieValidation.IsValid)
                {
                    return cookieValidation;
                }

                // 测试API连接
                var apiTestResult = await TestApiConnectionAsync();
                if (!apiTestResult.IsValid)
                {
                    return apiTestResult;
                }

                // 测试天气接口（用于验证Cookie有效性）
                var weatherTestResult = await TestWeatherApiAsync();
                if (!weatherTestResult.IsValid)
                {
                    return weatherTestResult;
                }

                return new CookieValidationResult
                {
                    IsValid = true,
                    Message = "Cookie验证成功",
                    Details = "所有测试接口均能正常访问，Cookie有效",
                    CookieInfo = ExtractCookieInfo(_config.ApiSettings.Cookie)
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Cookie验证过程中发生异常");
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "Cookie验证失败",
                    Details = $"验证过程中发生异常: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// 验证Cookie格式
        /// </summary>
        private CookieValidationResult ValidateCookieFormat(string cookie)
        {
            var requiredKeys = new[]
            {
                "ASP.NET_SessionId",
                "Qianchen_ADMS_V7_Token",
                "Qianchen_ADMS_V7_Mark"
            };

            var missingKeys = new List<string>();
            foreach (var key in requiredKeys)
            {
                if (!cookie.Contains($"{key}="))
                {
                    missingKeys.Add(key);
                }
            }

            if (missingKeys.Any())
            {
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "Cookie格式不完整",
                    Details = $"缺少必要的Cookie键: {string.Join(", ", missingKeys)}"
                };
            }

            return new CookieValidationResult
            {
                IsValid = true,
                Message = "Cookie格式正确"
            };
        }

        /// <summary>
        /// 测试API连接
        /// </summary>
        private async Task<CookieValidationResult> TestApiConnectionAsync()
        {
            try
            {
                var url = $"{_config.ApiSettings.BaseUrl}/QC_FoundationPit/ResultsQuery/DataList";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                
                // 添加必要的请求头
                request.Headers.Add("User-Agent", _config.ApiSettings.UserAgent);
                request.Headers.Add("Referer", _config.ApiSettings.Referer);
                request.Headers.Add("Cookie", _config.ApiSettings.Cookie);
                request.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8");
                request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6");

                var response = await _httpClient.SendAsync(request);
                
                if (response.IsSuccessStatusCode)
                {
                    return new CookieValidationResult
                    {
                        IsValid = true,
                        Message = "API连接测试成功"
                    };
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    return new CookieValidationResult
                    {
                        IsValid = false,
                        Message = "API认证失败",
                        Details = $"HTTP状态码: {response.StatusCode} - Cookie可能已过期或无效"
                    };
                }
                else
                {
                    return new CookieValidationResult
                    {
                        IsValid = false,
                        Message = "API连接异常",
                        Details = $"HTTP状态码: {response.StatusCode} - {response.ReasonPhrase}"
                    };
                }
            }
            catch (HttpRequestException ex)
            {
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "API连接失败",
                    Details = $"网络连接异常: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "API测试异常",
                    Details = $"测试过程中发生异常: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// 测试天气API接口（用于验证Cookie有效性）
        /// </summary>
        private async Task<CookieValidationResult> TestWeatherApiAsync()
        {
            try
            {
                var url = $"{_config.ApiSettings.BaseUrl}/QC_FoundationPit/HomePage/GetWeather?_={DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                
                // 添加必要的请求头
                request.Headers.Add("User-Agent", _config.ApiSettings.UserAgent);
                request.Headers.Add("Accept", "application/json, text/javascript, */*; q=0.01");
                request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6");
                request.Headers.Add("Connection", "keep-alive");
                request.Headers.Add("Cookie", _config.ApiSettings.Cookie);
                request.Headers.Add("Referer", $"{_config.ApiSettings.BaseUrl}/Home/Index");
                request.Headers.Add("Sec-Fetch-Dest", "empty");
                request.Headers.Add("Sec-Fetch-Mode", "cors");
                request.Headers.Add("Sec-Fetch-Site", "same-origin");
                request.Headers.Add("X-Requested-With", "XMLHttpRequest");
                request.Headers.Add("account", "System");

                var response = await _httpClient.SendAsync(request);
                
                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    _logger.LogDebug("天气API响应: {Content}", content);
                    
                    // 尝试解析JSON响应
                    try
                    {
                        var jsonDoc = JsonDocument.Parse(content);
                        return new CookieValidationResult
                        {
                            IsValid = true,
                            Message = "天气API测试成功",
                            Details = "成功获取天气数据，Cookie有效"
                        };
                    }
                    catch (JsonException)
                    {
                        // 如果不是JSON格式，检查是否包含错误信息
                        if (content.Contains("error") || content.Contains("Error") || content.Contains("未授权"))
                        {
                            return new CookieValidationResult
                            {
                                IsValid = false,
                                Message = "天气API返回错误",
                                Details = $"响应内容: {content}"
                            };
                        }
                        
                        return new CookieValidationResult
                        {
                            IsValid = true,
                            Message = "天气API测试成功",
                            Details = "成功获取响应数据，Cookie有效"
                        };
                    }
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    return new CookieValidationResult
                    {
                        IsValid = false,
                        Message = "天气API认证失败",
                        Details = $"HTTP状态码: {response.StatusCode} - Cookie可能已过期或无效"
                    };
                }
                else
                {
                    var content = await response.Content.ReadAsStringAsync();
                    return new CookieValidationResult
                    {
                        IsValid = false,
                        Message = "天气API响应异常",
                        Details = $"HTTP状态码: {response.StatusCode} - 响应内容: {content}"
                    };
                }
            }
            catch (HttpRequestException ex)
            {
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "天气API连接失败",
                    Details = $"网络连接异常: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                return new CookieValidationResult
                {
                    IsValid = false,
                    Message = "天气API测试异常",
                    Details = $"测试过程中发生异常: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// 提取Cookie信息
        /// </summary>
        private Dictionary<string, string> ExtractCookieInfo(string cookie)
        {
            var cookieInfo = new Dictionary<string, string>();
            var pairs = cookie.Split(';', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var pair in pairs)
            {
                var trimmed = pair.Trim();
                var separatorIndex = trimmed.IndexOf('=');
                if (separatorIndex > 0)
                {
                    var key = trimmed.Substring(0, separatorIndex).Trim();
                    var value = trimmed.Substring(separatorIndex + 1).Trim();
                    
                    // 隐藏敏感信息
                    if (key.Contains("Token") || key.Contains("SessionId"))
                    {
                        if (value.Length > 8)
                        {
                            value = value.Substring(0, 4) + "..." + value.Substring(value.Length - 4);
                        }
                    }
                    
                    cookieInfo[key] = value;
                }
            }
            
            return cookieInfo;
        }

        /// <summary>
        /// 显示Cookie验证结果
        /// </summary>
        public void DisplayValidationResult(CookieValidationResult result)
        {
            Console.WriteLine("\n=== Cookie验证结果 ===");
            Console.WriteLine($"状态: {(result.IsValid ? "✓ 有效" : "✗ 无效")}");
            Console.WriteLine($"消息: {result.Message}");
            
            if (!string.IsNullOrEmpty(result.Details))
            {
                Console.WriteLine($"详情: {result.Details}");
            }
            
            if (result.CookieInfo != null && result.CookieInfo.Any())
            {
                Console.WriteLine("\nCookie信息:");
                foreach (var kvp in result.CookieInfo)
                {
                    Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                }
            }
            
            Console.WriteLine("=".PadRight(50, '='));
        }

        /// <summary>
        /// 更新Cookie配置
        /// </summary>
        /// <param name="newCookie">新的Cookie字符串</param>
        /// <returns>更新结果</returns>
        public async Task<bool> UpdateCookieAsync(string newCookie)
        {
            try
            {
                _logger.LogInformation("开始更新Cookie配置...");
                
                // 验证新Cookie格式
                var formatValidation = ValidateCookieFormat(newCookie);
                if (!formatValidation.IsValid)
                {
                    _logger.LogWarning("新Cookie格式验证失败: {Message}", formatValidation.Message);
                    return false;
                }
                
                // 更新配置
                _config.ApiSettings.Cookie = newCookie;
                
                // 保存配置到文件
                await SaveConfigurationAsync();
                
                _logger.LogInformation("Cookie配置更新成功");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "更新Cookie配置失败");
                return false;
            }
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        private async Task SaveConfigurationAsync()
        {
            try
            {
                var configPath = "appsettings.json";
                var jsonString = System.Text.Json.JsonSerializer.Serialize(_config, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });
                
                await File.WriteAllTextAsync(configPath, jsonString);
                _logger.LogInformation("配置已保存到: {ConfigPath}", configPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存配置失败");
                throw;
            }
        }

        /// <summary>
        /// 从Cookie字符串中提取必要的Cookie值
        /// </summary>
        /// <param name="cookieString">完整的Cookie字符串</param>
        /// <returns>提取的Cookie值</returns>
        public string ExtractRequiredCookies(string cookieString)
        {
            var requiredKeys = new[]
            {
                "ASP.NET_SessionId",
                "Qianchen_ADMS_V7_Token",
                "Qianchen_ADMS_V7_Mark"
            };

            var extractedCookies = new List<string>();
            var pairs = cookieString.Split(';', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var pair in pairs)
            {
                var trimmed = pair.Trim();
                var separatorIndex = trimmed.IndexOf('=');
                if (separatorIndex > 0)
                {
                    var key = trimmed.Substring(0, separatorIndex).Trim();
                    if (requiredKeys.Contains(key))
                    {
                        extractedCookies.Add(trimmed);
                    }
                }
            }
            
            return string.Join("; ", extractedCookies);
        }
    }

    /// <summary>
    /// Cookie验证结果
    /// </summary>
    public class CookieValidationResult
    {
        public bool IsValid { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Details { get; set; } = string.Empty;
        public Dictionary<string, string>? CookieInfo { get; set; }
    }
}
