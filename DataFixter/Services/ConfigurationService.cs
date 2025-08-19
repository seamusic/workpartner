using System;
using System.IO;
using DataFixter.Models;
using Microsoft.Extensions.Configuration;
using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 配置服务，负责读取和管理应用程序配置
    /// </summary>
    public class ConfigurationService
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger _logger;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public ConfigurationService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            // 构建配置
            _configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();
                
            _logger.Information("配置服务已初始化");
        }

        /// <summary>
        /// 获取验证选项
        /// </summary>
        /// <returns>验证选项</returns>
        public ValidationOptions GetValidationOptions()
        {
            var options = new ValidationOptions();
            
            try
            {
                var section = _configuration.GetSection("ValidationOptions");
                
                if (section.Exists())
                {
                    options.CumulativeTolerance = section.GetValue<double>("CumulativeTolerance", 0.01);
                    options.CriticalThreshold = section.GetValue<double>("CriticalThreshold", 1.0);
                    options.ErrorThreshold = section.GetValue<double>("ErrorThreshold", 0.5);
                    options.MinValueThreshold = section.GetValue<double>("MinValueThreshold", 0.01);
                    options.MaxCurrentPeriodValue = section.GetValue<double>("MaxCurrentPeriodValue", 1.0);
                    options.MileageTolerance = section.GetValue<double>("MileageTolerance", 0.01);
                    options.MaxTimeInterval = section.GetValue<double>("MaxTimeInterval", 30.0);
                    
                    // 新增的超时和批处理配置
                    options.MaxProcessingTimeMinutes = section.GetValue<int>("MaxProcessingTimeMinutes", 30);
                    options.BatchSize = section.GetValue<int>("BatchSize", 50);
                    options.EnableMemoryCleanup = section.GetValue<bool>("EnableMemoryCleanup", true);
                    options.MemoryCleanupFrequency = section.GetValue<int>("MemoryCleanupFrequency", 200);
                    options.MaxDegreeOfParallelism = section.GetValue<int>("MaxDegreeOfParallelism", 0);
                    
                    _logger.Information("从配置文件加载验证选项: 累计变化量容差={CumulativeTolerance}, 最大处理时间={MaxProcessingTimeMinutes}分钟, 批处理大小={BatchSize}, 并行度={MaxDegreeOfParallelism}", 
                        options.CumulativeTolerance, options.MaxProcessingTimeMinutes, options.BatchSize, options.MaxDegreeOfParallelism);
                }
                else
                {
                    _logger.Warning("配置文件中未找到ValidationOptions节点，使用默认值");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "读取验证选项配置时发生异常，使用默认值");
            }

            return options;
        }

        /// <summary>
        /// 获取修正选项
        /// </summary>
        /// <returns>修正选项</returns>
        public CorrectionOptions GetCorrectionOptions()
        {
            var options = new CorrectionOptions();
            
            try
            {
                var section = _configuration.GetSection("CorrectionOptions");
                
                if (section.Exists())
                {
                    options.CumulativeTolerance = section.GetValue<double>("CumulativeTolerance", 0.01);
                    options.MaxCurrentPeriodValue = section.GetValue<double>("MaxCurrentPeriodValue", 1.0);
                    options.MaxCumulativeValue = section.GetValue<double>("MaxCumulativeValue", 4.0);
                    options.EnableMinimalModification = section.GetValue<bool>("EnableMinimalModification", true);
                    options.RandomChangeRange = section.GetValue<double>("RandomChangeRange", 0.3);
                    
                    _logger.Information("从配置文件加载修正选项: 累计变化量容差={CumulativeTolerance}, 随机变化量范围={RandomChangeRange}", 
                        options.CumulativeTolerance, options.RandomChangeRange);
                }
                else
                {
                    _logger.Warning("配置文件中未找到CorrectionOptions节点，使用默认值");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "读取修正选项配置时发生异常，使用默认值");
            }

            return options;
        }

        /// <summary>
        /// 获取输出选项
        /// </summary>
        /// <returns>输出选项</returns>
        public OutputOptions GetOutputOptions()
        {
            var options = new OutputOptions();
            
            try
            {
                var section = _configuration.GetSection("OutputOptions");
                
                if (section.Exists())
                {
                    options.PreserveOriginalFormat = section.GetValue<bool>("PreserveOriginalFormat", true);
                    options.AddCorrectionMarks = section.GetValue<bool>("AddCorrectionMarks", true);
                    options.OutputEncoding = section.GetValue<string>("OutputEncoding", "UTF-8");
                    
                    _logger.Information("从配置文件加载输出选项: 保留原始格式={PreserveOriginalFormat}", options.PreserveOriginalFormat);
                }
                else
                {
                    _logger.Warning("配置文件中未找到OutputOptions节点，使用默认值");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "读取输出选项配置时发生异常，使用默认值");
            }

            return options;
        }

        /// <summary>
        /// 获取配置值
        /// </summary>
        /// <typeparam name="T">值类型</typeparam>
        /// <param name="key">配置键</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>配置值</returns>
        public T GetValue<T>(string key, T defaultValue = default(T))
        {
            try
            {
                return _configuration.GetValue<T>(key, defaultValue);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "读取配置值 {Key} 时发生异常，使用默认值 {DefaultValue}", key, defaultValue);
                return defaultValue;
            }
        }

        /// <summary>
        /// 获取连接字符串
        /// </summary>
        /// <param name="name">连接字符串名称</param>
        /// <returns>连接字符串</returns>
        public string GetConnectionString(string name)
        {
            try
            {
                return _configuration.GetConnectionString(name) ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "读取连接字符串 {Name} 时发生异常", name);
                return string.Empty;
            }
        }
    }
}
