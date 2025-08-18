using Serilog;
using Serilog.Events;
using Microsoft.Extensions.Configuration;

namespace DataFixter.Logging
{
    /// <summary>
    /// 日志配置类，负责配置和初始化Serilog日志记录器
    /// </summary>
    public static class LoggingConfiguration
    {
        /// <summary>
        /// 配置并初始化日志记录器
        /// </summary>
        /// <param name="logFilePath">日志文件路径，默认为logs目录下的DataFixter.log</param>
        public static void ConfigureLogging(string logFilePath = "logs/DataFixter.log")
        {
            // 确保日志目录存在
            var logDir = Path.GetDirectoryName(logFilePath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // 构建配置
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();

            // 配置Serilog
            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .CreateLogger();

            Log.Information("日志系统已初始化");
        }

        /// <summary>
        /// 关闭日志记录器
        /// </summary>
        public static void CloseLogging()
        {
            Log.CloseAndFlush();
        }
    }
}
