using Serilog;
using Serilog.Events;

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

            // 配置Serilog
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
                .WriteTo.Console(
                    outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File(
                    path: logFilePath,
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                    retainedFileCountLimit: 30,
                    fileSizeLimitBytes: 10 * 1024 * 1024) // 10MB
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
