namespace WorkPartner.Utils
{
    /// <summary>
    /// 日志级别枚举
    /// </summary>
    public enum LogLevel
    {
        /// <summary>
        /// 调试
        /// </summary>
        Debug,

        /// <summary>
        /// 信息
        /// </summary>
        Info,

        /// <summary>
        /// 警告
        /// </summary>
        Warning,

        /// <summary>
        /// 错误
        /// </summary>
        Error,

        /// <summary>
        /// 致命错误
        /// </summary>
        Fatal
    }

    /// <summary>
    /// 日志工具类
    /// </summary>
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static LogLevel _minLevel = LogLevel.Info;
        private static string? _logFilePath;
        private static bool _enableConsoleOutput = true;
        private static bool _enableFileOutput = true;

        /// <summary>
        /// 初始化日志器
        /// </summary>
        /// <param name="logFilePath">日志文件路径</param>
        /// <param name="minLevel">最小日志级别</param>
        /// <param name="enableConsole">是否启用控制台输出</param>
        /// <param name="enableFile">是否启用文件输出</param>
        public static void Initialize(string? logFilePath = null, LogLevel minLevel = LogLevel.Info, bool enableConsole = true, bool enableFile = true)
        {
            lock (_lock)
            {
                _logFilePath = logFilePath;
                _minLevel = minLevel;
                _enableConsoleOutput = enableConsole;
                _enableFileOutput = enableFile;

                if (_enableFileOutput && !string.IsNullOrEmpty(_logFilePath))
                {
                    try
                    {
                        var logDir = Path.GetDirectoryName(_logFilePath);
                        if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                        {
                            Directory.CreateDirectory(logDir);
                        }
                    }
                    catch
                    {
                        // 忽略目录创建失败
                    }
                }
            }
        }

        /// <summary>
        /// 记录调试日志
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        public static void Debug(string message, params object[] args)
        {
            Log(LogLevel.Debug, message, args);
        }

        /// <summary>
        /// 记录信息日志
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        public static void Info(string message, params object[] args)
        {
            Log(LogLevel.Info, message, args);
        }

        /// <summary>
        /// 记录警告日志
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        public static void Warning(string message, params object[] args)
        {
            Log(LogLevel.Warning, message, args);
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        public static void Error(string message, params object[] args)
        {
            Log(LogLevel.Error, message, args);
        }

        /// <summary>
        /// 记录错误日志（带异常）
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="exception">异常</param>
        /// <param name="args">格式化参数</param>
        public static void Error(string message, Exception exception, params object[] args)
        {
            var fullMessage = string.Format(message, args) + $"\nException: {exception}";
            Log(LogLevel.Error, fullMessage);
        }

        /// <summary>
        /// 记录致命错误日志
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        public static void Fatal(string message, params object[] args)
        {
            Log(LogLevel.Fatal, message, args);
        }

        /// <summary>
        /// 记录日志
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <param name="message">日志消息</param>
        /// <param name="args">格式化参数</param>
        private static void Log(LogLevel level, string message, params object[] args)
        {
            if (level < _minLevel)
            {
                return;
            }

            var formattedMessage = args.Length > 0 ? string.Format(message, args) : message;
            var logEntry = CreateLogEntry(level, formattedMessage);

            lock (_lock)
            {
                if (_enableConsoleOutput)
                {
                    WriteToConsole(level, logEntry);
                }

                if (_enableFileOutput && !string.IsNullOrEmpty(_logFilePath))
                {
                    WriteToFile(logEntry);
                }
            }
        }

        /// <summary>
        /// 创建日志条目
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <param name="message">日志消息</param>
        /// <returns>日志条目</returns>
        private static string CreateLogEntry(LogLevel level, string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var levelStr = level.ToString().ToUpper().PadRight(5);
            var threadId = Environment.CurrentManagedThreadId;
            return $"[{timestamp}] [{levelStr}] [Thread-{threadId}] {message}";
        }

        /// <summary>
        /// 写入控制台
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <param name="logEntry">日志条目</param>
        private static void WriteToConsole(LogLevel level, string logEntry)
        {
            var originalColor = Console.ForegroundColor;
            var color = level switch
            {
                LogLevel.Debug => ConsoleColor.Gray,
                LogLevel.Info => ConsoleColor.White,
                LogLevel.Warning => ConsoleColor.Yellow,
                LogLevel.Error => ConsoleColor.Red,
                LogLevel.Fatal => ConsoleColor.DarkRed,
                _ => ConsoleColor.White
            };

            Console.ForegroundColor = color;
            Console.WriteLine(logEntry);
            Console.ForegroundColor = originalColor;
        }

        /// <summary>
        /// 写入文件
        /// </summary>
        /// <param name="logEntry">日志条目</param>
        private static void WriteToFile(string logEntry)
        {
            try
            {
                File.AppendAllText(_logFilePath!, logEntry + Environment.NewLine);
            }
            catch
            {
                // 忽略文件写入失败
            }
        }

        /// <summary>
        /// 记录进度信息
        /// </summary>
        /// <param name="current">当前进度</param>
        /// <param name="total">总进度</param>
        /// <param name="message">进度消息</param>
        public static void Progress(int current, int total, string message = "处理进度")
        {
            if (total <= 0) return;

            var percentage = (double)current / total * 100;
            var progressBar = CreateProgressBar(percentage);
            var progressMessage = $"{message}: {current}/{total} ({percentage:F1}%) {progressBar}";

            // 清除当前行并重新写入
            Console.Write($"\r{progressMessage}");
            
            if (current >= total)
            {
                Console.WriteLine(); // 完成后换行
            }
        }

        /// <summary>
        /// 创建进度条
        /// </summary>
        /// <param name="percentage">百分比</param>
        /// <returns>进度条字符串</returns>
        private static string CreateProgressBar(double percentage)
        {
            const int barLength = 20;
            var filledLength = (int)(percentage / 100 * barLength);
            var bar = new string('█', filledLength) + new string('░', barLength - filledLength);
            return $"[{bar}]";
        }

        /// <summary>
        /// 记录性能信息
        /// </summary>
        /// <param name="operation">操作名称</param>
        /// <param name="elapsedTime">耗时</param>
        /// <param name="details">详细信息</param>
        public static void Performance(string operation, TimeSpan elapsedTime, string? details = null)
        {
            var message = $"性能 - {operation}: {elapsedTime.TotalMilliseconds:F2}ms";
            if (!string.IsNullOrEmpty(details))
            {
                message += $" ({details})";
            }
            Info(message);
        }

        /// <summary>
        /// 记录内存使用情况
        /// </summary>
        /// <param name="operation">操作名称</param>
        public static void MemoryUsage(string operation = "当前")
        {
            var memory = GC.GetTotalMemory(false);
            var memoryMB = memory / (1024.0 * 1024.0);
            Info($"内存使用 - {operation}: {memoryMB:F2}MB");
        }

        /// <summary>
        /// 清理日志文件
        /// </summary>
        /// <param name="maxSizeMB">最大文件大小（MB）</param>
        /// <param name="maxDays">最大保留天数</param>
        public static void CleanupLogFile(double maxSizeMB = 100, int maxDays = 30)
        {
            if (string.IsNullOrEmpty(_logFilePath) || !File.Exists(_logFilePath))
            {
                return;
            }

            try
            {
                var fileInfo = new FileInfo(_logFilePath);
                var maxSizeBytes = maxSizeMB * 1024 * 1024;

                // 检查文件大小
                if (fileInfo.Length > maxSizeBytes)
                {
                    var backupPath = _logFilePath + ".backup";
                    if (File.Exists(backupPath))
                    {
                        File.Delete(backupPath);
                    }
                    File.Move(_logFilePath, backupPath);
                }

                // 检查文件时间
                if (fileInfo.CreationTime < DateTime.Now.AddDays(-maxDays))
                {
                    File.Delete(_logFilePath);
                }
            }
            catch
            {
                // 忽略清理失败
            }
        }
    }
} 