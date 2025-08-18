using Serilog;

namespace DataFixter.Services
{
    /// <summary>
    /// 双日志服务，用于分别处理控制台和文件的日志输出
    /// 控制台输出简洁友好，文件输出详细完整
    /// </summary>
    public class DualLoggerService
    {
        private readonly ILogger _consoleLogger;
        private readonly ILogger _fileLogger;
        private int _lastProgress = -1;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="serviceType">服务类型</param>
        public DualLoggerService(Type serviceType)
        {
            _consoleLogger = Log.ForContext(serviceType).ForContext("Output", "Console");
            _fileLogger = Log.ForContext(serviceType).ForContext("Output", "File");
        }

        /// <summary>
        /// 控制台：简洁信息
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void ConsoleInfo(string message, params object[] args)
        {
            _consoleLogger.Information(message, args);
        }

        /// <summary>
        /// 控制台：警告信息
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void ConsoleWarning(string message, params object[] args)
        {
            _consoleLogger.Warning(message, args);
        }

        /// <summary>
        /// 控制台：错误信息
        /// </summary>
        /// <param name="ex">异常</param>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void ConsoleError(Exception ex, string message, params object[] args)
        {
            _consoleLogger.Error(ex, message, args);
        }

        /// <summary>
        /// 文件：详细信息
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void FileInfo(string message, params object[] args)
        {
            _fileLogger.Information(message, args);
        }

        /// <summary>
        /// 文件：警告信息
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void FileWarning(string message, params object[] args)
        {
            _fileLogger.Warning(message, args);
        }

        /// <summary>
        /// 文件：错误信息
        /// </summary>
        /// <param name="ex">异常</param>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void FileError(Exception ex, string message, params object[] args)
        {
            _fileLogger.Error(ex, message, args);
        }

        /// <summary>
        /// 文件：调试信息
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void FileDebug(string message, params object[] args)
        {
            _fileLogger.Debug(message, args);
        }

        /// <summary>
        /// 进度更新（只显示在控制台）
        /// </summary>
        /// <param name="current">当前进度</param>
        /// <param name="total">总数量</param>
        /// <param name="operation">操作名称</param>
        public void UpdateProgress(int current, int total, string operation = "处理")
        {
            var progress = (int)((double)current / total * 100);
            
            // 只在进度变化时更新控制台
            if (progress != _lastProgress)
            {
                Console.Write($"\r{operation}进度: {progress}%");
                if (progress == 100)
                {
                    Console.WriteLine(); // 换行
                }
                _lastProgress = progress;
            }
        }

        /// <summary>
        /// 显示步骤信息（控制台）
        /// </summary>
        /// <param name="step">步骤信息</param>
        public void ShowStep(string step)
        {
            Console.WriteLine($"\n{step}");
        }

        /// <summary>
        /// 显示完成信息（控制台）
        /// </summary>
        /// <param name="message">完成消息</param>
        public void ShowComplete(string message)
        {
            Console.WriteLine($"✅ {message}");
        }

        /// <summary>
        /// 显示错误信息（控制台）
        /// </summary>
        /// <param name="message">错误消息</param>
        public void ShowError(string message)
        {
            Console.WriteLine($"❌ {message}");
        }

        /// <summary>
        /// 显示警告信息（控制台）
        /// </summary>
        /// <param name="message">警告消息</param>
        public void ShowWarning(string message)
        {
            Console.WriteLine($"⚠️ {message}");
        }

        /// <summary>
        /// 同时记录到控制台和文件
        /// </summary>
        /// <param name="consoleMessage">控制台消息（简洁）</param>
        /// <param name="fileMessage">文件消息（详细）</param>
        /// <param name="args">参数</param>
        public void LogBoth(string consoleMessage, string fileMessage, params object[] args)
        {
            ConsoleInfo(consoleMessage, args);
            FileInfo(fileMessage, args);
        }

        /// <summary>
        /// 记录操作开始（控制台简洁，文件详细）
        /// </summary>
        /// <param name="operation">操作名称</param>
        /// <param name="details">详细信息</param>
        public void LogOperationStart(string operation, string details = "")
        {
            ConsoleInfo("开始{Operation}", operation);
            if (!string.IsNullOrEmpty(details))
            {
                FileInfo("开始{Operation}，详细信息：{Details}", operation, details);
            }
            else
            {
                FileInfo("开始{Operation}", operation);
            }
        }

        /// <summary>
        /// 记录操作完成（控制台简洁，文件详细）
        /// </summary>
        /// <param name="operation">操作名称</param>
        /// <param name="result">结果信息</param>
        /// <param name="details">详细信息</param>
        public void LogOperationComplete(string operation, string result, string details = "")
        {
            ShowComplete($"{operation}完成：{result}");
            if (!string.IsNullOrEmpty(details))
            {
                FileInfo("{Operation}完成，结果：{Result}，详细信息：{Details}", operation, result, details);
            }
            else
            {
                FileInfo("{Operation}完成，结果：{Result}", operation, result);
            }
        }
    }
}
