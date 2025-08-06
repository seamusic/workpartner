namespace WorkPartner.Utils
{
    /// <summary>
    /// 自定义异常类型
    /// </summary>
    public class WorkPartnerException : Exception
    {
        public string Category { get; }
        public string? FilePath { get; }
        public Dictionary<string, object> Context { get; }

        public WorkPartnerException(string category, string message, string? filePath = null, Exception? innerException = null) 
            : base(message, innerException)
        {
            Category = category;
            FilePath = filePath;
            Context = new Dictionary<string, object>();
        }

        public WorkPartnerException AddContext(string key, object value)
        {
            Context[key] = value;
            return this;
        }
    }

    /// <summary>
    /// 异常处理工具类
    /// </summary>
    public static class ExceptionHandler
    {
        private static readonly Dictionary<string, int> _errorCounts = new();
        private static readonly object _lock = new object();

        /// <summary>
        /// 处理文件读取异常
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="maxRetries">最大重试次数</param>
        /// <returns>操作结果</returns>
        public static async Task<T> HandleFileReadAsync<T>(Func<Task<T>> action, string filePath, int maxRetries = 3)
        {
            int attempts = 0;
            Exception? lastException = null;

            while (attempts < maxRetries)
            {
                try
                {
                    attempts++;
                    Logger.Debug($"尝试读取文件 (第{attempts}次): {Path.GetFileName(filePath)}");
                    return await action();
                }
                catch (FileNotFoundException ex)
                {
                    RecordError("FileNotFound", filePath);
                    Logger.Error($"文件不存在: {Path.GetFileName(filePath)}", ex);
                    throw new WorkPartnerException("FileNotFound", $"文件不存在: {Path.GetFileName(filePath)}", filePath, ex);
                }
                catch (UnauthorizedAccessException ex)
                {
                    RecordError("UnauthorizedAccess", filePath);
                    Logger.Error($"文件访问被拒绝: {Path.GetFileName(filePath)}", ex);
                    throw new WorkPartnerException("UnauthorizedAccess", $"文件访问被拒绝: {Path.GetFileName(filePath)}", filePath, ex);
                }
                catch (IOException ex) when (attempts < maxRetries)
                {
                    lastException = ex;
                    RecordError("IOError", filePath);
                    Logger.Warning($"文件读取IO错误 (第{attempts}次重试): {Path.GetFileName(filePath)} - {ex.Message}");
                    await Task.Delay(1000 * attempts); // 递增延迟
                }
                catch (Exception ex)
                {
                    RecordError("UnknownFileReadError", filePath);
                    Logger.Error($"文件读取未知错误: {Path.GetFileName(filePath)}", ex);
                    throw new WorkPartnerException("UnknownFileReadError", $"文件读取失败: {Path.GetFileName(filePath)}", filePath, ex);
                }
            }

            // 所有重试都失败
            RecordError("FileReadMaxRetriesExceeded", filePath);
            Logger.Error($"文件读取超过最大重试次数: {Path.GetFileName(filePath)}", lastException!);
            throw new WorkPartnerException("FileReadMaxRetriesExceeded", 
                $"文件读取失败，已重试{maxRetries}次: {Path.GetFileName(filePath)}", filePath, lastException);
        }

        /// <summary>
        /// 处理文件读取异常（同步版本）
        /// </summary>
        public static T HandleFileRead<T>(Func<T> action, string filePath, int maxRetries = 3)
        {
            return HandleFileReadAsync(() => Task.FromResult(action()), filePath, maxRetries).GetAwaiter().GetResult();
        }

        /// <summary>
        /// 处理文件写入异常
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="maxRetries">最大重试次数</param>
        public static async Task HandleFileWriteAsync(Func<Task> action, string filePath, int maxRetries = 3)
        {
            int attempts = 0;
            Exception? lastException = null;

            while (attempts < maxRetries)
            {
                try
                {
                    attempts++;
                    Logger.Debug($"尝试写入文件 (第{attempts}次): {Path.GetFileName(filePath)}");
                    await action();
                    return;
                }
                catch (UnauthorizedAccessException ex)
                {
                    RecordError("UnauthorizedAccess", filePath);
                    Logger.Error($"文件写入权限被拒绝: {Path.GetFileName(filePath)}", ex);
                    throw new WorkPartnerException("UnauthorizedAccess", $"文件写入权限被拒绝: {Path.GetFileName(filePath)}", filePath, ex);
                }
                catch (DirectoryNotFoundException ex)
                {
                    RecordError("DirectoryNotFound", filePath);
                    Logger.Error($"目标目录不存在: {Path.GetDirectoryName(filePath)}", ex);
                    
                    // 尝试创建目录
                    try
                    {
                        var dir = Path.GetDirectoryName(filePath);
                        if (!string.IsNullOrEmpty(dir))
                        {
                            Directory.CreateDirectory(dir);
                            Logger.Info($"已创建目录: {dir}");
                            continue; // 重新尝试写入
                        }
                    }
                    catch (Exception createEx)
                    {
                        Logger.Error($"创建目录失败: {Path.GetDirectoryName(filePath)}", createEx);
                    }
                    
                    throw new WorkPartnerException("DirectoryNotFound", $"目标目录不存在且无法创建: {Path.GetDirectoryName(filePath)}", filePath, ex);
                }
                catch (IOException ex) when (attempts < maxRetries)
                {
                    lastException = ex;
                    RecordError("IOError", filePath);
                    Logger.Warning($"文件写入IO错误 (第{attempts}次重试): {Path.GetFileName(filePath)} - {ex.Message}");
                    await Task.Delay(1000 * attempts); // 递增延迟
                }
                catch (Exception ex)
                {
                    RecordError("UnknownFileWriteError", filePath);
                    Logger.Error($"文件写入未知错误: {Path.GetFileName(filePath)}", ex);
                    throw new WorkPartnerException("UnknownFileWriteError", $"文件写入失败: {Path.GetFileName(filePath)}", filePath, ex);
                }
            }

            // 所有重试都失败
            RecordError("FileWriteMaxRetriesExceeded", filePath);
            Logger.Error($"文件写入超过最大重试次数: {Path.GetFileName(filePath)}", lastException!);
            throw new WorkPartnerException("FileWriteMaxRetriesExceeded", 
                $"文件写入失败，已重试{maxRetries}次: {Path.GetFileName(filePath)}", filePath, lastException);
        }

        /// <summary>
        /// 处理文件写入异常（同步版本）
        /// </summary>
        public static void HandleFileWrite(Action action, string filePath, int maxRetries = 3)
        {
            HandleFileWriteAsync(() => { action(); return Task.CompletedTask; }, filePath, maxRetries).GetAwaiter().GetResult();
        }

        /// <summary>
        /// 处理数据格式异常
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <param name="dataContext">数据上下文信息</param>
        /// <param name="filePath">相关文件路径</param>
        public static T HandleDataFormat<T>(Func<T> action, string dataContext, string? filePath = null)
        {
            try
            {
                Logger.Debug($"处理数据格式: {dataContext}");
                return action();
            }
            catch (FormatException ex)
            {
                RecordError("DataFormatError", filePath);
                Logger.Error($"数据格式错误 - {dataContext}: {ex.Message}", ex);
                throw new WorkPartnerException("DataFormatError", $"数据格式错误 - {dataContext}: {ex.Message}", filePath, ex)
                    .AddContext("DataContext", dataContext);
            }
            catch (InvalidCastException ex)
            {
                RecordError("DataCastError", filePath);
                Logger.Error($"数据类型转换错误 - {dataContext}: {ex.Message}", ex);
                throw new WorkPartnerException("DataCastError", $"数据类型转换错误 - {dataContext}: {ex.Message}", filePath, ex)
                    .AddContext("DataContext", dataContext);
            }
            catch (ArgumentException ex)
            {
                RecordError("DataArgumentError", filePath);
                Logger.Error($"数据参数错误 - {dataContext}: {ex.Message}", ex);
                throw new WorkPartnerException("DataArgumentError", $"数据参数错误 - {dataContext}: {ex.Message}", filePath, ex)
                    .AddContext("DataContext", dataContext);
            }
            catch (Exception ex)
            {
                RecordError("UnknownDataError", filePath);
                Logger.Error($"数据处理未知错误 - {dataContext}: {ex.Message}", ex);
                throw new WorkPartnerException("UnknownDataError", $"数据处理失败 - {dataContext}: {ex.Message}", filePath, ex)
                    .AddContext("DataContext", dataContext);
            }
        }

        /// <summary>
        /// 安全执行操作，捕获并记录所有异常
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <param name="operationName">操作名称</param>
        /// <param name="continueOnError">出错时是否继续</param>
        /// <returns>操作是否成功</returns>
        public static bool SafeExecute(Action action, string operationName, bool continueOnError = true)
        {
            try
            {
                Logger.Debug($"开始执行操作: {operationName}");
                action();
                Logger.Debug($"操作完成: {operationName}");
                return true;
            }
            catch (WorkPartnerException ex)
            {
                Logger.Error($"操作失败 ({operationName}): {ex.Message}", ex);
                if (!continueOnError) throw;
                return false;
            }
            catch (Exception ex)
            {
                RecordError("UnknownOperationError", null);
                Logger.Error($"操作发生未知错误 ({operationName}): {ex.Message}", ex);
                if (!continueOnError) throw;
                return false;
            }
        }

        /// <summary>
        /// 记录错误统计
        /// </summary>
        /// <param name="category">错误类别</param>
        /// <param name="filePath">相关文件路径</param>
        private static void RecordError(string category, string? filePath)
        {
            lock (_lock)
            {
                var key = string.IsNullOrEmpty(filePath) ? category : $"{category}:{Path.GetFileName(filePath)}";
                _errorCounts[key] = _errorCounts.GetValueOrDefault(key, 0) + 1;
            }
        }

        /// <summary>
        /// 获取错误统计
        /// </summary>
        /// <returns>错误统计信息</returns>
        public static Dictionary<string, int> GetErrorStatistics()
        {
            lock (_lock)
            {
                return new Dictionary<string, int>(_errorCounts);
            }
        }

        /// <summary>
        /// 清空错误统计
        /// </summary>
        public static void ClearErrorStatistics()
        {
            lock (_lock)
            {
                _errorCounts.Clear();
            }
        }

        /// <summary>
        /// 生成错误报告
        /// </summary>
        /// <returns>错误报告字符串</returns>
        public static string GenerateErrorReport()
        {
            var stats = GetErrorStatistics();
            if (!stats.Any())
            {
                return "✅ 未发现错误";
            }

            var report = new System.Text.StringBuilder();
            report.AppendLine("🚨 错误统计报告:");
            report.AppendLine(new string('-', 50));

            var totalErrors = stats.Values.Sum();
            report.AppendLine($"总错误数: {totalErrors}");
            report.AppendLine();

            var groupedErrors = stats.GroupBy(kvp => kvp.Key.Split(':')[0])
                                   .OrderByDescending(g => g.Sum(kvp => kvp.Value));

            foreach (var group in groupedErrors)
            {
                var categoryTotal = group.Sum(kvp => kvp.Value);
                report.AppendLine($"📊 {group.Key}: {categoryTotal} 次");
                
                foreach (var item in group.OrderByDescending(kvp => kvp.Value))
                {
                    var parts = item.Key.Split(':', 2);
                    if (parts.Length > 1)
                    {
                        report.AppendLine($"  • {parts[1]}: {item.Value} 次");
                    }
                }
                report.AppendLine();
            }

            return report.ToString();
        }
    }
}