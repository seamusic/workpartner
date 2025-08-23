using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DataExport.Services
{
    /// <summary>
    /// 命令行界面服务
    /// </summary>
    public class CommandLineInterface
    {
        private readonly ILogger<CommandLineInterface> _logger;
        private readonly ExportModeManager _exportModeManager;
        private readonly ExportHistoryService _historyService;
        private readonly DataQualityService _qualityService;
        private readonly ExportFileManager _fileManager;
        private readonly CookieValidationService _cookieValidationService;
        private bool _isRunning = true;

        public CommandLineInterface(
            ILogger<CommandLineInterface> logger,
            ExportModeManager exportModeManager,
            ExportHistoryService historyService,
            DataQualityService qualityService,
            ExportFileManager fileManager,
            CookieValidationService cookieValidationService)
        {
            _logger = logger;
            _exportModeManager = exportModeManager;
            _historyService = historyService;
            _qualityService = qualityService;
            _fileManager = fileManager;
            _cookieValidationService = cookieValidationService;
        }

        /// <summary>
        /// 启动命令行界面
        /// </summary>
        public async Task RunAsync()
        {
            try
            {
                _logger.LogInformation("启动命令行界面...");
                ShowWelcomeMessage();
                ShowMainMenu();

                while (_isRunning)
                {
                    try
                    {
                        Console.Write("\n请输入命令 (输入 'help' 查看帮助，或使用数字快捷命令): ");
                        var input = Console.ReadLine()?.Trim();

                        if (string.IsNullOrEmpty(input))
                            continue;

                        await ProcessCommandAsync(input);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "命令执行失败");
                        Console.WriteLine($"错误: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "命令行界面运行失败");
                Console.WriteLine($"严重错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示欢迎信息
        /// </summary>
        private void ShowWelcomeMessage()
        {
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("          数据导出工具 - 命令行界面");
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("版本: 1.0.0");
            Console.WriteLine("基于基坑监测系统的自动化数据导出工具");
            Console.WriteLine("支持多种导出模式、历史记录管理、数据质量检查等功能");
            Console.WriteLine("=".PadRight(60, '='));
        }

        /// <summary>
        /// 显示主菜单
        /// </summary>
        private void ShowMainMenu()
        {
            Console.WriteLine("\n=== 主要功能菜单 ===");
            Console.WriteLine("  1. 导出管理 (export)     - 执行各种导出模式");
            Console.WriteLine("  2. 历史记录 (history)    - 查看和管理导出历史");
            Console.WriteLine("  3. 数据质量 (quality)    - 检查导出数据质量");
            Console.WriteLine("  4. 文件管理 (file)       - 管理导出文件");
            Console.WriteLine("  5. 系统状态 (status)     - 查看系统状态和配置");
            Console.WriteLine("  6. Cookie验证 (cookie)   - 验证API Cookie有效性");
            Console.WriteLine("  7. 帮助信息 (help)       - 查看详细帮助");
            Console.WriteLine("  8. 清屏 (clear)          - 清空屏幕");
            Console.WriteLine("  0. 退出系统 (exit)       - 退出程序");
            Console.WriteLine("\n提示: 可以直接输入数字快捷命令，或输入完整命令");
        }

        /// <summary>
        /// 处理用户命令
        /// </summary>
        private async Task ProcessCommandAsync(string command)
        {
            var parts = command.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var mainCommand = parts[0].ToLower();

            switch (mainCommand)
            {
                // 数字快捷命令
                case "1":
                    await HandleExportCommandsAsync(new[] { "export", "help" });
                    break;

                case "2":
                    await HandleHistoryCommandsAsync(new[] { "history", "help" });
                    break;

                case "3":
                    await HandleQualityCommandsAsync(new[] { "quality", "help" });
                    break;

                case "4":
                    await HandleFileCommandsAsync(new[] { "file", "help" });
                    break;

                case "5":
                    await HandleStatusCommandsAsync(new[] { "status", "help" });
                    break;

                case "6":
                    await HandleCookieValidationAsync();
                    break;

                case "7":
                    ShowHelp();
                    break;

                case "8":
                    Console.Clear();
                    ShowWelcomeMessage();
                    ShowMainMenu();
                    break;

                case "0":
                    _isRunning = false;
                    Console.WriteLine("感谢使用数据导出工具，再见！");
                    break;

                // 完整命令
                case "help":
                case "h":
                    if (parts.Length > 1)
                    {
                        ShowSpecificHelp(parts[1]);
                    }
                    else
                    {
                        ShowHelp();
                    }
                    break;

                case "export":
                case "e":
                    await HandleExportCommandsAsync(parts);
                    break;

                case "history":
                case "hist":
                    await HandleHistoryCommandsAsync(parts);
                    break;

                case "quality":
                case "q":
                    await HandleQualityCommandsAsync(parts);
                    break;

                case "file":
                case "f":
                    await HandleFileCommandsAsync(parts);
                    break;

                case "status":
                case "s":
                    await HandleStatusCommandsAsync(parts);
                    break;

                case "cookie":
                case "c":
                    await HandleCookieValidationAsync();
                    break;

                case "clear":
                case "cls":
                    Console.Clear();
                    ShowWelcomeMessage();
                    ShowMainMenu();
                    break;

                case "exit":
                case "quit":
                    _isRunning = false;
                    Console.WriteLine("感谢使用数据导出工具，再见！");
                    break;

                default:
                    Console.WriteLine($"未知命令: {mainCommand}");
                    Console.WriteLine("输入 'help' 查看可用命令，或使用数字快捷命令");
                    break;
            }
        }

        /// <summary>
        /// 处理导出相关命令
        /// </summary>
        private async Task HandleExportCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowExportHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "list":
                case "ls":
                    await ListExportModesAsync();
                    break;

                case "run":
                case "execute":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: export run <模式名称>");
                        return;
                    }
                    await ExecuteExportModeAsync(parts[2]);
                    break;

                case "default":
                    await ExecuteDefaultExportAsync();
                    break;

                case "all":
                    await ExecuteAllExportModesAsync();
                    break;

                case "validate":
                    await ValidateExportModesAsync();
                    break;

                default:
                    Console.WriteLine($"未知的导出命令: {subCommand}");
                    ShowExportHelp();
                    break;
            }
        }

        /// <summary>
        /// 处理历史记录相关命令
        /// </summary>
        private async Task HandleHistoryCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowHistoryHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "list":
                case "ls":
                    await ListExportHistoriesAsync();
                    break;

                case "show":
                case "view":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: history show <记录ID>");
                        return;
                    }
                    await ShowExportHistoryAsync(parts[2]);
                    break;

                case "stats":
                case "statistics":
                    await ShowHistoryStatisticsAsync();
                    break;

                case "cleanup":
                    await CleanupExpiredHistoriesAsync();
                    break;

                case "export":
                    await ExportHistoriesToFileAsync();
                    break;

                default:
                    Console.WriteLine($"未知的历史记录命令: {subCommand}");
                    ShowHistoryHelp();
                    break;
            }
        }

        /// <summary>
        /// 处理数据质量相关命令
        /// </summary>
        private async Task HandleQualityCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowQualityHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "check":
                    await CheckDataQualityAsync();
                    break;

                case "report":
                    await GenerateQualityReportAsync();
                    break;

                default:
                    Console.WriteLine($"未知的数据质量命令: {subCommand}");
                    ShowQualityHelp();
                    break;
            }
        }

        /// <summary>
        /// 处理文件管理相关命令
        /// </summary>
        private async Task HandleFileCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowFileHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "info":
                case "status":
                    await ShowFileStorageInfoAsync();
                    break;

                case "organize":
                    await OrganizeFilesAsync();
                    break;

                case "cleanup":
                    await CleanupExpiredFilesAsync();
                    break;

                case "search":
                    await SearchFilesAsync(parts);
                    break;

                case "archive":
                    await ArchiveFilesAsync(parts);
                    break;

                default:
                    Console.WriteLine($"未知的文件管理命令: {subCommand}");
                    ShowFileHelp();
                    break;
            }
        }

        /// <summary>
        /// 处理状态相关命令
        /// </summary>
        private async Task HandleStatusCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowStatusHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "system":
                    await ShowSystemStatusAsync();
                    break;

                case "config":
                    await ShowConfigurationAsync();
                    break;

                case "storage":
                    await ShowStorageStatusAsync();
                    break;

                default:
                    Console.WriteLine($"未知的状态命令: {subCommand}");
                    ShowStatusHelp();
                    break;
            }
        }

        /// <summary>
        /// 显示帮助信息
        /// </summary>
        private void ShowHelp()
        {
            Console.WriteLine("\n=== 命令帮助 ===");
            Console.WriteLine("数字快捷命令:");
            Console.WriteLine("  1                    - 导出管理");
            Console.WriteLine("  2                    - 历史记录管理");
            Console.WriteLine("  3                    - 数据质量管理");
            Console.WriteLine("  4                    - 文件管理");
            Console.WriteLine("  5                    - 系统状态");
            Console.WriteLine("  6                    - Cookie验证");
            Console.WriteLine("  7                    - 显示此帮助信息");
            Console.WriteLine("  8                    - 清屏");
            Console.WriteLine("  0                    - 退出系统");
            Console.WriteLine("\n完整命令:");
            Console.WriteLine("  help, h              - 显示此帮助信息");
            Console.WriteLine("  clear, cls           - 清屏");
            Console.WriteLine("  exit, quit           - 退出系统");
            Console.WriteLine("\n功能命令:");
            Console.WriteLine("  export <命令>        - 导出管理");
            Console.WriteLine("  history <命令>       - 历史记录管理");
            Console.WriteLine("  quality <命令>       - 数据质量管理");
            Console.WriteLine("  file <命令>          - 文件管理");
            Console.WriteLine("  status <命令>        - 系统状态");
            Console.WriteLine("  cookie               - Cookie验证");
            Console.WriteLine("\n输入 'help <功能>' 查看具体功能的详细帮助");
            Console.WriteLine("示例: help export, help history");
        }

        /// <summary>
        /// 显示特定功能的帮助信息
        /// </summary>
        private void ShowSpecificHelp(string function)
        {
            switch (function.ToLower())
            {
                case "export":
                case "e":
                    ShowExportHelp();
                    break;

                case "history":
                case "hist":
                    ShowHistoryHelp();
                    break;

                case "quality":
                case "q":
                    ShowQualityHelp();
                    break;

                case "file":
                case "f":
                    ShowFileHelp();
                    break;

                case "status":
                case "s":
                    ShowStatusHelp();
                    break;

                case "cookie":
                case "c":
                    ShowCookieHelp();
                    break;

                default:
                    Console.WriteLine($"未知功能: {function}");
                    Console.WriteLine("可用的功能: export, history, quality, file, status, cookie");
                    Console.WriteLine("示例: help export, help history");
                    break;
            }
        }

        /// <summary>
        /// 显示导出帮助
        /// </summary>
        private void ShowExportHelp()
        {
            Console.WriteLine("\n=== 导出管理帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  export help           - 显示此帮助信息");
            Console.WriteLine("  export list, ls       - 列出所有导出模式");
            Console.WriteLine("  export run <模式名称> - 执行指定导出模式");
            Console.WriteLine("  export default        - 执行默认导出模式");
            Console.WriteLine("  export all            - 执行所有导出模式");
            Console.WriteLine("  export validate       - 验证导出模式配置");
            Console.WriteLine("\n导出模式:");
            Console.WriteLine("  AllProjects           - 导出所有项目");
            Console.WriteLine("  SingleProject         - 导出单个项目");
            Console.WriteLine("  CustomTimeRange       - 自定义时间范围导出");
            Console.WriteLine("  BatchExport           - 批量导出");
            Console.WriteLine("  IncrementalExport     - 增量导出");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  export list           - 查看所有导出模式");
            Console.WriteLine("  export run AllProjects - 执行所有项目导出");
        }

        /// <summary>
        /// 显示历史记录帮助
        /// </summary>
        private void ShowHistoryHelp()
        {
            Console.WriteLine("\n=== 历史记录管理帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  history help          - 显示此帮助信息");
            Console.WriteLine("  history list, ls      - 列出导出历史记录");
            Console.WriteLine("  history show <ID>     - 显示指定记录详情");
            Console.WriteLine("  history stats          - 显示统计信息");
            Console.WriteLine("  history cleanup       - 清理过期记录");
            Console.WriteLine("  history export        - 导出历史记录到文件");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  history list          - 查看历史记录列表");
            Console.WriteLine("  history show abc123   - 查看ID为abc123的记录详情");
            Console.WriteLine("  history stats         - 查看统计信息");
        }

        /// <summary>
        /// 显示数据质量帮助
        /// </summary>
        private void ShowQualityHelp()
        {
            Console.WriteLine("\n=== 数据质量管理帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  quality help          - 显示此帮助信息");
            Console.WriteLine("  quality check         - 检查数据质量");
            Console.WriteLine("  quality report        - 生成质量报告");
            Console.WriteLine("\n功能说明:");
            Console.WriteLine("  check                 - 对导出的文件进行质量检查");
            Console.WriteLine("  report                - 生成详细的质量检查报告");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  quality check         - 执行数据质量检查");
            Console.WriteLine("  quality report        - 生成质量报告");
        }

        /// <summary>
        /// 显示文件管理帮助
        /// </summary>
        private void ShowFileHelp()
        {
            Console.WriteLine("\n=== 文件管理帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  file help             - 显示此帮助信息");
            Console.WriteLine("  file info, status     - 显示存储信息");
            Console.WriteLine("  file organize         - 整理文件");
            Console.WriteLine("  file cleanup          - 清理过期文件");
            Console.WriteLine("  file search <条件>    - 搜索文件");
            Console.WriteLine("  file archive <名称> <文件> - 归档文件");
            Console.WriteLine("\n搜索条件格式:");
            Console.WriteLine("  size>1MB              - 文件大小大于1MB");
            Console.WriteLine("  size<100MB            - 文件大小小于100MB");
            Console.WriteLine("  date>2024-01-01       - 修改日期晚于2024-01-01");
            Console.WriteLine("  date<2024-12-31       - 修改日期早于2024-12-31");
            Console.WriteLine("  name=test             - 文件名包含test");
            Console.WriteLine("  ext=xlsx              - 文件扩展名为xlsx");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  file info             - 查看存储信息");
            Console.WriteLine("  file search size>1MB,date>2024-01-01 - 搜索大文件");
            Console.WriteLine("  file archive backup file1.xlsx file2.xlsx - 归档文件");
        }

        /// <summary>
        /// 显示状态帮助
        /// </summary>
        private void ShowStatusHelp()
        {
            Console.WriteLine("\n=== 系统状态帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  status help           - 显示此帮助信息");
            Console.WriteLine("  status system         - 显示系统状态");
            Console.WriteLine("  status config         - 显示配置信息");
            Console.WriteLine("  status storage        - 显示存储状态");
            Console.WriteLine("\n功能说明:");
            Console.WriteLine("  system                - 显示系统基本信息");
            Console.WriteLine("  config                - 显示当前配置信息");
            Console.WriteLine("  storage               - 显示存储使用情况");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  status system         - 查看系统状态");
            Console.WriteLine("  status config         - 查看配置信息");
            Console.WriteLine("  status storage        - 查看存储状态");
        }

        #region 导出管理命令实现

        private async Task ListExportModesAsync()
        {
            Console.WriteLine("\n=== 导出模式列表 ===");
            var modes = _exportModeManager.GetExportModes();
            
            if (!modes.Any())
            {
                Console.WriteLine("没有找到可用的导出模式");
                return;
            }

            foreach (var mode in modes.OrderBy(m => m.Priority))
            {
                Console.WriteLine($"  {mode.Mode,-15} - {mode.Description,-30} (优先级: {mode.Priority})");
            }
        }

        private async Task ExecuteExportModeAsync(string modeName)
        {
            Console.WriteLine($"\n开始执行导出模式: {modeName}");
            try
            {
                var result = await _exportModeManager.ExecuteModeAsync(modeName);
                Console.WriteLine($"执行完成: {result.GetSummary()}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行失败: {ex.Message}");
            }
        }

        private async Task ExecuteDefaultExportAsync()
        {
            Console.WriteLine("\n开始执行默认导出模式");
            try
            {
                var result = await _exportModeManager.ExecuteDefaultModeAsync();
                Console.WriteLine($"执行完成: {result.GetSummary()}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行失败: {ex.Message}");
            }
        }

        private async Task ExecuteAllExportModesAsync()
        {
            Console.WriteLine("\n开始执行所有导出模式");
            try
            {
                var results = await _exportModeManager.ExecuteAllModesAsync();
                Console.WriteLine($"执行完成，共 {results.Count} 个模式:");
                foreach (var result in results)
                {
                    Console.WriteLine($"  {result.Mode}: {result.GetSummary()}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行失败: {ex.Message}");
            }
        }

        private async Task ValidateExportModesAsync()
        {
            Console.WriteLine("\n开始验证导出模式配置");
            try
            {
                var validationResults = _exportModeManager.ValidateModes();
                if (!validationResults.Any())
                {
                    Console.WriteLine("✓ 所有导出模式配置验证通过");
                }
                else
                {
                    Console.WriteLine("⚠ 发现配置验证问题:");
                    foreach (var result in validationResults)
                    {
                        Console.WriteLine($"  - {result.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"验证失败: {ex.Message}");
            }
        }

        #endregion

        #region 历史记录命令实现

        private async Task ListExportHistoriesAsync()
        {
            Console.WriteLine("\n=== 导出历史记录 ===");
            try
            {
                var query = new ExportHistoryQuery { PageSize = 10 };
                var (histories, totalCount) = await _historyService.QueryHistoriesAsync(query);
                
                if (!histories.Any())
                {
                    Console.WriteLine("没有找到导出历史记录");
                    return;
                }

                Console.WriteLine($"共 {totalCount} 条记录，显示前 {histories.Count} 条:");
                foreach (var history in histories)
                {
                    Console.WriteLine($"  {history.Id,-8} | {history.ExportMode,-12} | {history.Status,-8} | {history.StartTime:yyyy-MM-dd HH:mm} | {history.GetSummary()}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取历史记录失败: {ex.Message}");
            }
        }

        private async Task ShowExportHistoryAsync(string id)
        {
            Console.WriteLine($"\n=== 导出历史记录详情: {id} ===");
            try
            {
                var history = _historyService.GetHistory(id);
                if (history == null)
                {
                    Console.WriteLine("未找到指定的历史记录");
                    return;
                }

                Console.WriteLine($"记录ID: {history.Id}");
                Console.WriteLine($"导出模式: {history.ExportMode}");
                Console.WriteLine($"描述: {history.Description}");
                Console.WriteLine($"状态: {history.Status}");
                Console.WriteLine($"开始时间: {history.StartTime:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"结束时间: {history.EndTime:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"总数量: {history.TotalCount}");
                Console.WriteLine($"成功数量: {history.SuccessCount}");
                Console.WriteLine($"失败数量: {history.FailedCount}");
                Console.WriteLine($"成功率: {history.SuccessRate:F1}%");
                Console.WriteLine($"执行耗时: {history.Duration}ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取历史记录详情失败: {ex.Message}");
            }
        }

        private async Task ShowHistoryStatisticsAsync()
        {
            Console.WriteLine("\n=== 导出历史统计信息 ===");
            try
            {
                var stats = await _historyService.GetStatisticsAsync();
                Console.WriteLine($"总记录数: {stats.TotalRecords}");
                Console.WriteLine($"成功记录数: {stats.SuccessRecords}");
                Console.WriteLine($"失败记录数: {stats.FailedRecords}");
                Console.WriteLine($"成功率: {stats.SuccessRate:F1}%");
                Console.WriteLine($"总导出文件数: {stats.TotalExportedFiles}");
                Console.WriteLine($"总合并文件数: {stats.TotalMergedFiles}");
                Console.WriteLine($"平均执行时间: {stats.AverageDuration:F0}ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取统计信息失败: {ex.Message}");
            }
        }

        private async Task CleanupExpiredHistoriesAsync()
        {
            Console.WriteLine("\n开始清理过期历史记录");
            try
            {
                var retentionPeriod = TimeSpan.FromDays(30);
                var count = await _historyService.CleanupExpiredHistoriesAsync(retentionPeriod);
                Console.WriteLine($"清理完成，删除了 {count} 条过期记录");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理失败: {ex.Message}");
            }
        }

        private async Task ExportHistoriesToFileAsync()
        {
            Console.WriteLine("\n开始导出历史记录到文件");
            try
            {
                var query = new ExportHistoryQuery();
                var outputPath = "./exports";
                var filePath = await _historyService.ExportHistoriesToFileAsync(query, outputPath);
                Console.WriteLine($"导出完成: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出失败: {ex.Message}");
            }
        }

        #endregion

        #region 数据质量命令实现

        private async Task CheckDataQualityAsync()
        {
            Console.WriteLine("\n开始检查数据质量");
            try
            {
                // 这里需要从实际的导出结果中获取数据
                // 暂时使用空列表作为示例
                var exportResults = new List<ExportResult>();
                var mergeResults = new List<MergeResult>();
                
                var report = await _qualityService.CheckExportQualityAsync(exportResults, mergeResults);
                Console.WriteLine($"质量检查完成，总体评分: {report.OverallQualityScore:F1}/100");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"质量检查失败: {ex.Message}");
            }
        }

        private async Task GenerateQualityReportAsync()
        {
            Console.WriteLine("\n开始生成质量报告");
            try
            {
                // 这里需要从实际的导出结果中获取数据
                var exportResults = new List<ExportResult>();
                var mergeResults = new List<MergeResult>();
                
                var report = await _qualityService.CheckExportQualityAsync(exportResults, mergeResults);
                var outputPath = "./exports";
                var filePath = await _qualityService.GenerateQualityReportAsync(report, outputPath);
                Console.WriteLine($"质量报告生成完成: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成质量报告失败: {ex.Message}");
            }
        }

        #endregion

        #region 文件管理命令实现

        private async Task ShowFileStorageInfoAsync()
        {
            Console.WriteLine("\n=== 文件存储信息 ===");
            try
            {
                var info = await _fileManager.GetStorageInfoAsync();
                Console.WriteLine($"基础目录: {info.BaseDirectory}");
                Console.WriteLine($"总文件数: {info.TotalFiles}");
                Console.WriteLine($"总大小: {FormatFileSize(info.TotalSize)}");
                Console.WriteLine($"小文件 (<1MB): {info.SmallFiles}");
                Console.WriteLine($"中等文件 (1MB-10MB): {info.MediumFiles}");
                Console.WriteLine($"大文件 (>10MB): {info.LargeFiles}");
                
                if (info.FilesByExtension.Any())
                {
                    Console.WriteLine("\n按扩展名统计:");
                    foreach (var ext in info.FilesByExtension.OrderByDescending(x => x.Value))
                    {
                        Console.WriteLine($"  {ext.Key}: {ext.Value} 个文件");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取存储信息失败: {ex.Message}");
            }
        }

        private async Task OrganizeFilesAsync()
        {
            Console.WriteLine("\n开始整理文件");
            try
            {
                var result = await _fileManager.OrganizeFilesAsync();
                Console.WriteLine($"文件整理完成:");
                Console.WriteLine($"  处理文件数: {result.ProcessedFiles}");
                Console.WriteLine($"  移动文件数: {result.MovedFiles}");
                Console.WriteLine($"  删除文件数: {result.DeletedCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文件整理失败: {ex.Message}");
            }
        }

        private async Task CleanupExpiredFilesAsync()
        {
            Console.WriteLine("\n开始清理过期文件");
            try
            {
                var retentionPeriod = TimeSpan.FromDays(90);
                var result = await _fileManager.CleanupExpiredFilesAsync(retentionPeriod);
                Console.WriteLine($"过期文件清理完成:");
                Console.WriteLine($"  删除文件数: {result.DeletedCount}");
                Console.WriteLine($"  失败文件数: {result.FailedCount}");
                Console.WriteLine($"  释放空间: {FormatFileSize(result.TotalSize)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理过期文件失败: {ex.Message}");
            }
        }

        private async Task SearchFilesAsync(string[] parts)
        {
            if (parts.Length < 3)
            {
                Console.WriteLine("用法: file search <条件>");
                Console.WriteLine("条件示例: size>1MB, date>2024-01-01, name=test");
                return;
            }

            Console.WriteLine($"\n开始搜索文件，条件: {parts[2]}");
            try
            {
                var criteria = ParseSearchCriteria(parts[2]);
                var files = await _fileManager.FindFilesAsync(criteria);
                Console.WriteLine($"搜索完成，找到 {files.Count} 个文件:");
                
                foreach (var file in files.Take(20)) // 只显示前20个
                {
                    Console.WriteLine($"  {file.Name} ({FormatFileSize(file.Length)}) - {file.LastWriteTime:yyyy-MM-dd HH:mm}");
                }
                
                if (files.Count > 20)
                {
                    Console.WriteLine($"  ... 还有 {files.Count - 20} 个文件");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文件搜索失败: {ex.Message}");
            }
        }

        private async Task ArchiveFilesAsync(string[] parts)
        {
            if (parts.Length < 4)
            {
                Console.WriteLine("用法: file archive <归档名称> <文件路径1> [文件路径2] ...");
                return;
            }

            var archiveName = parts[2];
            var filePaths = parts.Skip(3).ToList();
            
            Console.WriteLine($"\n开始归档文件: {archiveName}");
            Console.WriteLine($"文件数量: {filePaths.Count}");
            
            try
            {
                var result = await _fileManager.ArchiveFilesAsync(archiveName, filePaths);
                Console.WriteLine($"文件归档完成:");
                Console.WriteLine($"  归档路径: {result.ArchivePath}");
                Console.WriteLine($"  归档大小: {FormatFileSize(result.ArchiveSize)}");
                Console.WriteLine($"  源文件数: {result.SourceFileCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文件归档失败: {ex.Message}");
            }
        }

        #endregion

        #region 状态命令实现

        private async Task ShowSystemStatusAsync()
        {
            Console.WriteLine("\n=== 系统状态 ===");
            Console.WriteLine($"工作目录: {Environment.CurrentDirectory}");
            Console.WriteLine($".NET版本: {Environment.Version}");
            Console.WriteLine($"内存使用: {GC.GetTotalMemory(false) / 1024 / 1024:F1} MB");
        }

        private async Task ShowConfigurationAsync()
        {
            Console.WriteLine("\n=== 配置信息 ===");
            // 这里可以显示当前的配置信息
            Console.WriteLine("配置信息显示功能待实现");
        }

        private async Task ShowStorageStatusAsync()
        {
            Console.WriteLine("\n=== 存储状态 ===");
            try
            {
                var info = await _fileManager.GetStorageInfoAsync();
                Console.WriteLine($"存储目录: {info.BaseDirectory}");
                Console.WriteLine($"文件总数: {info.TotalFiles}");
                Console.WriteLine($"存储大小: {FormatFileSize(info.TotalSize)}");
                Console.WriteLine($"检查时间: {info.CheckTime:yyyy-MM-dd HH:mm:ss}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取存储状态失败: {ex.Message}");
            }
        }

        #endregion

        #region 辅助方法

        private FileSearchCriteria ParseSearchCriteria(string criteriaString)
        {
            var criteria = new FileSearchCriteria();
            var parts = criteriaString.Split(',', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.StartsWith("size>"))
                {
                    var sizeStr = trimmed.Substring(5);
                    if (long.TryParse(sizeStr.Replace("MB", "").Replace("KB", "").Replace("GB", ""), out var size))
                    {
                        if (sizeStr.Contains("MB")) size *= 1024 * 1024;
                        else if (sizeStr.Contains("KB")) size *= 1024;
                        else if (sizeStr.Contains("GB")) size *= 1024 * 1024 * 1024;
                        criteria.MinSize = size;
                    }
                }
                else if (trimmed.StartsWith("size<"))
                {
                    var sizeStr = trimmed.Substring(5);
                    if (long.TryParse(sizeStr.Replace("MB", "").Replace("KB", "").Replace("GB", ""), out var size))
                    {
                        if (sizeStr.Contains("MB")) size *= 1024 * 1024;
                        else if (sizeStr.Contains("KB")) size *= 1024;
                        else if (sizeStr.Contains("GB")) size *= 1024 * 1024 * 1024;
                        criteria.MaxSize = size;
                    }
                }
                else if (trimmed.StartsWith("date>"))
                {
                    var dateStr = trimmed.Substring(5);
                    if (DateTime.TryParse(dateStr, out var date))
                    {
                        criteria.MinDate = date;
                    }
                }
                else if (trimmed.StartsWith("date<"))
                {
                    var dateStr = trimmed.Substring(5);
                    if (DateTime.TryParse(dateStr, out var date))
                    {
                        criteria.MaxDate = date;
                    }
                }
                else if (trimmed.StartsWith("name="))
                {
                    criteria.FileNamePattern = trimmed.Substring(5);
                }
                else if (trimmed.StartsWith("ext="))
                {
                    criteria.Extension = trimmed.Substring(4);
                }
            }
            
            return criteria;
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;
            
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            
            return $"{len:0.##} {sizes[order]}";
        }

        #endregion

        #region Cookie验证命令实现

        /// <summary>
        /// 处理Cookie验证命令
        /// </summary>
        private async Task HandleCookieValidationAsync()
        {
            Console.WriteLine("\n=== Cookie验证 ===");
            Console.WriteLine("开始验证API设置的Cookie有效性...");
            
            try
            {
                var result = await _cookieValidationService.ValidateCookieAsync();
                _cookieValidationService.DisplayValidationResult(result);
                
                if (result.IsValid)
                {
                    Console.WriteLine("\n✓ Cookie验证成功！可以正常使用导出功能。");
                }
                else
                {
                    Console.WriteLine("\n⚠ Cookie验证失败！请检查以下问题：");
                    Console.WriteLine("  1. 确认Cookie是否已过期");
                    Console.WriteLine("  2. 检查appsettings.json中的Cookie配置");
                    Console.WriteLine("  3. 确认服务器地址是否正确");
                    Console.WriteLine("  4. 尝试重新登录系统获取新的Cookie");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cookie验证过程中发生异常: {ex.Message}");
                _logger.LogError(ex, "Cookie验证失败");
            }
        }

        /// <summary>
        /// 显示Cookie验证帮助
        /// </summary>
        private void ShowCookieHelp()
        {
            Console.WriteLine("\n=== Cookie验证帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  cookie                - 验证API Cookie有效性");
            Console.WriteLine("  c                     - 快捷命令");
            Console.WriteLine("\n功能说明:");
            Console.WriteLine("  验证配置的Cookie是否有效，包括:");
            Console.WriteLine("  - Cookie格式检查");
            Console.WriteLine("  - API连接测试");
            Console.WriteLine("  - 天气接口测试");
            Console.WriteLine("  - 认证状态验证");
            Console.WriteLine("\n验证项目:");
            Console.WriteLine("  - ASP.NET_SessionId");
            Console.WriteLine("  - Qianchen_ADMS_V7_Token");
            Console.WriteLine("  - Qianchen_ADMS_V7_Mark");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  cookie                - 执行Cookie验证");
            Console.WriteLine("  c                     - 快捷方式验证");
        }

        #endregion
    }
}
