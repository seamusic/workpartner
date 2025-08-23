using DataExport.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport.Services
{
    /// <summary>
    /// 配置文件编辑器服务
    /// </summary>
    public class ConfigurationEditor
    {
        private readonly ILogger<ConfigurationEditor> _logger;
        private readonly string _configDirectory;
        private readonly string _exportModesConfigPath;
        private readonly string _mainConfigPath;

        public ConfigurationEditor(ILogger<ConfigurationEditor> logger, string configDirectory = "./")
        {
            _logger = logger;
            _configDirectory = configDirectory;
            _exportModesConfigPath = Path.Combine(configDirectory, "appsettings.export-modes.json");
            _mainConfigPath = Path.Combine(configDirectory, "appsettings.json");
        }

        /// <summary>
        /// 启动配置编辑器
        /// </summary>
        public async Task RunEditorAsync()
        {
            try
            {
                _logger.LogInformation("启动配置编辑器...");
                ShowWelcomeMessage();
                ShowMainMenu();

                var isRunning = true;
                while (isRunning)
                {
                    try
                    {
                        Console.Write("\n请输入命令 (输入 'help' 查看帮助): ");
                        var input = Console.ReadLine()?.Trim();

                        if (string.IsNullOrEmpty(input))
                            continue;

                        if (input.ToLower() == "exit" || input.ToLower() == "quit")
                        {
                            isRunning = false;
                            break;
                        }

                        await ProcessCommandAsync(input);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "配置编辑命令执行失败");
                        Console.WriteLine($"错误: {ex.Message}");
                    }
                }

                Console.WriteLine("配置编辑器已退出");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "配置编辑器运行失败");
                Console.WriteLine($"严重错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示欢迎信息
        /// </summary>
        private void ShowWelcomeMessage()
        {
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("          数据导出工具 - 配置编辑器");
            Console.WriteLine("=".PadRight(60, '='));
            Console.WriteLine("版本: 1.0.0");
            Console.WriteLine("提供交互式的配置文件编辑功能");
            Console.WriteLine("支持导出模式配置、项目配置、全局设置等");
            Console.WriteLine("=".PadRight(60, '='));
        }

        /// <summary>
        /// 显示主菜单
        /// </summary>
        private void ShowMainMenu()
        {
            Console.WriteLine("\n主要功能:");
            Console.WriteLine("  1. 导出模式配置 - 管理各种导出模式");
            Console.WriteLine("  2. 项目配置 - 管理项目信息");
            Console.WriteLine("  3. 全局设置 - 管理系统全局配置");
            Console.WriteLine("  4. 配置验证 - 验证配置文件有效性");
            Console.WriteLine("  5. 配置备份 - 备份和恢复配置");
            Console.WriteLine("  6. 帮助信息 - 查看详细帮助");
            Console.WriteLine("  0. 退出编辑器");
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
                case "help":
                case "h":
                    ShowHelp();
                    break;

                case "mode":
                case "export":
                    await HandleExportModeCommandsAsync(parts);
                    break;

                case "project":
                case "proj":
                    await HandleProjectCommandsAsync(parts);
                    break;

                case "global":
                case "g":
                    await HandleGlobalCommandsAsync(parts);
                    break;

                case "validate":
                case "val":
                    await ValidateConfigurationAsync();
                    break;

                case "backup":
                case "bkp":
                    await HandleBackupCommandsAsync(parts);
                    break;

                case "list":
                case "ls":
                    await ListConfigurationAsync();
                    break;

                case "show":
                    await ShowConfigurationAsync(parts);
                    break;

                case "edit":
                    await EditConfigurationAsync(parts);
                    break;

                case "clear":
                case "cls":
                    Console.Clear();
                    ShowWelcomeMessage();
                    ShowMainMenu();
                    break;

                default:
                    Console.WriteLine($"未知命令: {mainCommand}");
                    Console.WriteLine("输入 'help' 查看可用命令");
                    break;
            }
        }

        /// <summary>
        /// 显示帮助信息
        /// </summary>
        private void ShowHelp()
        {
            Console.WriteLine("\n=== 配置编辑器帮助 ===");
            Console.WriteLine("基本命令:");
            Console.WriteLine("  help, h          - 显示此帮助信息");
            Console.WriteLine("  clear, cls       - 清屏");
            Console.WriteLine("  exit, quit       - 退出编辑器");
            Console.WriteLine("\n配置管理命令:");
            Console.WriteLine("  mode <命令>      - 导出模式配置管理");
            Console.WriteLine("  project <命令>   - 项目配置管理");
            Console.WriteLine("  global <命令>    - 全局设置管理");
            Console.WriteLine("  validate, val    - 验证配置文件");
            Console.WriteLine("  backup <命令>    - 配置备份管理");
            Console.WriteLine("\n查看和编辑命令:");
            Console.WriteLine("  list, ls         - 列出配置信息");
            Console.WriteLine("  show <配置项>    - 显示指定配置");
            Console.WriteLine("  edit <配置项>    - 编辑指定配置");
        }

        #region 导出模式配置管理

        private async Task HandleExportModeCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowExportModeHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "list":
                case "ls":
                    await ListExportModesAsync();
                    break;

                case "show":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: mode show <模式名称>");
                        return;
                    }
                    await ShowExportModeAsync(parts[2]);
                    break;

                case "add":
                    await AddExportModeAsync();
                    break;

                case "edit":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: mode edit <模式名称>");
                        return;
                    }
                    await EditExportModeAsync(parts[2]);
                    break;

                case "delete":
                case "remove":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: mode delete <模式名称>");
                        return;
                    }
                    await DeleteExportModeAsync(parts[2]);
                    break;

                default:
                    Console.WriteLine($"未知的导出模式命令: {subCommand}");
                    ShowExportModeHelp();
                    break;
            }
        }

        private void ShowExportModeHelp()
        {
            Console.WriteLine("\n=== 导出模式配置帮助 ===");
            Console.WriteLine("  mode list, ls           - 列出所有导出模式");
            Console.WriteLine("  mode show <名称>        - 显示指定模式配置");
            Console.WriteLine("  mode add                - 添加新的导出模式");
            Console.WriteLine("  mode edit <名称>        - 编辑指定模式配置");
            Console.WriteLine("  mode delete <名称>      - 删除指定模式");
        }

        private async Task ListExportModesAsync()
        {
            Console.WriteLine("\n=== 导出模式列表 ===");
            try
            {
                var config = await LoadExportModesConfigAsync();
                if (config?.ExportModes == null || !config.ExportModes.Any())
                {
                    Console.WriteLine("没有找到配置的导出模式");
                    return;
                }

                foreach (var mode in config.ExportModes.OrderBy(m => m.Priority))
                {
                    var status = mode.EnableParallel ? "并行" : "串行";
                    Console.WriteLine($"  {mode.Mode,-15} - 优先级: {mode.Priority}, {status}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取导出模式列表失败: {ex.Message}");
            }
        }

        private async Task ShowExportModeAsync(string modeName)
        {
            Console.WriteLine($"\n=== 导出模式配置: {modeName} ===");
            try
            {
                var config = await LoadExportModesConfigAsync();
                var mode = config?.ExportModes?.FirstOrDefault(m => m.Mode.ToString().Equals(modeName, StringComparison.OrdinalIgnoreCase));

                if (mode == null)
                {
                    Console.WriteLine($"未找到导出模式: {modeName}");
                    return;
                }

                Console.WriteLine($"模式名称: {mode.Mode}");
                Console.WriteLine($"优先级: {mode.Priority}");
                Console.WriteLine($"启用并行: {mode.EnableParallel}");
                Console.WriteLine($"最大并行数: {mode.MaxParallelCount}");
                Console.WriteLine($"导出间隔: {mode.ExportInterval}ms");
                Console.WriteLine($"重试次数: {mode.RetryCount}");
                Console.WriteLine($"重试间隔: {mode.RetryInterval}ms");
                Console.WriteLine($"自动合并: {mode.AutoMerge}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取导出模式配置失败: {ex.Message}");
            }
        }

        private async Task AddExportModeAsync()
        {
            Console.WriteLine("\n=== 添加新的导出模式 ===");
            Console.WriteLine("导出模式添加功能待实现");
        }

        private async Task EditExportModeAsync(string modeName)
        {
            Console.WriteLine($"\n=== 编辑导出模式: {modeName} ===");
            Console.WriteLine("导出模式编辑功能待实现");
        }

        private async Task DeleteExportModeAsync(string modeName)
        {
            Console.WriteLine($"\n=== 删除导出模式: {modeName} ===");
            Console.WriteLine("导出模式删除功能待实现");
        }

        #endregion

        #region 项目配置管理

        private async Task HandleProjectCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowProjectHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "list":
                case "ls":
                    await ListProjectsAsync();
                    break;

                case "show":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: project show <项目名称>");
                        return;
                    }
                    await ShowProjectAsync(parts[2]);
                    break;

                case "add":
                    await AddProjectAsync();
                    break;

                case "edit":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: project edit <项目名称>");
                        return;
                    }
                    await EditProjectAsync(parts[2]);
                    break;

                case "delete":
                case "remove":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: project delete <项目名称>");
                        return;
                    }
                    await DeleteProjectAsync(parts[2]);
                    break;

                default:
                    Console.WriteLine($"未知的项目命令: {subCommand}");
                    ShowProjectHelp();
                    break;
            }
        }

        private void ShowProjectHelp()
        {
            Console.WriteLine("\n=== 项目配置管理帮助 ===");
            Console.WriteLine("  project list, ls         - 列出所有项目");
            Console.WriteLine("  project show <名称>      - 显示指定项目配置");
            Console.WriteLine("  project add              - 添加新项目");
            Console.WriteLine("  project edit <名称>      - 编辑指定项目");
            Console.WriteLine("  project delete <名称>    - 删除指定项目");
        }

        private async Task ListProjectsAsync()
        {
            Console.WriteLine("\n=== 项目列表 ===");
            Console.WriteLine("项目列表功能待实现");
        }

        private async Task ShowProjectAsync(string projectName)
        {
            Console.WriteLine($"\n=== 项目配置: {projectName} ===");
            Console.WriteLine("项目配置显示功能待实现");
        }

        private async Task AddProjectAsync()
        {
            Console.WriteLine("\n=== 添加新项目 ===");
            Console.WriteLine("项目添加功能待实现");
        }

        private async Task EditProjectAsync(string projectName)
        {
            Console.WriteLine($"\n=== 编辑项目: {projectName} ===");
            Console.WriteLine("项目编辑功能待实现");
        }

        private async Task DeleteProjectAsync(string projectName)
        {
            Console.WriteLine($"\n=== 删除项目: {projectName} ===");
            Console.WriteLine("项目删除功能待实现");
        }

        #endregion

        #region 全局设置管理

        private async Task HandleGlobalCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowGlobalHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "show":
                    await ShowGlobalSettingsAsync();
                    break;

                case "edit":
                    await EditGlobalSettingsAsync();
                    break;

                case "reset":
                    await ResetGlobalSettingsAsync();
                    break;

                default:
                    Console.WriteLine($"未知的全局设置命令: {subCommand}");
                    ShowGlobalHelp();
                    break;
            }
        }

        private void ShowGlobalHelp()
        {
            Console.WriteLine("\n=== 全局设置管理帮助 ===");
            Console.WriteLine("  global show              - 显示全局设置");
            Console.WriteLine("  global edit              - 编辑全局设置");
            Console.WriteLine("  global reset             - 重置全局设置");
        }

        private async Task ShowGlobalSettingsAsync()
        {
            Console.WriteLine("\n=== 全局设置 ===");
            Console.WriteLine("全局设置显示功能待实现");
        }

        private async Task EditGlobalSettingsAsync()
        {
            Console.WriteLine("\n=== 编辑全局设置 ===");
            Console.WriteLine("全局设置编辑功能待实现");
        }

        private async Task ResetGlobalSettingsAsync()
        {
            Console.WriteLine("\n=== 重置全局设置 ===");
            Console.WriteLine("全局设置重置功能待实现");
        }

        #endregion

        #region 配置验证

        private async Task ValidateConfigurationAsync()
        {
            Console.WriteLine("\n=== 配置验证 ===");
            Console.WriteLine("配置验证功能待实现");
        }

        #endregion

        #region 配置备份管理

        private async Task HandleBackupCommandsAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                ShowBackupHelp();
                return;
            }

            var subCommand = parts[1].ToLower();

            switch (subCommand)
            {
                case "create":
                    await CreateBackupAsync();
                    break;

                case "list":
                case "ls":
                    await ListBackupsAsync();
                    break;

                case "restore":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: backup restore <备份文件名>");
                        return;
                    }
                    await RestoreBackupAsync(parts[2]);
                    break;

                case "delete":
                    if (parts.Length < 3)
                    {
                        Console.WriteLine("用法: backup delete <备份文件名>");
                        return;
                    }
                    await DeleteBackupAsync(parts[2]);
                    break;

                default:
                    Console.WriteLine($"未知的备份命令: {subCommand}");
                    ShowBackupHelp();
                    break;
            }
        }

        private void ShowBackupHelp()
        {
            Console.WriteLine("\n=== 配置备份管理帮助 ===");
            Console.WriteLine("  backup create            - 创建配置备份");
            Console.WriteLine("  backup list, ls          - 列出所有备份");
            Console.WriteLine("  backup restore <文件>    - 从备份恢复");
            Console.WriteLine("  backup delete <文件>     - 删除备份文件");
        }

        private async Task CreateBackupAsync()
        {
            Console.WriteLine("\n=== 创建配置备份 ===");
            Console.WriteLine("配置备份创建功能待实现");
        }

        private async Task ListBackupsAsync()
        {
            Console.WriteLine("\n=== 配置备份列表 ===");
            Console.WriteLine("配置备份列表功能待实现");
        }

        private async Task RestoreBackupAsync(string backupFileName)
        {
            Console.WriteLine($"\n=== 从备份恢复: {backupFileName} ===");
            Console.WriteLine("配置恢复功能待实现");
        }

        private async Task DeleteBackupAsync(string backupFileName)
        {
            Console.WriteLine($"\n=== 删除备份文件: {backupFileName} ===");
            Console.WriteLine("备份文件删除功能待实现");
        }

        #endregion

        #region 配置查看和编辑

        private async Task ListConfigurationAsync()
        {
            Console.WriteLine("\n=== 配置概览 ===");
            Console.WriteLine("配置概览功能待实现");
        }

        private async Task ShowConfigurationAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                Console.WriteLine("用法: show <配置项>");
                Console.WriteLine("可用配置项: modes, projects, global, all");
                return;
            }

            var configItem = parts[1].ToLower();

            switch (configItem)
            {
                case "modes":
                    await ListExportModesAsync();
                    break;

                case "projects":
                    await ListProjectsAsync();
                    break;

                case "global":
                    await ShowGlobalSettingsAsync();
                    break;

                case "all":
                    await ListConfigurationAsync();
                    break;

                default:
                    Console.WriteLine($"未知的配置项: {configItem}");
                    break;
            }
        }

        private async Task EditConfigurationAsync(string[] parts)
        {
            if (parts.Length < 2)
            {
                Console.WriteLine("用法: edit <配置项>");
                Console.WriteLine("可用配置项: modes, projects, global");
                return;
            }

            var configItem = parts[1].ToLower();

            switch (configItem)
            {
                case "modes":
                    Console.WriteLine("使用 'mode add' 添加模式，'mode edit <名称>' 编辑模式");
                    break;

                case "projects":
                    Console.WriteLine("使用 'project add' 添加项目，'project edit <名称>' 编辑项目");
                    break;

                case "global":
                    await EditGlobalSettingsAsync();
                    break;

                default:
                    Console.WriteLine($"未知的配置项: {configItem}");
                    break;
            }
        }

        #endregion

        #region 配置文件加载和保存

        private async Task<ExportModesConfig?> LoadExportModesConfigAsync()
        {
            try
            {
                if (!File.Exists(_exportModesConfigPath))
                {
                    return null;
                }

                var json = await File.ReadAllTextAsync(_exportModesConfigPath);
                return JsonSerializer.Deserialize<ExportModesConfig>(json, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载导出模式配置失败");
                throw;
            }
        }

        private async Task SaveExportModesConfigAsync(ExportModesConfig config)
        {
            try
            {
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                await File.WriteAllTextAsync(_exportModesConfigPath, json);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存导出模式配置失败");
                throw;
            }
        }

        private async Task<ExportConfig?> LoadMainConfigAsync()
        {
            try
            {
                if (!File.Exists(_mainConfigPath))
                {
                    return null;
                }

                var json = await File.ReadAllTextAsync(_mainConfigPath);
                return JsonSerializer.Deserialize<ExportConfig>(json, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载主配置失败");
                throw;
            }
        }

        private async Task SaveMainConfigAsync(ExportConfig config)
        {
            try
            {
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                await File.WriteAllTextAsync(_mainConfigPath, json);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存主配置失败");
                throw;
            }
        }

        #endregion
    }

    /// <summary>
    /// 导出模式配置
    /// </summary>
    public class ExportModesConfig
    {
        public List<ExportModeConfig> ExportModes { get; set; } = new();
        public GlobalExportSettings GlobalSettings { get; set; } = new();
    }
}
