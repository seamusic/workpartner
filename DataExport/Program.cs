using DataExport.Models;
using DataExport.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DataExport
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("=== 数据导出工具 ===");
            Console.WriteLine();

            try
            {
                // 构建配置
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .AddJsonFile("appsettings.export-modes.json", optional: true, reloadOnChange: true)
                    .Build();

                // 创建服务容器
                var services = new ServiceCollection();
                
                // 添加配置
                var exportConfig = new ExportConfig();
                configuration.Bind(exportConfig);
                services.AddSingleton(exportConfig);

                // 读取导出模式配置
                var exportModeConfig = new ExportModeConfigList();
                configuration.Bind("ExportModes", exportModeConfig);
                services.AddSingleton(exportModeConfig.ExportModes);

                // 读取全局导出设置
                var globalSettings = new GlobalExportSettings();
                configuration.Bind("GlobalSettings", globalSettings);
                services.AddSingleton(globalSettings);

                // 添加日志
                services.AddLogging(builder =>
                {
                    builder.AddConsole();
                    builder.AddDebug();
                    builder.SetMinimumLevel(LogLevel.Information);
                });

                // 添加HTTP客户端
                services.AddHttpClient();
                
                // 添加服务
                services.AddSingleton<DataExportService>();
                services.AddSingleton<ExcelMergeService>();
                services.AddSingleton<BatchExportService>();
                services.AddSingleton<ExportModeService>();
                services.AddSingleton<ExportModeManager>();
                services.AddSingleton<ExportHistoryService>();
                services.AddSingleton<DataQualityService>();
                services.AddSingleton<ExportFileManager>();
                services.AddSingleton<ConfigurationEditor>();
                services.AddSingleton<CommandLineInterface>();

                var serviceProvider = services.BuildServiceProvider();

                // 获取服务
                var logger = serviceProvider.GetRequiredService<ILogger<Program>>();
                var batchExportService = serviceProvider.GetRequiredService<BatchExportService>();
                var exportModeManager = serviceProvider.GetRequiredService<ExportModeManager>();
                var dataExportService = serviceProvider.GetRequiredService<DataExportService>();
                var cli = serviceProvider.GetRequiredService<CommandLineInterface>();

                // 显示配置信息
                DisplayConfiguration(exportConfig);

                Console.WriteLine();
                Console.WriteLine("选择操作模式:");
                Console.WriteLine("1. 批量导出所有项目 (不合并)");
                Console.WriteLine("2. 批量导出所有项目 + Excel合并 (推荐)");
                Console.WriteLine("3. 导出单个项目");
                Console.WriteLine("4. 命令行交互模式");
                Console.WriteLine("5. 退出");
                Console.Write("请输入选择 (1-5): ");

                var choice = Console.ReadLine()?.Trim();

                switch (choice)
                {
                    case "1":
                        Console.WriteLine();
                        Console.WriteLine("开始执行批量导出 (不合并Excel文件)...");
                        await batchExportService.ExecuteBatchExportAsync(false);
                        break;

                    case "2":
                        Console.WriteLine();
                        Console.WriteLine("开始执行批量导出 + Excel合并...");
                        Console.WriteLine("注意: Excel合并功能会将同一项目同一数据类型的所有月度文件合并为一个文件");
                        Console.WriteLine();
                        await batchExportService.ExecuteBatchExportAsync(true);
                        break;

                    case "3":
                        Console.WriteLine();
                        await ExportSingleProjectAsync(exportConfig, dataExportService);
                        break;

                    case "4":
                        Console.WriteLine();
                        Console.WriteLine("启动命令行交互模式...");
                        await cli.RunAsync();
                        break;

                    case "5":
                        Console.WriteLine("退出程序");
                        return;

                    default:
                        Console.WriteLine("无效选择，退出程序");
                        return;
                }

                Console.WriteLine();
                Console.WriteLine("程序执行完成，按任意键退出...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序执行失败: {ex.Message}");
                Console.WriteLine($"详细错误: {ex}");
                Console.WriteLine();
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
            }
        }

        /// <summary>
        /// 导出单个项目
        /// </summary>
        private static async Task ExportSingleProjectAsync(ExportConfig config, DataExportService dataExportService)
        {
            try
            {
                if (!config.Projects.Any())
                {
                    Console.WriteLine("配置中没有找到项目信息");
                    return;
                }

                // 显示项目列表
                Console.WriteLine("可用的项目:");
                for (int i = 0; i < config.Projects.Count; i++)
                {
                    var project = config.Projects[i];
                    Console.WriteLine($"{i + 1}. {project.ProjectName} (ID: {project.ProjectId})");
                    Console.WriteLine($"   数据类型: {project.DataTypes.Count} 个");
                    foreach (var dataType in project.DataTypes)
                    {
                        Console.WriteLine($"     - {dataType.DataName} ({dataType.DataCode})");
                    }
                    Console.WriteLine();
                }

                // 选择项目
                Console.Write("请选择项目编号 (1-" + config.Projects.Count + "): ");
                if (!int.TryParse(Console.ReadLine(), out int projectIndex) || 
                    projectIndex < 1 || projectIndex > config.Projects.Count)
                {
                    Console.WriteLine("无效的项目编号");
                    return;
                }

                var selectedProject = config.Projects[projectIndex - 1];

                // 选择数据类型
                Console.WriteLine($"\n项目 '{selectedProject.ProjectName}' 的数据类型:");
                for (int i = 0; i < selectedProject.DataTypes.Count; i++)
                {
                    var dataType = selectedProject.DataTypes[i];
                    Console.WriteLine($"{i + 1}. {dataType.DataName} ({dataType.DataCode})");
                }

                Console.Write("请选择数据类型编号 (1-" + selectedProject.DataTypes.Count + ")，或输入 'all' 选择所有: ");
                var dataTypeChoice = Console.ReadLine()?.Trim().ToLower();

                List<DataTypeConfig> selectedDataTypes;
                if (dataTypeChoice == "all")
                {
                    selectedDataTypes = selectedProject.DataTypes.ToList();
                    Console.WriteLine("已选择所有数据类型");
                }
                else if (int.TryParse(dataTypeChoice, out int dataTypeIndex) && 
                         dataTypeIndex >= 1 && dataTypeIndex <= selectedProject.DataTypes.Count)
                {
                    selectedDataTypes = new List<DataTypeConfig> { selectedProject.DataTypes[dataTypeIndex - 1] };
                    Console.WriteLine($"已选择数据类型: {selectedDataTypes[0].DataName}");
                }
                else
                {
                    Console.WriteLine("无效的数据类型选择");
                    return;
                }

                // 选择时间范围
                Console.WriteLine($"\n可用的时间范围:");
                for (int i = 0; i < config.ExportSettings.MonthlyExport.Months.Count; i++)
                {
                    var month = config.ExportSettings.MonthlyExport.Months[i];
                    Console.WriteLine($"{i + 1}. {month.Name}: {month.StartTime:yyyy-MM-dd} 至 {month.EndTime:yyyy-MM-dd}");
                }

                Console.Write("请选择时间范围编号 (1-" + config.ExportSettings.MonthlyExport.Months.Count + ")，或输入 'all' 选择所有: ");
                var timeChoice = Console.ReadLine()?.Trim().ToLower();

                List<MonthConfig> selectedTimeRanges;
                if (timeChoice == "all")
                {
                    selectedTimeRanges = config.ExportSettings.MonthlyExport.Months.ToList();
                    Console.WriteLine("已选择所有时间范围");
                }
                else if (int.TryParse(timeChoice, out int timeIndex) && 
                         timeIndex >= 1 && timeIndex <= config.ExportSettings.MonthlyExport.Months.Count)
                {
                    selectedTimeRanges = new List<MonthConfig> { config.ExportSettings.MonthlyExport.Months[timeIndex - 1] };
                    Console.WriteLine($"已选择时间范围: {selectedTimeRanges[0].Name}");
                }
                else
                {
                    Console.WriteLine("无效的时间范围选择");
                    return;
                }

                // 确认导出
                Console.WriteLine($"\n导出配置确认:");
                Console.WriteLine($"项目: {selectedProject.ProjectName}");
                Console.WriteLine($"数据类型: {selectedDataTypes.Count} 个");
                Console.WriteLine($"时间范围: {selectedTimeRanges.Count} 个");
                Console.WriteLine($"预计导出文件数: {selectedDataTypes.Count * selectedTimeRanges.Count}");
                Console.WriteLine($"输出目录: {config.ExportSettings.OutputDirectory}");

                Console.Write("\n确认开始导出? (y/n): ");
                var confirm = Console.ReadLine()?.Trim().ToLower();
                if (confirm != "y" && confirm != "yes")
                {
                    Console.WriteLine("导出已取消");
                    return;
                }

                // 开始导出
                Console.WriteLine("\n开始导出单个项目...");
                var startTime = DateTime.Now;

                var exportResults = new List<ExportResult>();
                var totalCount = selectedDataTypes.Count * selectedTimeRanges.Count;
                var currentCount = 0;

                foreach (var dataType in selectedDataTypes)
                {
                    foreach (var timeRange in selectedTimeRanges)
                    {
                        currentCount++;
                        Console.WriteLine($"导出进度: {currentCount}/{totalCount} - {selectedProject.ProjectName} - {dataType.DataName} - {timeRange.Name}");

                        try
                        {
                            var exportParams = new ExportParameters
                            {
                                ProjectId = selectedProject.ProjectId,
                                ProjectName = selectedProject.ProjectName,
                                DataCode = dataType.DataCode,
                                DataName = dataType.DataName,
                                StartTime = timeRange.StartTime,
                                EndTime = timeRange.EndTime,
                                WithDetail = 1,
                                PointCodes = string.Empty
                            };

                            var result = await dataExportService.ExportDataAsync(exportParams);

                            exportResults.Add(result);

                            if (result.Success)
                            {
                                Console.WriteLine($"  ✓ 导出成功: {result.FileName}");
                            }
                            else
                            {
                                Console.WriteLine($"  ✗ 导出失败: {result.ErrorMessage}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  ✗ 导出异常: {ex.Message}");
                            exportResults.Add(new ExportResult
                            {
                                Success = false,
                                ErrorMessage = ex.Message,
                                ProjectName = selectedProject.ProjectName,
                                DataName = dataType.DataName
                            });
                        }
                    }
                }

                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                var successCount = exportResults.Count(r => r.Success);
                var failedCount = exportResults.Count(r => !r.Success);

                Console.WriteLine($"\n导出完成!");
                Console.WriteLine($"总耗时: {duration.TotalSeconds:F1} 秒");
                Console.WriteLine($"成功: {successCount} 个文件");
                Console.WriteLine($"失败: {failedCount} 个文件");
                Console.WriteLine($"成功率: {(double)successCount / totalCount * 100:F1}%");

                if (successCount > 0)
                {
                    Console.WriteLine($"\n成功导出的文件:");
                    foreach (var result in exportResults.Where(r => r.Success))
                    {
                        Console.WriteLine($"  {result.FileName}");
                    }
                }

                if (failedCount > 0)
                {
                    Console.WriteLine($"\n失败的文件:");
                    foreach (var result in exportResults.Where(r => !r.Success))
                    {
                        Console.WriteLine($"  {result.ProjectName} - {result.DataName}: {result.ErrorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"单个项目导出失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示配置信息
        /// </summary>
        private static void DisplayConfiguration(ExportConfig config)
        {
            Console.WriteLine("当前配置:");
            Console.WriteLine($"API地址: {config.ApiSettings.BaseUrl}{config.ApiSettings.Endpoint}");
            Console.WriteLine($"输出目录: {config.ExportSettings.OutputDirectory}");
            Console.WriteLine($"月度导出: {(config.ExportSettings.MonthlyExport.Enabled ? "启用" : "禁用")}");
            Console.WriteLine($"项目数量: {config.Projects.Count}");
            Console.WriteLine($"月度数量: {config.ExportSettings.MonthlyExport.Months.Count}");
            
            var totalExports = config.Projects.Count * 
                             config.ExportSettings.MonthlyExport.Months.Count * 
                             config.Projects.Sum(p => p.DataTypes.Count);
            Console.WriteLine($"预计导出文件数: {totalExports}");
            
            Console.WriteLine();
            Console.WriteLine("项目列表:");
            foreach (var project in config.Projects)
            {
                Console.WriteLine($"  {project.ProjectName} (ID: {project.ProjectId})");
                Console.WriteLine($"    数据类型: {project.DataTypes.Count} 个");
                foreach (var dataType in project.DataTypes)
                {
                    Console.WriteLine($"      - {dataType.DataName} ({dataType.DataCode})");
                }
            }

            Console.WriteLine();
            Console.WriteLine("月度配置:");
            foreach (var month in config.ExportSettings.MonthlyExport.Months)
            {
                Console.WriteLine($"  {month.Name}: {month.StartTime} 至 {month.EndTime}");
            }
        }
    }
}
