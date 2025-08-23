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
                    .Build();

                // 创建服务容器
                var services = new ServiceCollection();
                
                // 添加配置
                var exportConfig = new ExportConfig();
                configuration.Bind(exportConfig);
                services.AddSingleton(exportConfig);

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

                var serviceProvider = services.BuildServiceProvider();

                // 获取服务
                var logger = serviceProvider.GetRequiredService<ILogger<Program>>();
                var batchExportService = serviceProvider.GetRequiredService<BatchExportService>();

                // 显示配置信息
                DisplayConfiguration(exportConfig);

                Console.WriteLine();
                Console.WriteLine("选择操作模式:");
                Console.WriteLine("1. 批量导出所有项目 (不合并)");
                Console.WriteLine("2. 批量导出所有项目 + Excel合并 (推荐)");
                Console.WriteLine("3. 导出单个项目");
                Console.WriteLine("4. 退出");
                Console.Write("请输入选择 (1-4): ");

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
                        Console.WriteLine("单个项目导出功能待实现...");
                        break;

                    case "4":
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
