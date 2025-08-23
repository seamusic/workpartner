using DataExport.Models;
using DataExport.Services;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using System.Text.Json;

namespace DataExport
{
    /// <summary>
    /// 导出模式管理器测试程序
    /// </summary>
    public class TestExportModeManager
    {
        private readonly ExportModeManager _exportModeManager;
        private readonly ILogger<TestExportModeManager> _logger;

        public TestExportModeManager(ExportModeManager exportModeManager, ILogger<TestExportModeManager> logger)
        {
            _exportModeManager = exportModeManager;
            _logger = logger;
        }

        /// <summary>
        /// 测试导出模式管理器
        /// </summary>
        public async Task RunTestAsync()
        {
            _logger.LogInformation("开始测试导出模式管理器...");

            try
            {
                // 1. 获取所有导出模式
                var modes = _exportModeManager.GetExportModes();
                _logger.LogInformation($"找到 {modes.Count} 个导出模式:");
                foreach (var mode in modes)
                {
                    _logger.LogInformation($"  - {mode.Mode}: {mode.Description} (优先级: {mode.Priority})");
                }

                // 2. 验证导出模式配置
                var validationResults = _exportModeManager.ValidateModes();
                if (validationResults.Any())
                {
                    _logger.LogWarning("发现配置验证问题:");
                    foreach (var result in validationResults)
                    {
                        _logger.LogWarning($"  - {result.Message}");
                    }
                }
                else
                {
                    _logger.LogInformation("所有导出模式配置验证通过");
                }

                // 3. 测试执行默认模式
                _logger.LogInformation("测试执行默认导出模式...");
                try
                {
                    var defaultResult = await _exportModeManager.ExecuteDefaultModeAsync();
                    _logger.LogInformation($"默认模式执行结果: {defaultResult.GetSummary()}");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"默认模式执行失败: {ex.Message}");
                }

                // 4. 测试执行指定模式
                _logger.LogInformation("测试执行指定导出模式...");
                try
                {
                    var singleProjectResult = await _exportModeManager.ExecuteModeAsync("SingleProject");
                    _logger.LogInformation($"单个项目导出模式执行结果: {singleProjectResult.GetSummary()}");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"单个项目导出模式执行失败: {ex.Message}");
                }

                // 5. 测试执行所有模式
                _logger.LogInformation("测试执行所有导出模式...");
                try
                {
                    var allResults = await _exportModeManager.ExecuteAllModesAsync();
                    _logger.LogInformation($"所有模式执行完成，总计 {allResults.Count} 个模式:");
                    foreach (var result in allResults)
                    {
                        _logger.LogInformation($"  - {result.Mode}: {result.GetSummary()}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"执行所有模式失败: {ex.Message}");
                }

                _logger.LogInformation("导出模式管理器测试完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "测试过程中发生错误");
            }
        }
    }


}
