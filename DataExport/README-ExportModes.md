# 数据导出工具 - 导出模式功能说明

## 概述

数据导出工具现在支持多种导出模式，可以根据不同的业务需求选择合适的导出方式。每种模式都有独立的配置和优化选项。

## 支持的导出模式

### 1. 导出所有项目 (AllProjects)
**默认模式**，适用于需要导出所有配置项目的场景。

**特点：**
- 自动遍历所有配置的项目
- 使用配置的月度时间范围
- 支持自动Excel合并
- 适合定期全量数据导出

**配置示例：**
```json
{
  "Mode": "AllProjects",
  "AutoMerge": true,
  "Priority": 1,
  "EnableParallel": false,
  "ExportInterval": 1000
}
```

### 2. 单个项目导出 (SingleProject)
适用于只需要导出特定项目数据的场景。

**特点：**
- 指定单个项目ID和名称
- 可选择特定的数据类型
- 支持自定义时间范围
- 适合项目级别的数据导出

**配置示例：**
```json
{
  "Mode": "SingleProject",
  "SingleProject": {
    "ProjectId": "28fd93ef-7b19-4886-98dd-dadc79deed03",
    "ProjectName": "基坑监测项目A",
    "DataTypes": ["QC001", "QC002"],
    "TimeRange": {
      "StartTime": "2025-01-01 00:00",
      "EndTime": "2025-01-31 23:59"
    }
  },
  "AutoMerge": true,
  "Priority": 2
}
```

### 3. 自定义时间范围导出 (CustomTimeRange)
适用于需要导出特定时间范围数据的场景。

**特点：**
- 指定时间范围
- 可选择多个项目
- 可选择特定数据类型
- 适合跨项目的特定时间段导出

**配置示例：**
```json
{
  "Mode": "CustomTimeRange",
  "CustomTimeRange": {
    "ProjectIds": ["project-id-1", "project-id-2"],
    "DataTypeCodes": ["QC001", "QC002", "QC003"],
    "TimeRange": {
      "StartTime": "2025-02-01 00:00",
      "EndTime": "2025-02-28 23:59"
    },
    "Description": "2025年2月数据导出"
  },
  "EnableParallel": true,
  "MaxParallelCount": 5
}
```

### 4. 批量导出 (BatchExport)
适用于需要批量处理多个项目、多个时间范围的场景。

**特点：**
- 支持多个项目批量处理
- 支持多个时间范围
- 可配置批量大小
- 支持并行导出优化
- 适合大规模数据导出任务

**配置示例：**
```json
{
  "Mode": "BatchExport",
  "BatchExport": {
    "ProjectIds": ["project-id-1", "project-id-2", "project-id-3"],
    "DataTypeCodes": ["QC001", "QC002"],
    "TimeRanges": [
      {
        "StartTime": "2025-03-01 00:00",
        "EndTime": "2025-03-31 23:59"
      },
      {
        "StartTime": "2025-04-01 00:00",
        "EndTime": "2025-04-30 23:59"
      }
    ],
    "BatchSize": 10,
    "GroupByProject": true,
    "GroupByDataType": false
  },
  "EnableParallel": true,
  "MaxParallelCount": 8,
  "ExportInterval": 600
}
```

### 5. 增量导出 (IncrementalExport)
适用于需要定期同步最新数据的场景。

**特点：**
- 基于上次导出时间自动计算增量范围
- 可配置增量时间间隔
- 支持最大增量范围限制
- 可检查数据完整性
- 适合实时数据同步

**配置示例：**
```json
{
  "Mode": "IncrementalExport",
  "IncrementalExport": {
    "LastExportTime": "2025-01-31T23:59:59",
    "IncrementHours": 24,
    "AutoUpdateLastExportTime": true,
    "MaxIncrementHours": 168,
    "CheckDataIntegrity": true
  },
  "EnableParallel": true,
  "MaxParallelCount": 5,
  "ExportInterval": 500
}
```

## 通用配置选项

每种导出模式都支持以下通用配置：

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `AutoMerge` | 是否自动合并Excel文件 | `true` |
| `Priority` | 执行优先级（数字越小优先级越高） | `1` |
| `EnableParallel` | 是否启用并行导出 | `false` |
| `MaxParallelCount` | 最大并行数量 | `3` |
| `ExportInterval` | 导出间隔（毫秒） | `1000` |
| `RetryCount` | 重试次数 | `3` |
| `RetryInterval` | 重试间隔（毫秒） | `5000` |

## 全局设置

支持全局配置来管理多个导出模式：

```json
{
  "GlobalSettings": {
    "DefaultExportMode": "AllProjects",
    "EnableModeSwitching": true,
    "ModeExecutionOrder": "Priority",
    "MaxConcurrentModes": 2,
    "GlobalRetryCount": 3,
    "GlobalRetryInterval": 10000,
    "EnableProgressTracking": true,
    "EnableResultPersistence": true,
    "ResultStoragePath": "./export-results",
    "StopOnFirstFailure": false
  }
}
```

## 使用方法

### 1. 配置导出模式
在 `appsettings.export-modes.json` 中配置所需的导出模式。

### 2. 执行导出模式
```csharp
// 执行所有模式
var results = await exportModeManager.ExecuteAllModesAsync();

// 执行指定模式
var result = await exportModeManager.ExecuteModeAsync("BatchExport");

// 执行默认模式
var result = await exportModeManager.ExecuteDefaultModeAsync();
```

### 3. 监控导出进度
每种导出模式都会返回详细的执行结果，包括：
- 成功/失败数量统计
- 详细的导出结果列表
- Excel合并结果
- 执行时间和耗时统计

## 性能优化建议

1. **并行导出**：对于大量数据导出，建议启用并行导出并调整并行数量
2. **导出间隔**：根据API限制调整导出间隔，避免请求过于频繁
3. **批量大小**：合理设置批量大小，平衡内存使用和性能
4. **优先级管理**：使用优先级控制导出模式的执行顺序
5. **重试机制**：配置合适的重试次数和间隔，提高导出成功率

## 错误处理

- 每种导出模式都有完善的异常处理
- 支持重试机制
- 详细的错误日志记录
- 可配置的失败处理策略

## 扩展性

导出模式系统设计为可扩展的架构：
- 可以轻松添加新的导出模式
- 支持自定义配置选项
- 支持插件式的导出策略
- 支持动态配置加载
