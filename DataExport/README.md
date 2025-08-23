# 数据导出工具

基于基坑监测系统的自动化数据导出工具，支持参数化配置和批量导出功能。

## 功能特性

- ✅ **参数化配置**: 支持项目ID、数据类型、时间范围等灵活配置
- ✅ **HTTP客户端封装**: 完整的HTTP请求封装，支持认证和头信息管理
- ✅ **认证管理**: 基于Cookie的会话认证，支持自动配置
- ✅ **配置验证**: 完整的参数验证和错误提示
- ✅ **错误处理**: 完善的异常处理和重试机制
- ✅ **批量导出**: 支持多项目、多时间段的批量导出
- ✅ **日志记录**: 完整的操作日志和错误记录
- ✅ **配置文件管理**: JSON格式的配置文件，支持热重载

## 快速开始

### 1. 环境要求

- .NET 8.0 或更高版本
- Windows 10/11 或 Linux/macOS

### 2. 安装依赖

```bash
dotnet restore
```

### 3. 配置设置

编辑 `appsettings.json` 文件，配置以下参数：

```json
{
  "Export": {
    "ProjectId": "your-project-id",
    "ProjectCode": "your-project-code",
    "ProjectName": "your-project-name",
    "DataCode": "your-data-code",
    "DataName": "your-data-name",
    "StartTime": "2025-07-01 00:00",
    "EndTime": "2025-07-31 00:00",
    "OutputPath": "./exports"
  },
  "Auth": {
    "BaseUrl": "http://localhost:20472",
    "CookieString": "your-cookie-string-here"
  }
}
```

### 4. 运行程序

```bash
dotnet run
```

## 配置说明

### 导出配置 (Export)

| 参数 | 说明 | 必填 | 默认值 |
|------|------|------|--------|
| ProjectId | 项目唯一标识 | 是 | - |
| ProjectCode | 项目代码 | 是 | - |
| ProjectName | 项目名称 | 否 | - |
| DataCode | 数据类型代码 | 是 | - |
| DataName | 数据类型名称 | 否 | - |
| StartTime | 开始时间 | 是 | 30天前 |
| EndTime | 结束时间 | 是 | 今天 |
| PointCodes | 监测点代码 | 否 | 空 |
| WithDetail | 是否包含详细信息 | 否 | true |
| OutputPath | 输出目录 | 否 | ./exports |
| FileNameFormat | 文件名格式 | 否 | 自动生成 |

### 认证配置 (Auth)

| 参数 | 说明 | 必填 | 默认值 |
|------|------|------|--------|
| BaseUrl | 服务器基础URL | 否 | http://localhost:20472 |
| CookieString | Cookie认证字符串 | 是 | - |
| UserAgent | 用户代理 | 否 | Chrome 139 |
| AcceptLanguage | 接受语言 | 否 | zh-CN,en |
| Accept | 接受类型 | 否 | text/html,application/xml |

## 使用方法

### 基本导出

程序启动后会自动：
1. 验证配置文件
2. 测试网络连接
3. 显示导出参数
4. 确认后开始导出

### 批量导出

支持以下批量导出方式：

```csharp
// 多项目批量导出
var configs = new List<ExportConfig> { /* 配置列表 */ };
var result = await batchService.ExportMultipleProjectsAsync(configs);

// 按时间范围批量导出
var result = await batchService.ExportByTimeRangeAsync(baseConfig, startDate, endDate, 7);

// 按数据类型批量导出
var dataTypes = new List<(string, string)> { ("Type1", "名称1"), ("Type2", "名称2") };
var result = await batchService.ExportByDataTypesAsync(baseConfig, dataTypes);
```

### 文件名格式

支持以下占位符：
- `{ProjectCode}`: 项目代码
- `{DataCode}`: 数据类型代码
- `{StartTime:yyyyMMdd}`: 开始时间（格式：yyyyMMdd）
- `{EndTime:yyyyMMdd}`: 结束时间（格式：yyyyMMdd）

示例：`{ProjectCode}_{DataCode}_{StartTime:yyyyMMdd}_{EndTime:yyyyMMdd}.xlsx`

## 项目结构

```
DataExport/
├── Models/                 # 数据模型
│   ├── ExportConfig.cs    # 导出配置
│   └── AuthConfig.cs      # 认证配置
├── Services/              # 业务服务
│   ├── DataExportService.cs      # 数据导出服务
│   ├── ConfigurationService.cs   # 配置管理服务
│   └── BatchExportService.cs     # 批量导出服务
├── Program.cs             # 主程序
├── appsettings.json       # 配置文件
└── README.md              # 说明文档
```

## 开发说明

### 添加新的数据类型

1. 在 `ExportConfig` 中添加新属性
2. 在 `ConfigurationService` 中添加配置加载逻辑
3. 在 `DataExportService` 中添加相应的处理逻辑

### 扩展导出格式

1. 修改 `DataExportService.ExportDataAsync` 方法
2. 添加新的文件格式支持
3. 更新配置验证逻辑

### 添加新的认证方式

1. 在 `AuthConfig` 中添加新属性
2. 在 `DataExportService.ConfigureHttpClient` 中添加配置逻辑
3. 更新认证验证方法

## 故障排除

### 常见问题

1. **配置验证失败**
   - 检查 `appsettings.json` 文件格式
   - 确保必填参数已填写
   - 验证时间格式是否正确

2. **认证失败**
   - 检查Cookie是否有效
   - 确认服务器地址是否正确
   - 验证网络连接

3. **导出失败**
   - 检查服务器状态
   - 确认项目ID和数据类型是否正确
   - 查看日志获取详细错误信息

### 日志查看

程序运行时会输出详细的日志信息，包括：
- 配置加载状态
- 网络连接测试结果
- 导出进度和结果
- 错误详情和堆栈信息

## 许可证

本项目采用 MIT 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个工具。

---

*最后更新: 2024-12-19*
