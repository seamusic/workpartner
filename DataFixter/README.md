# DataFixter - Excel数据修正工具

## 项目简介

DataFixter是一个基于.NET 8.0开发的Excel数据修正工具，专门用于处理Excel文件中的数据完整性问题。该工具提供了数据验证、缺失数据修正和数据统计分析等功能。

## 主要功能

- **数据验证**: 检查Excel文件中指定列的数据完整性
- **数据修正**: 自动填充缺失数据，支持自定义默认值
- **统计分析**: 生成详细的数据统计报告
- **日志记录**: 完整的操作日志记录，支持控制台和文件输出
- **批量处理**: 支持大批量数据的处理

## 技术特性

- 基于.NET 8.0开发
- 使用NPOI库处理Excel文件
- 集成Serilog日志系统
- 支持.xlsx和.xls格式
- 内存优化的数据处理
- 完整的错误处理和异常管理

## 系统要求

- Windows 10/11 或 Windows Server 2016+
- .NET 8.0 Runtime
- 至少100MB可用磁盘空间

## 安装说明

### 1. 克隆项目

```bash
git clone <repository-url>
cd WorkPartner/DataFixter
```

### 2. 还原NuGet包

```bash
dotnet restore
```

### 3. 编译项目

```bash
dotnet build
```

### 4. 发布项目（可选）

```bash
dotnet publish -c Release -o ./publish
```

## 使用方法

### 基本语法

```bash
DataFixter <命令> <文件路径> <开始行> <结束行>
```

### 可用命令

#### 1. 数据验证 (validate)

验证Excel文件中指定范围的数据完整性：

```bash
DataFixter validate data.xlsx 1 100
```

这将验证data.xlsx文件中第1行到第100行的数据完整性。

#### 2. 数据修正 (correct)

自动修正Excel文件中的缺失数据：

```bash
DataFixter correct data.xlsx 1 100
```

这将检查并修正data.xlsx文件中第1行到第100行的缺失数据。

#### 3. 数据统计 (stats)

生成Excel文件的数据统计报告：

```bash
DataFixter stats data.xlsx 1 100
```

这将生成data.xlsx文件中第1行到第100行的详细统计报告。

### 参数说明

- **文件路径**: Excel文件的完整路径
- **开始行**: 开始处理的行索引（从0开始）
- **结束行**: 结束处理的行索引（包含）

### 注意事项

- 行索引从0开始计算（即第1行在Excel中对应索引0）
- 工具会自动验证A列和B列（索引0和1）的数据完整性
- 修正后的文件会保存为原文件名加上"_corrected"后缀
- 所有操作都会记录详细的日志信息

## 配置说明

### 日志配置

在`appsettings.json`中可以配置日志相关参数：

```json
{
  "Logging": {
    "LogFilePath": "logs/DataFixter.log",
    "ConsoleOutput": true,
    "FileOutput": true,
    "RetentionDays": 30,
    "MaxFileSizeMB": 10
  }
}
```

### 数据处理配置

```json
{
  "DataProcessing": {
    "DefaultValidationColumns": [0, 1],
    "DefaultValues": {
      "0": "默认名称",
      "1": "0"
    },
    "BatchSize": 1000
  }
}
```

## 输出文件

### 日志文件

- 位置: `logs/DataFixter.log`
- 格式: 按日期滚动的日志文件
- 保留: 默认保留30天

### 修正后的文件

- 命名规则: `原文件名_corrected.xlsx`
- 位置: 与原文件相同目录
- 备份: 自动备份原文件（如果启用）

## 错误处理

工具提供了完善的错误处理机制：

- 文件不存在或无法访问
- Excel文件格式错误
- 数据验证失败
- 内存不足等异常情况

所有错误都会记录到日志文件中，并在控制台显示相应的错误信息。

## 性能优化

- 批量处理数据，减少内存占用
- 流式读取Excel文件，支持大文件处理
- 异步操作支持（计划中）
- 内存使用监控和优化

## 开发说明

### 项目结构

```
DataFixter/
├── Excel/                 # Excel处理相关类
│   └── ExcelProcessor.cs
├── Logging/              # 日志配置类
│   └── LoggingConfiguration.cs
├── Services/             # 业务逻辑服务
│   └── DataProcessingService.cs
├── Program.cs            # 主程序入口
├── appsettings.json      # 配置文件
└── README.md             # 项目说明文档
```

### 扩展开发

工具采用模块化设计，可以轻松扩展新功能：

1. 在`Services`目录下添加新的服务类
2. 在`Program.cs`中添加新的命令处理
3. 修改配置文件以支持新功能

## 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 贡献指南

欢迎提交Issue和Pull Request来改进这个工具。

## 联系方式

如有问题或建议，请通过以下方式联系：

- 提交GitHub Issue
- 发送邮件至：[your-email@example.com]

## 更新日志

### v1.0.0 (2024-01-XX)
- 初始版本发布
- 支持基本的Excel数据验证和修正功能
- 集成Serilog日志系统
- 支持命令行操作
