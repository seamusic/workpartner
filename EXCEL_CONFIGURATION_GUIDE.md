# Excel配置管理使用指南

## 概述

`ExcelConfiguration` 类用于管理Excel文件读取的各种配置参数，解决了代码中魔法数字的问题，提供了灵活的配置管理方案。

## 主要功能

### 1. 配置管理
- 从JSON配置文件读取配置
- 支持动态配置修改
- 自动保存配置到文件
- 配置验证和重置功能

### 2. 扩展性支持
- 支持从Excel文件动态读取配置
- 自动分析Excel工作表结构
- 支持不同Excel格式的配置

### 3. 单例模式
- 全局统一的配置管理
- 线程安全的配置访问

## 配置参数说明

| 参数 | 默认值 | 说明 |
|------|--------|------|
| StartRow | 5 | 数据开始行号 |
| EndRow | 368 | 数据结束行号 |
| StartCol | 4 | 数据开始列号（D列） |
| EndCol | 9 | 数据结束列号（I列） |
| NameCol | 2 | 名称列号（B列） |

## 计算属性

| 属性 | 说明 |
|------|------|
| TotalRows | 总行数（EndRow - StartRow + 1） |
| TotalCols | 总列数（EndCol - StartCol + 1） |
| NpoiStartRow | NPOI库的0基索引开始行 |
| NpoiEndRow | NPOI库的0基索引结束行 |
| NpoiStartCol | NPOI库的0基索引开始列 |
| NpoiEndCol | NPOI库的0基索引结束列 |
| NpoiNameCol | NPOI库的0基索引名称列 |

## 使用方法

### 基本使用

```csharp
// 获取配置实例
var config = ExcelConfiguration.Instance;

// 读取配置参数
int startRow = config.StartRow;
int endRow = config.EndRow;

// 修改配置
config.StartRow = 10;
config.EndRow = 100;

// 保存配置
config.SaveConfiguration();

// 重置为默认配置
config.ResetToDefault();
```

### 在ExcelService中使用

```csharp
// 获取配置
var config = ExcelConfiguration.Instance;

// 使用配置读取Excel数据
for (int row = config.StartRow; row <= config.EndRow; row++)
{
    var nameCell = worksheet.Cells[row, config.NameCol];
    // ... 处理数据
}
```

### 动态配置读取

```csharp
// 从Excel文件读取配置
var config = ExcelConfiguration.Instance;
if (config.LoadConfigurationFromExcel("path/to/excel.xlsx"))
{
    Console.WriteLine("成功从Excel文件读取配置");
}
```

## 配置文件格式

配置文件 `excel_config.json` 位于应用程序根目录：

```json
{
  "StartRow": 5,
  "EndRow": 368,
  "StartCol": 4,
  "EndCol": 9,
  "NameCol": 2
}
```

## 配置验证

配置类会自动验证配置的有效性：

- StartRow > 0
- EndRow >= StartRow
- StartCol > 0
- EndCol >= StartCol
- NameCol > 0

如果配置无效，会自动重置为默认配置。

## 扩展性设计

### 1. 支持不同Excel格式

可以通过修改配置来支持不同的Excel格式：

```csharp
// 标准格式
config.StartRow = 5; config.EndRow = 368;

// 紧凑格式
config.StartRow = 3; config.EndRow = 200;

// 扩展格式
config.StartRow = 10; config.EndRow = 500;
```

### 2. 动态配置读取

支持从Excel文件自动分析并读取配置：

```csharp
// 自动分析Excel文件结构
config.LoadConfigurationFromExcel(filePath);
```

### 3. 配置文件热更新

支持运行时修改配置文件，程序会自动重新加载：

```csharp
// 修改配置文件后，重新加载
config.LoadConfiguration();
```

## 最佳实践

### 1. 配置管理
- 使用配置文件而不是硬编码
- 定期备份配置文件
- 在部署前测试配置

### 2. 错误处理
- 检查配置的有效性
- 提供默认配置作为备选
- 记录配置加载和修改日志

### 3. 性能优化
- 使用单例模式避免重复加载
- 缓存配置值减少文件IO
- 异步加载配置文件

## 示例代码

参考 `WorkPartner/Examples/ExcelConfigurationExample.cs` 文件中的完整示例。

## 注意事项

1. 配置文件路径：`AppDomain.CurrentDomain.BaseDirectory/excel_config.json`
2. 配置修改后需要调用 `SaveConfiguration()` 保存
3. NPOI库使用0基索引，EPPlus使用1基索引
4. 动态配置读取功能需要Excel文件可访问
5. 配置验证失败时会自动重置为默认值

## 更新日志

- v1.0: 初始版本，支持基本配置管理
- v1.1: 添加动态配置读取功能
- v1.2: 增强配置验证和错误处理 