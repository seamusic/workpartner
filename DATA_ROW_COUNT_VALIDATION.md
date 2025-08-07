# 数据行数验证功能说明

## 功能概述

在Excel数据读取过程中，增加了数据行数长短不一致的判断功能。当读取的Excel文件数据行数与预期不符时，系统会提示用户是否继续处理。

## 功能特性

### 1. 自动检测数据行数
- **预期行数**: 364行 (B5-B368行)
- **检测范围**: 读取B列中非空的数据名称行
- **支持格式**: .xls 和 .xlsx 文件

### 2. 用户交互提示
当检测到数据行数不一致时，系统会显示：
```
警告：文件 [文件名] 的数据行数不一致。
预期行数：364，实际读取行数：[实际行数]
是否继续处理？(Y/N):
```

### 3. 用户选择处理
- **输入 Y 或 YES**: 继续处理文件
- **输入其他内容**: 取消处理，抛出 `WorkPartnerException` 异常

### 4. 日志记录
- 记录警告信息到日志文件
- 记录用户的选择和处理结果

## 实现细节

### 代码位置
- **主要实现**: `WorkPartner/Services/ExcelService.cs`
- **验证逻辑**: `ReadExcelFile` 方法中的行数检查部分

### 核心代码逻辑
```csharp
// 检查数据行数长短不一致的情况
if (dataRows.Count > 0)
{
    var expectedRowCount = 364; // B5-B368行，应该是364行数据
    if (dataRows.Count != expectedRowCount)
    {
        var message = $"警告：文件 {Path.GetFileName(filePath)} 的数据行数不一致。\n" +
                     $"预期行数：{expectedRowCount}，实际读取行数：{dataRows.Count}\n" +
                     $"是否继续处理？(Y/N): ";
        
        Console.Write(message);
        var response = Console.ReadLine()?.Trim().ToUpper();
        
        if (response != "Y" && response != "YES")
        {
            throw new WorkPartnerException("UserCancelled", "用户取消了处理操作", filePath);
        }
        
        Logger.Warning($"数据行数不一致：预期{expectedRowCount}行，实际{dataRows.Count}行，用户选择继续处理");
    }
}
```

## 使用场景

### 1. 正常情况
- 文件包含完整的B5-B368行数据
- 系统正常读取364行数据
- 无需用户干预

### 2. 数据缺失情况
- 某些行被删除或为空
- 实际读取行数少于364行
- 用户可选择继续或取消

### 3. 数据冗余情况
- 文件中包含额外的数据行
- 实际读取行数多于364行
- 用户可选择继续或取消

## 测试用例

### 单元测试
- `DataRowCountValidationTests.cs`: 验证行数检测逻辑
- 测试预期行数计算
- 测试异常处理机制

### 示例代码
- `DataRowCountValidationExample.cs`: 演示功能使用
- 单个文件验证
- 批量文件验证

## 错误处理

### 异常类型
- `WorkPartnerException`: 用户取消处理时抛出
- 异常类别: "UserCancelled"
- 包含文件路径信息

### 日志记录
- 警告级别日志记录
- 记录预期行数和实际行数
- 记录用户选择结果

## 配置选项

### 预期行数
- 当前固定为364行 (B5-B368)
- 可根据需要修改 `expectedRowCount` 变量

### 用户提示
- 支持中文提示信息
- 支持 Y/N 和 YES/NO 输入
- 大小写不敏感

## 注意事项

1. **交互性**: 此功能需要控制台交互，不适合自动化批处理
2. **性能**: 行数检查在数据读取完成后进行，不影响读取性能
3. **兼容性**: 支持所有已支持的Excel文件格式
4. **日志**: 所有操作都会记录到日志文件中

## 未来改进

1. **配置化**: 将预期行数设为可配置参数
2. **批处理模式**: 增加非交互模式选项
3. **详细报告**: 生成数据行数统计报告
4. **自动修复**: 提供自动修复数据行数的选项 