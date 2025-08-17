# Excel格式保持功能说明

## 重要改进概述

在最新版本中，DataFixter实现了一个重要的功能改进：**基于原始文件进行修正，完全保持原始格式**。这解决了之前版本中修正后文件格式与原始文件完全不同的问题。

## 改进前后对比

### ❌ 之前的实现（问题）
- 完全重新创建Excel文件
- 丢失原始格式、样式、公式
- 列结构可能发生变化
- 用户需要重新调整格式

### ✅ 现在的实现（解决方案）
- 基于原始文件进行修正
- 完全保持原始格式、样式、公式
- 列结构完全一致
- 只修改需要修正的数据

## 技术实现原理

### 1. 文件读取策略
```csharp
// 根据文件扩展名选择合适的工作簿类型
IWorkbook workbook;
using (var originalStream = new FileStream(originalFilePath, FileMode.Open, FileAccess.Read))
{
    if (Path.GetExtension(originalFileName).ToLower() == ".xlsx")
    {
        workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(originalStream);
    }
    else
    {
        workbook = new HSSFWorkbook(originalStream);
    }
}
```

### 2. 智能列检测
```csharp
private ColumnMapping DetectColumnMapping(ISheet sheet)
{
    var mapping = new ColumnMapping();
    
    // 查找标题行（通常是第4行，索引3）
    var headerRow = sheet.GetRow(3);
    if (headerRow == null) return mapping;
    
    // 遍历标题行，查找相关列
    for (int i = 0; i < headerRow.LastCellNum; i++)
    {
        var cell = headerRow.GetCell(i);
        if (cell == null) continue;
        
        var cellValue = cell.StringCellValue?.ToLower() ?? "";
        
        // 检测累计变化量列（支持中英文）
        if (cellValue.Contains("累计") && cellValue.Contains("x"))
            mapping.CumulativeX = i;
        else if (cellValue.Contains("cumulative") && cellValue.Contains("x"))
            mapping.CumulativeX = i;
        // ... 其他方向的检测
    }
    
    return mapping;
}
```

### 3. 精确数据修正
```csharp
// 只修正需要修正的单元格，保持原始格式
if (columnMapping.CumulativeX >= 0)
{
    var cell = row.GetCell(columnMapping.CumulativeX) ?? row.CreateCell(columnMapping.CumulativeX);
    cell.SetCellValue(data.CumulativeX);
    correctedRows++;
}
```

## 保持的原始元素

### ✅ 完全保持
- **文件格式**: .xls 或 .xlsx
- **工作表结构**: 工作表名称、行数、列数
- **列标题**: 所有列标题和位置
- **数据格式**: 数字格式、日期格式、文本格式
- **样式设置**: 字体、颜色、边框、背景色
- **公式**: 所有Excel公式
- **列宽行高**: 列宽、行高设置
- **合并单元格**: 合并单元格信息
- **条件格式**: 条件格式规则
- **数据验证**: 数据验证规则

### 🔧 智能修正
- **累计变化量**: 只修正验证失败的累计变化量数据
- **数据精度**: 保持原始数据精度
- **错误处理**: 记录修正原因和调整量

## 使用场景

### 1. 监测数据处理
- 保持原始监测数据的格式和结构
- 只修正累计变化量的计算错误
- 维护数据的专业性和可读性

### 2. 报告生成
- 修正后的文件可以直接用于报告
- 无需重新调整格式
- 保持专业文档的外观

### 3. 数据审计
- 可以清楚看到哪些数据被修正
- 修正记录完整可追溯
- 支持数据质量评估

## 技术优势

### 1. 兼容性
- 支持Excel 97-2003 (.xls) 和 Excel 2007+ (.xlsx)
- 自动识别文件格式
- 向后兼容性好

### 2. 智能性
- 自动检测列位置
- 支持中英文列名
- 容错处理能力强

### 3. 可靠性
- 基于原始文件，不会丢失数据
- 修正过程可追溯
- 支持批量处理

## 注意事项

### 1. 文件权限
- 确保有读取原始文件的权限
- 确保有写入输出目录的权限

### 2. 文件完整性
- 原始文件不应被其他程序占用
- 建议在修正前备份原始文件

### 3. 列名识别
- 列名应包含关键词（如"累计"、"cumulative"）
- 支持模糊匹配，容错性强

## 总结

这个重要改进使DataFixter从一个简单的数据修正工具升级为一个专业的、用户友好的数据处理解决方案。用户现在可以：

1. **保持工作成果**: 不会丢失已经调整好的格式和样式
2. **提高工作效率**: 修正后的文件可以直接使用
3. **降低使用门槛**: 无需重新学习Excel格式设置
4. **保证数据质量**: 在保持格式的同时确保数据正确性

这个改进体现了DataFixter项目对用户体验的重视，以及对专业数据处理需求的深度理解。
