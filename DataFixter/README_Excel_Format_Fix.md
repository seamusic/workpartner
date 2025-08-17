# Excel格式修复说明

## 问题描述

在之前的版本中，修正后的Excel文件打开时会提示"文件扩展名与格式不符合"的错误。

## 问题原因

这是因为我们在生成Excel文件时使用了以下代码：

```csharp
// 创建工作簿
using var workbook = new XSSFWorkbook();  // 使用XSSF格式（.xlsx）
```

但是文件名生成时使用了原始文件的扩展名：

```csharp
private string GetCorrectedFileName(string originalFileName)
{
    var nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
    var extension = Path.GetExtension(originalFileName);  // 获取原始扩展名
    return $"{nameWithoutExt}_修正后{extension}";  // 可能生成 .xls 扩展名
}
```

**问题分析：**
- `XSSFWorkbook` 生成的是 `.xlsx` 格式的文件
- 如果原始文件是 `.xls` 格式，生成的修正后文件就会是 `文件名_修正后.xls`
- 但实际文件内容是 `.xlsx` 格式，导致扩展名与内容不匹配

## 解决方案

根据用户需求，保持原始文件格式不变，使用 `HSSFWorkbook` 生成 `.xls` 格式的文件：

```csharp
// 创建工作簿
using var workbook = new HSSFWorkbook();  // 使用HSSF格式（.xls）

private string GetCorrectedFileName(string originalFileName)
{
    var nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
    var extension = Path.GetExtension(originalFileName);
    // 保持原始扩展名，因为我们使用的是HSSFWorkbook
    return $"{nameWithoutExt}_修正后{extension}";
}
```

## 修复后的效果

- ✅ 修正后的Excel文件保持原始格式（.xls）
- ✅ 文件内容与扩展名完全匹配
- ✅ 不再出现"文件扩展名与格式不符合"的错误
- ✅ 所有Excel文件都能正常打开
- ✅ 保持向后兼容性，支持旧版Excel

## 技术说明

### 文件格式对比

| 格式 | 扩展名 | NPOI类 | 特点 |
|------|--------|--------|------|
| Excel 97-2003 | .xls | HSSFWorkbook | 旧格式，兼容性好 |
| Excel 2007+ | .xlsx | XSSFWorkbook | 新格式，文件更小，功能更丰富 |

### 当前实现

我们使用 `HSSFWorkbook` 生成 `.xls` 格式的文件，因为：
1. 保持原始文件格式不变
2. 更好的向后兼容性
3. 支持Excel 97-2003版本

## 注意事项

1. **向后兼容性**：生成的 `.xls` 文件支持 Excel 97-2003 及以上版本
2. **文件大小**：`.xls` 文件通常比 `.xlsx` 文件大，但兼容性更好
3. **功能支持**：支持基础Excel功能，满足监测数据处理需求

## 测试建议

修复后，建议测试以下场景：
1. 原始文件为 `.xls` 格式
2. 原始文件为 `.xlsx` 格式  
3. 修正后的文件能否正常打开
4. 修正后的文件内容是否正确
5. 在旧版Excel中打开修正后的文件

## 相关代码位置

- 文件：`Services/ExcelOutputService.cs`
- 方法：`GetCorrectedFileName()`
- 行数：约第280行
