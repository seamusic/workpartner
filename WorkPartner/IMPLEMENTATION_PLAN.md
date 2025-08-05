# WorkPartner 技术实现计划

## 项目架构设计

### 1. 项目结构
```
WorkPartner/
├── Models/
│   ├── ExcelFile.cs          # Excel文件信息模型
│   ├── DataRow.cs            # 数据行模型
│   └── ProcessingResult.cs   # 处理结果模型
├── Services/
│   ├── IExcelService.cs      # Excel服务接口
│   ├── ExcelService.cs       # Excel服务实现
│   ├── IFileService.cs       # 文件服务接口
│   └── FileService.cs        # 文件服务实现
├── Utils/
│   ├── FileNameParser.cs     # 文件名解析工具
│   ├── DataProcessor.cs      # 数据处理工具
│   └── Logger.cs             # 日志工具
├── Program.cs                # 主程序入口
└── WorkPartner.csproj        # 项目文件
```

### 2. 核心类设计

#### ExcelFile 模型
```csharp
public class ExcelFile
{
    public string FilePath { get; set; }
    public string FileName { get; set; }
    public DateTime Date { get; set; }
    public int Hour { get; set; }
    public string ProjectName { get; set; }
    public List<DataRow> DataRows { get; set; }
}
```

#### DataRow 模型
```csharp
public class DataRow
{
    public string Name { get; set; }        // B列数据名称
    public List<double?> Values { get; set; } // D-I列数据值
    public int RowIndex { get; set; }       // 在Excel中的行索引
}
```

## 实现步骤详解

### 步骤1：项目依赖配置
1. 添加EPPlus包用于Excel操作
2. 添加System.IO.Abstractions用于文件操作抽象
3. 配置项目属性

### 步骤2：文件名解析器实现
```csharp
public class FileNameParser
{
    private static readonly Regex FileNameRegex = 
        new Regex(@"^(\d{4}\.\d{1,2}\.\d{1,2})-(\d{1,2})(.+)$");
    
    public static ExcelFile ParseFileName(string filePath)
    {
        // 实现文件名解析逻辑
    }
}
```

### 步骤3：Excel数据读取
```csharp
public class ExcelService : IExcelService
{
    public List<DataRow> ReadExcelData(string filePath)
    {
        // 读取B5-B368行的名称
        // 读取D5-I5列的数据
        // 处理空值
    }
}
```

### 步骤4：数据补充算法
```csharp
public class DataProcessor
{
    public List<ExcelFile> ProcessMissingData(List<ExcelFile> files)
    {
        // 1. 按时间顺序排序文件
        // 2. 识别每个文件中的缺失数据
        // 3. 计算前后平均值
        // 4. 补充缺失数据
    }
    
    private double? CalculateAverageValue(string dataName, List<ExcelFile> files, int currentIndex)
    {
        // 实现平均值计算逻辑
    }
}
```

### 步骤5：完整性检查
```csharp
public class CompletenessChecker
{
    public List<ExcelFile> CheckAndFillMissingFiles(List<ExcelFile> files)
    {
        // 1. 按日期分组
        // 2. 检查每天是否有0、8、16三份数据
        // 3. 复制缺失的文件
    }
}
```

## 关键算法实现

### 数据补充算法
1. **缺失数据识别**：遍历每个Excel文件的B5-B368行，检查D-I列是否有空值
2. **前后数据查找**：对于每个缺失的数据点，在前后文件中查找相同B列名称的非空值
3. **平均值计算**：计算找到的所有非空值的平均值
4. **数据补充**：用计算出的平均值替换空值

### 完整性检查算法
1. **日期分组**：按日期对文件进行分组
2. **时间点检查**：检查每组中是否包含0、8、16三个时间点
3. **文件复制**：对于缺失的时间点，复制当天任意文件并重命名

## 错误处理策略

### 1. 文件读取错误
- 检查文件是否存在
- 验证文件格式是否为Excel
- 处理文件被占用的情况

### 2. 数据格式错误
- 验证文件名格式
- 检查Excel数据结构
- 处理空文件或损坏文件

### 3. 数据补充错误
- 处理没有前后数据的情况
- 处理所有数据都缺失的情况
- 记录补充失败的数据点

## 性能优化考虑

### 1. 内存管理
- 使用流式读取大文件
- 及时释放Excel对象
- 避免一次性加载所有文件到内存

### 2. 并行处理
- 并行读取多个Excel文件
- 并行处理数据补充
- 控制并发数量避免资源竞争

### 3. 缓存策略
- 缓存文件名解析结果
- 缓存数据补充计算结果
- 避免重复计算

## 测试策略

### 1. 单元测试
- 文件名解析测试
- 数据读取测试
- 数据补充算法测试

### 2. 集成测试
- 完整流程测试
- 边界条件测试
- 错误处理测试

### 3. 性能测试
- 大文件处理测试
- 内存使用测试
- 处理速度测试

## 部署和配置

### 1. 命令行参数
```bash
WorkPartner.exe <input_folder> [output_folder] [options]
```

### 2. 配置文件
- 支持配置文件指定处理参数
- 支持日志级别配置
- 支持输出格式配置

### 3. 日志输出
- 控制台输出处理进度
- 文件日志记录详细信息
- 错误日志单独记录 