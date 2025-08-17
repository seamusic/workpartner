# DataFixter 项目结构说明

## 项目概述

DataFixter 是一个独立的.NET项目，专门用于修复监测数据中累计变化量计算错误的工具。该项目在WorkPartner解决方案中作为独立模块存在。

## 解决方案结构

```
WorkPartner.sln (解决方案文件)
├── WorkPartner/                    # 主项目
│   └── WorkPartner.csproj
├── Tests/                          # 主项目测试 (WorkPartner.Tests)
│   └── WorkPartner.Tests.csproj
├── DataFixter/                     # 数据修正工具项目 ⭐
│   ├── DataFixter.csproj          # 主项目文件
│   ├── Models/                     # 数据模型
│   ├── Services/                   # 业务服务
│   ├── Excel/                      # Excel处理
│   ├── Configuration/              # 配置文件
│   ├── Logging/                    # 日志配置
│   └── Program.cs                  # 主程序
└── DataFixter.Tests/               # 数据修正工具测试项目 ⭐
    └── DataFixter.Tests.csproj
```

## 为什么需要独立的测试项目？

### 1. 项目独立性
- DataFixter 是一个功能完整的独立工具
- 有自己的数据模型、业务逻辑和Excel处理功能
- 需要专门的测试来验证其功能正确性

### 2. 测试关注点不同
- **WorkPartner.Tests**: 专注于主项目的业务逻辑测试
- **DataFixter.Tests**: 专注于数据修正算法的测试
- 两个项目的测试需求和测试数据完全不同

### 3. 依赖管理
- DataFixter 使用特定的NuGet包（如NPOI）
- 测试项目需要引用这些包来创建测试数据
- 独立的测试项目可以更好地管理依赖关系

## 测试项目结构

```
DataFixter.Tests/
├── DataFixter.Tests.csproj        # 测试项目文件
├── DataValidationServiceTests.cs  # 数据验证服务测试
├── DataProcessingTests.cs         # 数据处理服务测试
└── TestData/                      # 测试数据文件夹（可选）
```

## 测试项目配置

### 1. 项目引用
```xml
<ProjectReference Include="../DataFixter.csproj" />
```

### 2. 测试框架
- **xUnit**: 主要的测试框架
- **Moq**: 用于创建模拟对象
- **FluentAssertions**: 用于更易读的断言

### 3. 测试数据
- 使用内存中的模拟数据
- 不依赖外部Excel文件
- 专注于算法逻辑的验证

## 运行测试

### 1. 运行所有测试
```bash
cd DataFixter.Tests
dotnet test
```

### 2. 运行特定测试
```bash
dotnet test --filter "TestCumulativeChangeValidation_ValidData"
```

### 3. 生成测试报告
```bash
dotnet test --collect:"XPlat Code Coverage"
```

## 测试策略

### 1. 单元测试
- **数据验证逻辑**: 测试累计变化量计算验证
- **数据处理逻辑**: 测试分组、排序、完整性检查
- **Excel处理逻辑**: 测试文件读取和数据解析

### 2. 集成测试
- **完整流程测试**: 从Excel读取到数据修正的完整流程
- **边界条件测试**: 处理异常数据和边界情况
- **性能测试**: 大文件处理性能验证

### 3. 测试数据管理
- **模拟数据**: 创建各种测试场景的模拟数据
- **真实数据**: 使用小规模的真实Excel文件进行测试
- **异常数据**: 故意创建错误数据来测试错误处理

## 与主项目的关系

### 1. 功能独立性
- DataFixter 可以独立运行和测试
- 不依赖 WorkPartner 主项目的功能
- 有自己的命令行接口和配置

### 2. 代码共享
- 如果将来需要，可以提取公共代码到共享库
- 目前保持项目独立性，便于维护和测试

### 3. 部署方式
- 可以独立编译和部署
- 可以作为工具库被其他项目引用
- 支持命令行和程序集两种使用方式

## 总结

DataFixter 项目采用独立的测试项目结构，这是基于以下考虑：

1. **项目独立性**: 作为独立工具，需要专门的测试验证
2. **测试专注性**: 专注于数据修正算法的测试，与主项目测试分离
3. **维护便利性**: 独立的测试项目便于维护和扩展
4. **部署灵活性**: 可以独立部署和测试

这种结构确保了项目的质量和可维护性，同时保持了与主项目的清晰分离。
