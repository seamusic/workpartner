# WorkPartner 测试配置说明

## 测试类型

### 1. 单元测试 (Unit Tests)
- **文件**: `DataProcessorRefactoredTests.cs`
- **覆盖范围**: 
  - 累计变化量计算逻辑
  - 连续缺失数据处理
  - 补充文件微调机制
  - 配置参数验证
  - 边界条件测试

### 2. 集成测试 (Integration Tests)
- **文件**: `DataProcessorIntegrationTests.cs`
- **覆盖范围**:
  - 完整数据处理流程
  - 大规模数据处理
  - 错误处理机制
  - 内存使用监控
  - 业务逻辑一致性

### 3. 性能测试 (Performance Tests)
- **文件**: `DataProcessorPerformanceTests.cs`
- **覆盖范围**:
  - 处理速度测试
  - 内存使用测试
  - 缓存效果验证
  - 扩展性测试
  - 压力测试

## 运行方式

### PowerShell 脚本
```powershell
# 运行所有测试
.\run_tests.ps1

# 运行特定类型测试
.\run_tests.ps1 -TestType unit
.\run_tests.ps1 -TestType integration
.\run_tests.ps1 -TestType performance

# 详细输出
.\run_tests.ps1 -Verbose

# 代码覆盖率
.\run_tests.ps1 -Coverage
```

### 直接运行
```bash
# 运行所有测试
dotnet test

# 运行特定测试
dotnet test --filter "Category=Unit"
dotnet test --filter "Category=Integration"
dotnet test --filter "Category=Performance"
```

## 测试数据

### 测试数据集大小
- **小数据集**: 50个文件
- **中数据集**: 100-200个文件
- **大数据集**: 300-500个文件
- **压力测试**: 并发处理多个数据集

### 数据特征
- 包含G列和D列（累计变化量和变化量）
- 随机缺失值（20-30%概率）
- 模拟真实业务场景
- 支持自定义配置参数

## 性能基准

### 处理时间要求
- 50个文件: < 1秒
- 100个文件: < 2秒
- 200个文件: < 5秒
- 500个文件: < 20秒

### 内存使用要求
- 峰值内存增长: < 200MB
- 最终内存增长: < 50MB
- 支持大规模数据处理

### 缓存效果
- 缓存版本应该比无缓存版本快
- 性能提升比例: > 1.0x

## 测试环境要求

### 硬件要求
- 内存: 至少4GB可用内存
- CPU: 支持多线程处理
- 磁盘: 足够的临时存储空间

### 软件要求
- .NET 6.0 或更高版本
- xUnit 测试框架
- PowerShell 5.0 或更高版本

## 测试结果

### 输出格式
- TRX格式测试结果文件
- 控制台详细输出
- 性能指标统计
- 内存使用报告

### 结果位置
- 测试结果目录: `TestResults/`
- TRX文件: `TestResults/*.trx`
- 覆盖率报告: `TestResults/coverage/`

## 故障排除

### 常见问题
1. **构建失败**: 检查依赖包和项目配置
2. **测试超时**: 调整性能测试的时间限制
3. **内存不足**: 减少测试数据集大小
4. **权限问题**: 确保有写入测试结果目录的权限

### 调试建议
- 使用 `-Verbose` 参数获取详细输出
- 检查测试日志文件
- 验证测试数据生成逻辑
- 监控系统资源使用情况
