# WorkPartner 发布指南

## 📋 概述

WorkPartner 支持多种发布方式，包括传统的 JIT 编译和现代的 AOT 编译。本指南将详细介绍各种发布选项。

## 🚀 快速发布

### 1. 标准发布（推荐）

```bash
# 框架依赖发布（需要 .NET Runtime）
dotnet publish -c Release -r win-x64

# 自包含发布（包含运行时）
dotnet publish -c Release -r win-x64 --self-contained

# 单文件发布（便携部署）
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

### 2. 自包含发布

```bash
# 自包含发布（包含运行时）
dotnet publish -c Release -r win-x64 --self-contained

# 自包含发布（单文件）
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## 🔧 AOT 编译发布

### 前提条件

AOT 编译需要安装 C++ 构建工具：

1. **Windows**：
   - 安装 Visual Studio 2022 Community/Professional/Enterprise
   - 在安装程序中选择 "Desktop development with C++" 工作负载
   - 或者安装 Visual Studio Build Tools 2022

2. **Linux**：
   ```bash
   sudo apt-get install clang zlib1g-dev
   ```

3. **macOS**：
   ```bash
   xcode-select --install
   ```

### AOT 发布命令

```bash
# AOT 发布（需要 C++ 工具）
dotnet publish -c Release -r win-x64 --self-contained

# AOT 发布（单文件）
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## 📦 发布配置说明

### 项目文件配置

```xml
<PropertyGroup>
  <OutputType>Exe</OutputType>
  <TargetFramework>net8.0</TargetFramework>
  <ImplicitUsings>enable</ImplicitUsings>
  <Nullable>enable</Nullable>
  
  <!-- AOT编译配置 - 仅在发布时启用 -->
  <PublishAot Condition="'$(Configuration)' == 'Release'">true</PublishAot>
  <InvariantGlobalization>true</InvariantGlobalization>
  
  <!-- 默认运行时标识符 -->
  <RuntimeIdentifier Condition="'$(RuntimeIdentifier)' == ''">win-x64</RuntimeIdentifier>
</PropertyGroup>
```

### 配置说明

- **PublishAot**: 启用 AOT 编译（仅在 Release 配置时）
- **InvariantGlobalization**: 减少运行时大小
- **RuntimeIdentifier**: 默认目标平台

## 🎯 不同平台的发布

### Windows

```bash
# x64 架构
dotnet publish -c Release -r win-x64

# x86 架构
dotnet publish -c Release -r win-x86

# ARM64 架构
dotnet publish -c Release -r win-arm64
```

### Linux

```bash
# x64 架构
dotnet publish -c Release -r linux-x64

# ARM64 架构
dotnet publish -c Release -r linux-arm64
```

### macOS

```bash
# x64 架构
dotnet publish -c Release -r osx-x64

# ARM64 架构 (Apple Silicon)
dotnet publish -c Release -r osx-arm64
```

## 📁 发布输出

### 标准发布输出

```
bin/Release/net8.0/win-x64/
├── WorkPartner.exe          # 主程序
├── WorkPartner.dll          # 程序集
├── appsettings.json         # 配置文件
├── *.dll                    # 依赖库
└── *.json                   # 运行时配置
```

### 自包含发布输出

```
bin/Release/net8.0/win-x64/
├── WorkPartner.exe          # 独立可执行文件
├── appsettings.json         # 配置文件
└── *.dll                    # .NET 运行时库
```

### 单文件发布输出

```
bin/Release/net8.0/win-x64/
├── WorkPartner.exe          # 单文件可执行文件
└── appsettings.json         # 配置文件（需要复制）
```

## ⚠️ 常见问题

### 1. RuntimeIdentifier 错误

**错误**: `RuntimeIdentifier is required for native compilation`

**解决方案**:
- 使用 `-r` 参数指定目标平台
- 或在项目文件中设置默认 RuntimeIdentifier

### 2. C++ 工具缺失

**错误**: `Platform linker not found`

**解决方案**:
- 安装 Visual Studio 2022 的 C++ 开发工具
- 或使用标准发布（不使用 AOT）

### 3. 配置文件缺失

**问题**: 发布后配置文件丢失

**解决方案**:
- 确保 `appsettings.json` 设置为 `CopyToOutputDirectory`
- 手动复制配置文件到发布目录

## 🔍 验证发布

### 1. 检查文件

```bash
# 检查发布文件
ls bin/Release/net8.0/win-x64/

# 检查文件大小
dir bin/Release/net8.0/win-x64/
```

### 2. 测试运行

```bash
# 测试可执行文件
./bin/Release/net8.0/win-x64/WorkPartner.exe --help
```

### 3. 性能测试

```bash
# 测试处理性能
./bin/Release/net8.0/win-x64/WorkPartner.exe --input excel/ --output output/
```

## 📊 发布选项对比

| 发布类型 | 文件大小 | 启动速度 | 依赖项 | 适用场景 |
|---------|---------|---------|--------|----------|
| 标准发布 | 中等 | 中等 | 需要 .NET Runtime | 开发环境 |
| 自包含发布 | 较大 | 中等 | 无外部依赖 | 生产环境 |
| AOT 发布 | 较大 | 快速 | 无外部依赖 | 高性能场景 |
| 单文件发布 | 最大 | 中等 | 无外部依赖 | 便携部署 |

## 🎉 最佳实践

1. **开发阶段**: 使用标准发布
2. **测试环境**: 使用自包含发布
3. **生产环境**: 使用 AOT 发布（如果性能要求高）
4. **便携部署**: 使用单文件发布

## 📞 技术支持

如果遇到发布问题，请：

1. 检查是否安装了必要的构建工具
2. 确认目标平台支持
3. 查看详细的错误日志
4. 参考 .NET 官方文档

---

*最后更新: 2025-08-07* 