# VSTO构建问题修复说明

## 已修复的问题

### 1. BaseOutputPath/OutputPath 错误
**错误信息：** 
```
The BaseOutputPath/OutputPath property is not set for project 'ppmeta.csproj'. Please check to make sure that you have specified a valid combination of Configuration and Platform for this project.
```

**原因：** Platform参数格式不正确
**修复：** 
- 将 `-p:Platform="Any CPU"` 改为 `-p:Platform=AnyCPU`
- MSBuild中Platform参数不应该包含空格

### 2. 证书导入问题
**问题：** 证书文件路径和错误处理不完善
**修复：**
- 使用完整的绝对路径：`${{ github.workspace }}\ppmeta\ppmeta_TemporaryKey.pfx`
- 添加证书文件创建验证
- 添加证书导入的错误处理
- 添加详细的日志输出

### 3. 构建路径问题
**问题：** 构建命令中使用了错误的项目路径
**修复：**
- 统一使用解决方案文件：`ppmeta.sln` 而不是单独的 `.csproj` 文件
- 确保所有路径都是绝对路径

## 修复后的关键配置

### 证书处理
```powershell
# 创建证书文件
$certPath = "${{ github.workspace }}\ppmeta\ppmeta_TemporaryKey.pfx"
[System.IO.File]::WriteAllBytes($certPath, $certBytes)

# 验证文件创建
if (Test-Path $certPath) {
    Write-Host "Certificate file verified"
} else {
    Write-Error "Certificate file creation failed"
    exit 1
}

# 导入证书
try {
    $cert = Import-PfxCertificate -FilePath $certPath -CertStoreLocation Cert:\CurrentUser\My -Password $pwd
    Write-Host "Certificate imported successfully with thumbprint: $($cert.Thumbprint)"
} catch {
    Write-Error "Certificate import failed: $($_.Exception.Message)"
    exit 1
}
```

### MSBuild 命令
```cmd
msbuild "${{ github.workspace }}\ppmeta.sln" -p:Configuration=Release -p:Platform=AnyCPU -p:VisualStudioVersion="17.0" -nologo
```

### 安装包构建
```cmd
devenv.com "${{ github.workspace }}\ppmeta.sln" /build "Release"
```

## 调试信息
添加了以下调试步骤来帮助问题排查：

1. **项目配置检查**：验证项目文件是否存在
2. **目录结构列表**：显示工作空间目录内容
3. **证书文件验证**：确认证书文件创建成功
4. **构建输出列表**：显示构建结果文件

## 建议的测试流程

1. **使用简化版workflow**：首先测试 `simple-build.yml`
2. **检查构建日志**：查看所有调试输出信息
3. **验证证书**：确保三个secrets正确设置
4. **检查artifacts**：下载并验证生成的安装包

## 常见问题排查

### 如果仍然出现Platform错误：
- 检查项目文件中的Platform配置
- 确认解决方案文件中的Platform设置
- 尝试使用 `"Any CPU"` 或 `AnyCPU`

### 如果证书导入失败：
- 验证VSTO_CERTIFICATE是否为有效的Base64编码
- 检查VSTO_CERT_PASSWORD是否正确
- 确认证书文件格式为.pfx

### 如果构建过程中断：
- 查看"Check project configuration"步骤的输出
- 确认所有必需的文件都存在
- 检查NuGet包还原是否成功
