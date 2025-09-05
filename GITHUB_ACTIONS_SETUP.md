# VSTO PowerPoint Add-in CI/CD Setup

本项目包含了自动化构建和发布VSTO PowerPoint插件的GitHub Actions workflows。

## GitHub Secrets 配置

在使用workflows之前，需要在GitHub仓库中配置以下secrets：

### 必需的Secrets

1. **VSTO_CERTIFICATE**
   - 描述：PFX证书文件的Base64编码字符串
   - 获取方法：
     ```powershell
     $certBytes = [System.IO.File]::ReadAllBytes("path\to\your\certificate.pfx")
     $base64Cert = [System.Convert]::ToBase64String($certBytes)
     Write-Output $base64Cert
     ```

2. **VSTO_CERT_PASSWORD**
   - 描述：PFX证书的密码
   - 注意：如果证书没有密码，请设置为空字符串

3. **VSTO_CERT_THUMBPRINT**
   - 描述：证书的指纹（用于代码签名）
   - 获取方法：
     ```powershell
     $cert = Get-PfxCertificate -FilePath "path\to\your\certificate.pfx"
     Write-Output $cert.Thumbprint
     ```

### 如何设置Secrets

1. 进入GitHub仓库
2. 点击 `Settings` 标签
3. 在左侧菜单中选择 `Secrets and variables` > `Actions`
4. 点击 `New repository secret`
5. 添加上述三个secrets

## Workflows 说明

### 1. build-installer.yml
- **触发条件**：push到master/main/release分支，PR，手动触发
- **功能**：
  - 构建VSTO项目
  - 生成MSI安装包
  - 上传构建结果作为artifacts

### 2. release-installer.yml
- **触发条件**：创建GitHub Release，手动触发
- **功能**：
  - 构建发布版本
  - 更新版本号
  - 对MSI进行代码签名
  - 自动上传到GitHub Release

## 使用方法

### 开发阶段
1. 提交代码到master/main分支
2. GitHub Actions自动触发构建
3. 在Actions页面查看构建结果
4. 下载artifacts查看生成的安装包

### 发布版本
1. 在GitHub上创建新的Release
2. 设置版本标签（如v1.0.0）
3. GitHub Actions自动构建并上传安装包到Release

### 手动触发
1. 进入Actions页面
2. 选择相应的workflow
3. 点击"Run workflow"按钮

## 项目结构要求

确保项目结构如下：
```
ppmeta/
├── .github/
│   └── workflows/
│       ├── build-installer.yml
│       └── release-installer.yml
├── ppmeta/
│   ├── ppmeta.csproj
│   ├── ppmeta_TemporaryKey.pfx (运行时创建)
│   └── Properties/
│       └── AssemblyInfo.cs
├── ppSetup/
│   └── ppSetup.vdproj
└── ppmeta.sln
```

## 故障排除

### 常见问题

1. **证书导入失败**
   - 检查VSTO_CERTIFICATE是否正确编码
   - 确认VSTO_CERT_PASSWORD是否正确

2. **构建失败**
   - 检查项目引用是否正确
   - 确认NuGet包是否可以正常还原

3. **安装包生成失败**
   - 确认ppSetup.vdproj配置正确
   - 检查Visual Studio Installer Project是否正确配置

4. **代码签名失败**
   - 确认证书指纹VSTO_CERT_THUMBPRINT是否正确
   - 检查证书是否有效且适用于代码签名

### 调试方法

1. 查看Actions运行日志
2. 检查"List output files for debugging"步骤的输出
3. 下载artifacts查看生成的文件

## 版本管理

- 版本号在AssemblyInfo.cs中自动更新
- 使用语义化版本号（如1.0.0）
- Release标签格式：v1.0.0

## 安全注意事项

1. 证书文件仅在构建过程中临时创建
2. 构建完成后自动清理证书文件
3. 所有敏感信息都存储在GitHub Secrets中
4. 证书密码使用SecureString处理

## 支持的环境

- **操作系统**：Windows Server 2022
- **Visual Studio**：2022 (v17.0)
- **MSBuild**：17.0
- **.NET Framework**：4.8
- **Office**：支持VSTO的各版本PowerPoint
