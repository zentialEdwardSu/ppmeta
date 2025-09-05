# PPMeta GitHub Actions 工作流简化说明

## 简化内容

根据 StackOverflow 的最佳实践建议，已将复杂的工作流文件简化为核心功能：

### 主要改进

1. **移除冗余调试信息** - 删除了大量的诊断输出和错误检查代码
2. **采用标准解决方案** - 使用 `DisableOutOfProcBuild.exe` 解决 HRESULT = '8000000A' 错误
3. **简化证书处理** - 保留核心证书导入功能，移除复杂的错误处理
4. **标准化构建流程** - 按照 StackOverflow 推荐的步骤序列

### 核心工作流程

#### build-installer.yml
- 证书导入（如果提供）
- DisableOutOfProcBuild 修复（解决 VSTO 项目构建问题）
- 使用 devenv 构建 VSTO 项目和安装程序
- 基础错误处理

#### nightly-build.yml
- 检查24小时内的提交
- 构建安装程序
- 创建 MSI 和 VSTO 两种包格式
- 上传构件

#### release-build.yml  
- 标签触发的发布构建
- 自动创建 GitHub Release
- 包含安装指南和证书

### 关键技术点

1. **HRESULT = '8000000A' 解决方案**：
   ```powershell
   $vsWherePath = "${Env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
   $vsPath = & $vsWherePath -latest -products * -property 'installationPath'
   $disableOutOfProcPath = "$vsPath\Common7\IDE\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild"
   & ".\DisableOutOfProcBuild.exe"
   ```

2. **VSTO 项目构建**：
   ```bash
   msbuild ppmeta\ppmeta.csproj -t:rebuild /p:Platform="Any CPU" /p:Configuration="Release"
   devenv.com ppSetup\ppSetup.vdproj /build Release
   ```

3. **证书管理**：
   ```powershell
   Import-PfxCertificate -FilePath "cert.pfx" -CertStoreLocation Cert:\CurrentUser\My -Password $password
   ```

## 参考资料

- [StackOverflow: Build installer using GitHub Actions](https://stackoverflow.com/questions/71823928/build-installer-using-github-actions)
- 采用了 flywhc 和 sweetfa 的解决方案
- DisableOutOfProcBuild 工具解决 Visual Studio 2019+ 的构建问题
