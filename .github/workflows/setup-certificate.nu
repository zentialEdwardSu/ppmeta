#!/usr/bin/env nu
# VSTO证书设置助手脚本 (Nushell版本)
# 用于帮助用户准备GitHub Actions所需的证书配置

def main [
    --certificate-path (-c): string   # PFX证书文件路径
    --certificate-password (-p): string # 证书密码
    --create-self-signed (-s)          # 创建自签名证书用于开发测试
    --help (-h)                        # 显示帮助信息
] {
    if $help {
        show_help
        return
    }

    print "🔐 VSTO证书设置助手 (Nushell版本)"
    print "========================================"

    let cert_info = if $create_self_signed {
        create_development_certificate
    } else if ($certificate_path | is-not-empty) {
        let password = if ($certificate_password | is-empty) {
            input "请输入证书密码: " --suppress-output
        } else {
            $certificate_password
        }
        
        print "🔍 正在验证证书..."
        get_certificate_info $certificate_path $password
    } else {
        show_help
        return
    }

    if ($cert_info | is-not-empty) {
        print "📦 正在转换证书为Base64格式..."
        let base64_cert = convert_to_base64_certificate $cert_info.path
        
        if ($base64_cert | is-not-empty) {
            show_github_secrets_config $cert_info $base64_cert
            save_config_to_file $cert_info $base64_cert
        }
    } else {
        print "❌ 证书设置失败"
        exit 1
    }

    print "✅ 证书设置完成!"
}

def show_help [] {
    print $"
VSTO证书设置助手 \(Nushell版本\)

用法:
  nu setup-certificate.nu --certificate-path \"path/to/cert.pfx\" --certificate-password \"password\"
  nu setup-certificate.nu --create-self-signed
  nu setup-certificate.nu --help

参数:
  --certificate-path, -c    PFX证书文件路径
  --certificate-password, -p 证书密码
  --create-self-signed, -s  创建自签名证书用于开发测试
  --help, -h               显示此帮助信息

输出:
  脚本会生成GitHub Secrets所需的配置信息，包括：
  - VSTO_CERTIFICATE \(Base64编码的证书\)
  - VSTO_CERT_PASSWORD \(证书密码\)
  - VSTO_CERT_THUMBPRINT \(证书指纹\)

示例:
  # 使用现有证书
  nu setup-certificate.nu -c \"my-cert.pfx\" -p \"MyPassword123\"
  
  # 创建开发用自签名证书
  nu setup-certificate.nu --create-self-signed
"
}

def create_development_certificate [] {
    print "🔧 正在创建自签名开发证书..."
    
    try {
        let random_num = (date now | format date "%f" | str substring 0..3)
        let password = $"PPMeta-Dev-($random_num)"
        let cert_path = "ppmeta-dev-certificate.pfx"
        
        let result = (powershell -c $"
            \$cert = New-SelfSignedCertificate -Subject 'CN=PPMeta Development Certificate' -Type CodeSigning -KeyUsage DigitalSignature -FriendlyName 'PPMeta Development Certificate' -CertStoreLocation Cert:\\CurrentUser\\My -HashAlgorithm SHA256 -NotAfter \(Get-Date\).AddYears\(2\)
            \$securePassword = ConvertTo-SecureString -String '($password)' -Force -AsPlainText
            Export-PfxCertificate -Cert \$cert -FilePath '($cert_path)' -Password \$securePassword | Out-Null
            Write-Output \$cert.Thumbprint
        ")
        
        let thumbprint = ($result | str trim)
        
        print "✅ 开发证书创建成功!"
        print $"📄 证书文件: ($cert_path)"
        print $"🔒 证书密码: ($password)"
        print $"👆 证书指纹: ($thumbprint)"
        
        {
            path: $cert_path,
            password: $password,
            thumbprint: $thumbprint,
            subject: "CN=PPMeta Development Certificate",
            not_after: "2027-08-30 00:00:00"
        }
    } catch {
        print "❌ 创建证书失败"
        {}
    }
}

def get_certificate_info [path: string, password: string] {
    try {
        if not ($path | path exists) {
            error make { msg: $"证书文件不存在: ($path)" }
        }
        
        let cert_info = (powershell -c $"
            try {
                \$cert = Get-PfxCertificate -FilePath '($path)'
                \$info = @{
                    Subject = \$cert.Subject
                    Thumbprint = \$cert.Thumbprint
                    NotAfter = \$cert.NotAfter.ToString\('yyyy-MM-dd HH:mm:ss'\)
                    HasCodeSigning = \$cert.EnhancedKeyUsageList -like '*Code Signing*'
                    IsExpired = \$cert.NotAfter -lt \(Get-Date\)
                    ExpiresIn30Days = \$cert.NotAfter -lt \(Get-Date\).AddDays\(30\)
                }
                \$info | ConvertTo-Json
            } catch {
                Write-Error \$_.Exception.Message
                exit 1
            }
        ")
        
        let cert_data = ($cert_info | from json)
        
        # 验证证书
        if not $cert_data.HasCodeSigning {
            print "⚠️  警告: 证书可能不支持代码签名"
        }
        
        if $cert_data.IsExpired {
            error make { msg: $"证书已过期: ($cert_data.NotAfter)" }
        }
        
        if $cert_data.ExpiresIn30Days {
            print $"⚠️  警告: 证书将在30天内过期: ($cert_data.NotAfter)"
        }
        
        {
            path: $path,
            password: $password,
            thumbprint: $cert_data.Thumbprint,
            subject: $cert_data.Subject,
            not_after: $cert_data.NotAfter
        }
    } catch {
        print $"❌ 证书验证失败: ($in)"
        {}
    }
}

def convert_to_base64_certificate [path: string] {
    try {
        let base64_result = (powershell -c $"
            try {
                \$certBytes = [System.IO.File]::ReadAllBytes\('($path)'\)
                [System.Convert]::ToBase64String\(\$certBytes\)
            } catch {
                Write-Error \$_.Exception.Message
                exit 1
            }
        ")
        
        $base64_result | str trim
    } catch {
        print "❌ 转换证书失败"
        ""
    }
}

def show_github_secrets_config [cert_info: record, base64_cert: string] {
    print ""
    print "🚀 GitHub Secrets 配置"
    print "=================================================="
    print ""
    print "请在GitHub仓库设置中添加以下Secrets:"
    print ""
    
    print "1. VSTO_CERTIFICATE"
    print $"   ($base64_cert)"
    print ""
    
    print "2. VSTO_CERT_PASSWORD"
    print $"   ($cert_info.password)"
    print ""
    
    print "3. VSTO_CERT_THUMBPRINT"
    print $"   ($cert_info.thumbprint)"
    print ""
    
    print "📋 证书信息:"
    print $"   主题: ($cert_info.subject)"
    print $"   有效期至: ($cert_info.not_after)"
    print $"   指纹: ($cert_info.thumbprint)"
    print ""
    
    print "🔗 设置步骤:"
    print "   1. 访问 GitHub 仓库 -> Settings -> Secrets and variables -> Actions"
    print "   2. 点击 'New repository secret'"
    print "   3. 添加上述三个 secrets"
    print "   4. 运行 GitHub Actions workflow"
}

def save_config_to_file [cert_info: record, base64_cert: string] {
    let config_file = "github-secrets-config.txt"
    let current_time = (date now | format date "%Y-%m-%d %H:%M:%S")
    
    let config_content = $"GitHub Secrets Configuration for PPMeta VSTO Project
Generated: ($current_time)

VSTO_CERTIFICATE:
($base64_cert)

VSTO_CERT_PASSWORD:
($cert_info.password)

VSTO_CERT_THUMBPRINT:
($cert_info.thumbprint)

Certificate Information:
Subject: ($cert_info.subject)
Valid Until: ($cert_info.not_after)
Thumbprint: ($cert_info.thumbprint)
"
    
    $config_content | save $config_file
    
    print $"💾 配置已保存到: ($config_file)"
    print "⚠️  请妥善保管此文件，包含敏感信息!"
}

# 如果直接运行脚本（不是通过 source 加载）
# Nushell会自动调用main函数
