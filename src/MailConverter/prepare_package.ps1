# prepare_package.ps1
# 发布前准备：保留CardDAV配置，清空其他配置
param(
    [string]$SourceDir = "bin\Release\net48",
    [string]$OutputDir = "bin\Release\net48\publish-ready",
    [string]$CardDavSource = "bin\Debug\net48\Config\carddav"
)

$ErrorActionPreference = "Stop"

# 获取脚本所在目录
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host "=== 发布准备开始 ===" -ForegroundColor Cyan
Write-Host "源目录: $SourceDir"
Write-Host "输出目录: $OutputDir"
Write-Host ""

# 1. 复制整个发布目录到临时目录
if (Test-Path $OutputDir) {
    Write-Host "清理旧输出目录..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $OutputDir
}

Write-Host "复制发布文件..." -ForegroundColor Green
Copy-Item -Path $SourceDir -Destination $OutputDir -Recurse

# 2. 创建Config目录结构（如果不存在）
$ConfigDir = Join-Path $OutputDir "Config"
$OAuthDir = Join-Path $ConfigDir "oauth"
$ImapDir = Join-Path $ConfigDir "imap"
$CardDavDir = Join-Path $ConfigDir "carddav"
$PstDir = Join-Path $ConfigDir "pst"

# 确保目录存在
New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
New-Item -ItemType Directory -Path $OAuthDir -Force | Out-Null
New-Item -ItemType Directory -Path $ImapDir -Force | Out-Null
New-Item -ItemType Directory -Path $CardDavDir -Force | Out-Null
New-Item -ItemType Directory -Path $PstDir -Force | Out-Null

# 3. 清空所有配置目录内容
Write-Host "清空配置目录（保留结构）..." -ForegroundColor Yellow
Get-ChildItem $OAuthDir -File | Remove-Item -Force
Get-ChildItem $ImapDir -File | Remove-Item -Force
Get-ChildItem $PstDir -File | Remove-Item -Force

# 4. 保留CardDAV配置
Write-Host "处理CardDAV配置..." -ForegroundColor Green
if (Test-Path $CardDavSource) {
    # 复制CardDAV配置文件
    Copy-Item -Path (Join-Path $CardDavSource "*.inf") -Destination $CardDavDir -Force
    $copied = Get-ChildItem $CardDavDir -File
    if ($copied) {
        Write-Host "  已保留 CardDAV 配置:" -ForegroundColor Cyan
        foreach ($f in $copied) {
            Write-Host "    - $($f.Name)"
        }
    } else {
        Write-Host "  警告: CardDAV源目录没有.inf文件" -ForegroundColor Yellow
    }
} else {
    Write-Host "  警告: CardDAV源目录不存在: $CardDavSource" -ForegroundColor Yellow
}

# 5. 创建空的preferences.inf
$prefsFile = Join-Path $ConfigDir "preferences.inf"
@"
LastUsedEmail=
LastSourcePath=
LastTargetFolder=Inbox
PstTenantId=
PstClientId=
PstClientSecret=
PstAccountName=
PurviewLogPath=
PurviewOutputPath=
"@ | Out-File -FilePath $prefsFile -Encoding UTF8

# 6. 创建空的registration.inf（未注册状态）
$regFile = Join-Path $ConfigDir "registration.inf"
@"
IsRegistered=False
RegisteredUserName=
RegisteredUserEmail=
RegisteredOrganization=
RegisteredMacAddress=
RegisterDate=
RegisterSerialNumber=
RegisterRemainingDays=0
RegisterExpireDate=
FirstRunDate=
"@ | Out-File -FilePath $regFile -Encoding UTF8

# 7. 创建默认的features.inf（功能设置 - 默认全部开启）
$featuresFile = Join-Path $ConfigDir "features.inf"
@"
Feature_ToPst=True
Feature_ToPst_Eml=True
Feature_ToPst_Ost=True
Feature_ToPst_Imap=True
Feature_ToPst_MultiImap=True
Feature_Extract=True
Feature_Extract_Imap=True
Feature_Extract_Files=True
Feature_SingleUserSync=True
Feature_SingleUserSync_EmlImport=True
Feature_SingleUserSync_Contacts=True
Feature_BatchSync=True
Feature_BatchSync_Login=True
Feature_BatchSync_PstMail=True
Feature_BatchSync_PstContacts=True
Feature_BatchSync_PstCalendar=True
Feature_BatchSync_CsvContacts=True
Feature_BatchSync_VcfContacts=True
Feature_BatchSync_CsvCalendar=True
Feature_BatchSync_Purview=True
Feature_O365Toolkit=True
Feature_O365Toolkit_Login=True
Feature_O365Toolkit_Account=True
Feature_O365Toolkit_Group=True
Feature_O365Toolkit_Mobile=True
Feature_O365Toolkit_Traffic=True
Feature_O365Toolkit_Migration=True
Feature_O365Toolkit_Whois=True
Feature_O365Toolkit_Dns=True
Feature_OnPremiseToolkit=True
Feature_Preferences=True
"@ | Out-File -FilePath $featuresFile -Encoding UTF8

# 8. 创建空的Logs目录
$LogsDir = Join-Path $OutputDir "Logs"
if (Test-Path $LogsDir) {
    Remove-Item -Recurse -Force $LogsDir
}
New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "O365Online") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "O365Online\Login") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "O365Online\AccountManagement") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "O365Online\GroupManagement") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "Purview") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "batchToO365") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $LogsDir "Registration") -Force | Out-Null

# 9. 删除trial.dat（试用文件）
$trialFile = Join-Path $ConfigDir "trial.dat"
if (Test-Path $trialFile) {
    Remove-Item -Force $trialFile
    Write-Host "已删除 trial.dat" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== 发布准备完成 ===" -ForegroundColor Cyan
Write-Host "输出目录: $OutputDir"
Write-Host ""
Write-Host "Config目录内容:" -ForegroundColor White
Get-ChildItem $ConfigDir -Recurse -File | ForEach-Object {
    $relPath = $_.FullName.Replace($ConfigDir, "").TrimStart("\")
    Write-Host "  $relPath"
}
