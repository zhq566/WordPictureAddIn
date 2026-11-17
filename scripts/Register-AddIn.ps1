#requires -Version 3.0
<#
.SYNOPSIS
    注册 WordPictureAddIn（VSTO Word 加载项），无需管理员权限。

.DESCRIPTION
    在 HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn 写入 VSTO 加载项注册项，
    Manifest 指向本地编译输出的 .vsto 文件并附加 |vstolocal 后缀，
    使 VSTO 运行时直接从该目录加载，而不经过 ClickOnce 缓存。

.PARAMETER Target
    Debug（默认）、Release，或任意输出目录的绝对/相对路径。

.EXAMPLE
    .\Register-AddIn.ps1
    .\Register-AddIn.ps1 Release
    .\Register-AddIn.ps1 "D:\build\bin\Debug"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Target = 'Debug'
)

$ErrorActionPreference = 'Stop'

$addInName = 'WordPictureAddIn'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir   = Split-Path -Parent $scriptDir

# ---- 解析输出目录 ----
switch ($Target) {
    'Debug'   { $outDir = Join-Path $rootDir 'bin\Debug' }
    'Release' { $outDir = Join-Path $rootDir 'bin\Release' }
    default {
        if ([System.IO.Path]::IsPathRooted($Target)) {
            $outDir = $Target
        } else {
            $outDir = Join-Path $rootDir $Target
        }
    }
}
$outDir = [System.IO.Path]::GetFullPath($outDir).TrimEnd('\')

Write-Host "[信息] 输出目录: $outDir"

$vstoPath     = Join-Path $outDir "$addInName.vsto"
$manifestPath = Join-Path $outDir "$addInName.dll.manifest"
$dllPath      = Join-Path $outDir "$addInName.dll"

if (-not (Test-Path -LiteralPath $vstoPath)) {
    Write-Host "[错误] 找不到 `"$vstoPath`"" -ForegroundColor Red
    Write-Host "       请先编译项目，或确认路径参数是否正确。" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $manifestPath)) {
    Write-Host "[错误] 找不到 `"$manifestPath`"" -ForegroundColor Red
    exit 1
}

# ---- 注册表版本一致性检查（Word 位数无关，VSTO 走 HKCU） ----
$runtimeCandidates = @(
    (Join-Path $env:CommonProgramFiles   'microsoft shared\VSTO\vstoee.dll')
    (Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared\VSTO\vstoee.dll')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if ($runtimeCandidates.Count -eq 0) {
    Write-Host "[警告] 未检测到 VSTO 运行时 vstoee.dll，Word 可能无法加载此加载项。" -ForegroundColor Yellow
    Write-Host "       请安装 `"Microsoft Visual Studio 2010 Tools for Office Runtime`"。" -ForegroundColor Yellow
}

# ---- 检查 .NET Framework 4.7.2 ----
$ndp = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
if ($null -eq $ndp -or $ndp.Release -lt 461808) {
    Write-Host "[警告] 未检测到 .NET Framework 4.7.2 或更高版本。" -ForegroundColor Yellow
}

# ---- 解除文件锁定 ----
if (Test-Path -LiteralPath $dllPath) {
    try { Unblock-File -LiteralPath $dllPath -ErrorAction SilentlyContinue } catch { }
}

# ---- 写入注册表 ----
$regKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"
$manifestUrl = 'file:///' + ($vstoPath -replace '\\', '/') + '|vstolocal'

if (-not (Test-Path -LiteralPath $regKey)) {
    New-Item -Path $regKey -Force | Out-Null
}

New-ItemProperty -Path $regKey -Name 'Description'  -Value $addInName    -PropertyType String   -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'FriendlyName' -Value $addInName    -PropertyType String   -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'LoadBehavior' -Value 3             -PropertyType DWord    -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'Manifest'     -Value $manifestUrl  -PropertyType String   -Force | Out-Null

Write-Host "[信息] 已写入注册表: $regKey"
Write-Host "[信息] Manifest = $manifestUrl"

# ---- 回读校验 ----
$check = Get-ItemProperty -Path $regKey
Write-Host ''
Write-Host '当前注册项:'
Write-Host ("  LoadBehavior = {0}" -f $check.LoadBehavior)
Write-Host ("  Manifest     = {0}" -f $check.Manifest)

Write-Host ''
Write-Host '[成功] 加载项已注册。请重启 Word 生效。' -ForegroundColor Green
Write-Host '       若 Word 中未出现「加载项」选项卡，请检查:' -ForegroundColor Gray
Write-Host '         1) HKCU\Software\Microsoft\VSTO\SolutionMetadata 下的错误记录' -ForegroundColor Gray
Write-Host '         2) 事件查看器 -> 应用程序日志 -> .NET Runtime' -ForegroundColor Gray

exit 0
