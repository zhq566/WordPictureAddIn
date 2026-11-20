#requires -Version 3.0
<#
.SYNOPSIS
    卸载 WordPictureAddIn（VSTO Word 加载项）。

.DESCRIPTION
    1. 删除 HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn 注册项。
    2. 删除 HKCU\Software\Microsoft\VSTO\Security\Inclusion 中指向本加载项的信任项。
    3. 清理 Office 14/15/16 中 Word 的 DoNotDisableAddinList\WordPictureAddIn。
    4. 从 NGen 本机映像缓存移除加载项（需管理员权限，未提权时自动跳过）。

    加 -Force 可跳过 Word 运行中的确认提示。

.PARAMETER Force
    跳过 Word 运行中的确认提示。

.PARAMETER KeepTrust
    保留 VSTO 信任列表项（更换版本时想复用证书信任时可指定）。

.PARAMETER Target
    Release（默认）、Debug，或任意输出目录的绝对/相对路径；用于定位要卸载 NGen 的程序集。

.PARAMETER NoNGen
    跳过 NGen 缓存清理。
#>
[CmdletBinding()]
param(
    [switch]$Force,

    [switch]$KeepTrust,

    [Parameter(Position = 0)]
    [string]$Target = 'Release',

    [switch]$NoNGen
)

$ErrorActionPreference = 'Stop'

# 保证在 cmd / 旧控制台窗口中中文提示不乱码
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$addInName = 'WordPictureAddIn'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir   = Split-Path -Parent $scriptDir
$regKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"

# NGen 辅助函数（Get-WordPlatform / Uninstall-AddInNGen 等）
. (Join-Path $scriptDir 'NGen-Common.ps1')

# ---- 解析输出目录（用于定位 NGen 缓存中的程序集） ----
switch ($Target) {
    'Release' { $outDir = Join-Path $rootDir 'bin\Release' }
    'Debug'   { $outDir = Join-Path $rootDir 'bin\Debug' }
    default {
        if ([System.IO.Path]::IsPathRooted($Target)) {
            $outDir = $Target
        } else {
            $outDir = Join-Path $rootDir $Target
        }
    }
}
$outDir = [System.IO.Path]::GetFullPath($outDir).TrimEnd('\')
$dllPath = Join-Path $outDir "$addInName.dll"

# ---- 检查 Word 是否在运行 ----
$word = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
if ($word -and -not $Force) {
    Write-Host '[警告] 检测到 Word 正在运行，请先完全关闭 Word。' -ForegroundColor Yellow
    Write-Host '       否则加载项文件可能被锁定，且 Word 退出时可能重写注册项。' -ForegroundColor Yellow
    $answer = Read-Host '仍然继续吗? (y/N)'
    if ($answer -notmatch '^[Yy]') {
        Write-Host '已取消。'
        exit 1
    }
}

# ---- 删除加载项注册表项 ----
if (Test-Path -LiteralPath $regKey) {
    Remove-Item -Path $regKey -Recurse -Force
    if (Test-Path -LiteralPath $regKey) {
        Write-Host "[错误] 删除注册项失败: $regKey" -ForegroundColor Red
        exit 1
    }
    Write-Host "[成功] 已删除注册项: $regKey" -ForegroundColor Green
} else {
    Write-Host "[信息] 注册项不存在，无需删除: $regKey"
}

# ---- 清理 VSTO 信任列表 ----
if ($KeepTrust) {
    Write-Host '[信息] 已按 -KeepTrust 保留 VSTO 信任列表项。' -ForegroundColor Gray
} else {
    $inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (Test-Path -LiteralPath $inclusionRoot) {
        $removedTrust = 0
        Get-ChildItem -Path $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            $matchesAddIn = $false
            if ($entry) {
                if ("$($entry.Url)" -like "*$addInName*") { $matchesAddIn = $true }
            }
            if ($matchesAddIn) {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop
                    $removedTrust++
                    Write-Host "[成功] 已删除 VSTO 信任项: $($_.PSChildName)" -ForegroundColor Green
                } catch {
                    Write-Host "[警告] 删除信任项失败: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        if ($removedTrust -eq 0) {
            Write-Host '[信息] 未发现指向本加载项的 VSTO 信任项。'
        }
    }
}

# ---- 清理 DoNotDisableAddinList ----
$resiliencyVersions = @('14.0', '15.0', '16.0')
foreach ($version in $resiliencyVersions) {
    $doNotDisableKey = "HKCU:\Software\Microsoft\Office\$version\Word\Resiliency\DoNotDisableAddinList"
    if (-not (Test-Path -LiteralPath $doNotDisableKey)) { continue }

    $value = Get-ItemProperty -Path $doNotDisableKey -Name $addInName -ErrorAction SilentlyContinue
    if ($value) {
        try {
            Remove-ItemProperty -Path $doNotDisableKey -Name $addInName -Force -ErrorAction Stop
            Write-Host "[成功] 已清理 Word $version 的 DoNotDisableAddinList\$addInName" -ForegroundColor Green
        } catch {
            Write-Host "[警告] 清理 Word $version DoNotDisableAddinList 失败: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ---- 清理 NGen 本机映像缓存 ----
if ($NoNGen) {
    Write-Host '[信息] 已按 -NoNGen 跳过 NGen 缓存清理。' -ForegroundColor Gray
} elseif (-not (Test-IsAdministrator)) {
    Write-Host '[信息] 当前不是管理员权限，跳过 NGen 缓存清理。' -ForegroundColor Gray
    Write-Host '       NGen 缓存占用很小，残留不影响使用；如需彻底清理请用管理员身份运行。' -ForegroundColor Gray
} else {
    $wordPlatform = Get-WordPlatform
    if ($wordPlatform) {
        Uninstall-AddInNGen -AssemblyPath $dllPath -Platform $wordPlatform | Out-Null
    } else {
        Write-Host '[信息] 无法确定 Word 位数，跳过 NGen 缓存清理。' -ForegroundColor Gray
    }
}

# ---- 清理 VSTO 解决方案元数据（仅提示，不自动删除） ----
$vstoMeta = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
if (Test-Path -LiteralPath $vstoMeta) {
    Write-Host "[信息] 如需彻底清理 VSTO 缓存，可手动删除: $vstoMeta" -ForegroundColor Gray
}

Write-Host ''
Write-Host '[成功] 加载项已卸载。重启 Word 后「加载项」选项卡应消失。' -ForegroundColor Green
exit 0