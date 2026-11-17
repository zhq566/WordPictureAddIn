#requires -Version 3.0
<#
.SYNOPSIS
    卸载 WordPictureAddIn（VSTO Word 加载项）。

.DESCRIPTION
    删除 HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn 注册项。
    加 -Force 可跳过 Word 运行中的确认提示。
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$addInName = 'WordPictureAddIn'
$regKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"

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

# ---- 删除注册表项 ----
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

# ---- 清理 VSTO 解决方案元数据（仅提示，不自动删除） ----
$vstoMeta = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
if (Test-Path -LiteralPath $vstoMeta) {
    Write-Host "[信息] 如需彻底清理 VSTO 缓存，可手动删除: $vstoMeta" -ForegroundColor Gray
}

Write-Host ''
Write-Host '[成功] 加载项已卸载。重启 Word 后「加载项」选项卡应消失。' -ForegroundColor Green
exit 0
