#requires -Version 3.0
<#
.SYNOPSIS
    诊断 WordPictureAddIn（VSTO Word 加载项）为什么没能加载。

.DESCRIPTION
    只读检查以下内容并打印结果，不会修改系统状态：
      1. 加载项注册表项（LoadBehavior / Manifest）
      2. VSTO 信任列表（Security\Inclusion）
      3. Word 弹性机制（Resiliency\DisabledItems / DoNotDisableAddinList）
      4. VSTO 解决方案元数据中的错误记录（SolutionMetadata）
      5. 环境信息：.NET Framework、VSTO 运行时、Office 版本与位数
      6. 输出目录文件及其网络来源锁定（Zone.Identifier）状态
      7. 加载项自身日志（%LOCALAPPDATA%\WordPictureAddIn\addin.log）

    使用 -EnableAlerts 时会额外把 VSTO 的诊断环境变量写入当前用户，
    让下次启动 Word 时弹出真实错误框并生成 .vsto.log（属修改状态，可选）。

.PARAMETER Target
    Debug（默认）、Release，或任意输出目录的绝对/相对路径。

.PARAMETER EnableAlerts
    设置 VSTO_SUPPRESSDISPLAYALERTS=0、VSTO_LOGALERTS=1，使 VSTO 弹出详细错误。

.PARAMETER DisableAlerts
    还原为 VSTO_SUPPRESSDISPLAYALERTS=1、VSTO_LOGALERTS=0（关闭弹窗与日志）。

.EXAMPLE
    .\Diagnose-AddIn.ps1
    .\Diagnose-AddIn.ps1 Release
    .\Diagnose-AddIn.ps1 Release -EnableAlerts
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Target = 'Debug',

    [switch]$EnableAlerts,

    [switch]$DisableAlerts
)

$ErrorActionPreference = 'Continue'

# 保证在 cmd / 旧控制台窗口中中文提示不乱码
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$addInName = 'WordPictureAddIn'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir   = Split-Path -Parent $scriptDir

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor DarkGray
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor DarkGray
}

function Write-Ok {
    param([string]$Text)
    Write-Host "  [正常] $Text" -ForegroundColor Green
}

function Write-Bad {
    param([string]$Text)
    Write-Host "  [异常] $Text" -ForegroundColor Red
}

function Write-Note {
    param([string]$Text)
    Write-Host "  [说明] $Text" -ForegroundColor Gray
}

function Get-FirstPropertyValue {
    param(
        [string[]]$Paths,
        [string[]]$Names
    )

    foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }

        $item = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }

        foreach ($name in $Names) {
            $value = $null
            try { $value = $item.$name } catch { $value = $null }
            if ($null -ne $value -and "$value" -ne '') { return $value }
        }
    }

    return $null
}

# 从 DisabledItems 二进制中提取所有可读条目文本
function Get-DisabledItemTexts {
    param([byte[]]$Data)

    $results = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Data -or $Data.Length -eq 0) { return $results }

    # 直接按 Unicode 解码后按不可打印字符切分，再从文本块中筛选含 dll/vsto/exe 的片段
    $text = [System.Text.Encoding]::Unicode.GetString($Data)
    foreach ($chunk in ($text -split "[^\u0020-\u007E]{3,}")) {
        $clean = ($chunk -replace "[\u0000-\u001F]", '').Trim()
        if ($clean.Length -ge 4 -and $clean -match '(?i)(\.dll|\.vsto|\.exe|WordPicture)') {
            if (-not $results.Contains($clean)) { $results.Add($clean) }
        }
    }

    return $results
}

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

Write-Host ''
Write-Host "WordPictureAddIn 诊断报告" -ForegroundColor White
Write-Host ("生成时间: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Gray
Write-Host ("计算机名: {0}    用户: {1}" -f $env:COMPUTERNAME, $env:USERNAME) -ForegroundColor Gray
Write-Host ("输出目录: {0}" -f $outDir) -ForegroundColor Gray

#region ---- 1. 加载项注册表项 ----

Write-Section '1. 加载项注册表项'

$regKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"
if (-not (Test-Path -LiteralPath $regKey)) {
    Write-Bad "注册项不存在: $regKey"
    Write-Note '请先运行 scripts\Register-AddIn.ps1 注册加载项。'
} else {
    $addIn = Get-ItemProperty -Path $regKey
    Write-Host ("  LoadBehavior = {0}" -f $addIn.LoadBehavior)
    Write-Host ("  Manifest     = {0}" -f $addIn.Manifest)
    Write-Host ("  FriendlyName = {0}" -f $addIn.FriendlyName)

    switch ($addIn.LoadBehavior) {
        3 { Write-Ok 'LoadBehavior = 3（开机自动加载）' }
        2 { Write-Bad 'LoadBehavior = 2，加载项已被 Word 置为「不加载」。这是加载失败后 Word 的典型回写，重新运行注册脚本即可。' }
        0 { Write-Bad 'LoadBehavior = 0，加载项已连接但未加载。' }
        1 { Write-Note 'LoadBehavior = 1，仅按需加载。' }
        default { Write-Bad ("LoadBehavior = {0}，非预期值。" -f $addIn.LoadBehavior) }
    }

    if ($addIn.Manifest -and $addIn.Manifest -like '*|vstolocal') {
        Write-Ok 'Manifest 带有 |vstolocal 后缀（从本地目录加载）。'
    } else {
        Write-Note 'Manifest 未带 |vstolocal，可能走 ClickOnce 缓存加载。'
    }
}

#endregion

#region ---- 2. VSTO 信任列表 ----

Write-Section '2. VSTO 信任列表（Security\Inclusion）'

$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
$vstoPath = Join-Path $outDir "$addInName.vsto"
$expectedUrl = 'file:///' + ($vstoPath -replace '\\', '/')

if (-not (Test-Path -LiteralPath $inclusionRoot)) {
    Write-Bad '不存在 Security\Inclusion 项，自签名证书的加载项可能因此被拒绝加载。'
    Write-Note '解决方案: 运行 scripts\Register-AddIn.ps1（会自动写入信任项）。'
} else {
    $found = $false
    Get-ChildItem -Path $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
        $entry = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        if ($entry -and "$($entry.Url)" -like "*$addInName*") {
            $found = $true
            Write-Host ("  匹配项: {0}" -f $_.PSChildName)
            Write-Host ("    Url = {0}" -f $entry.Url)
            if ("$($entry.Url)" -eq $expectedUrl) {
                Write-Ok '信任项 Url 与当前输出目录一致。'
            } else {
                Write-Note '信任项 Url 与当前输出目录不一致（换目录后需重新注册）。'
            }
            if ([string]::IsNullOrWhiteSpace("$($entry.PublicKey)")) {
                Write-Bad '信任项缺少 PublicKey。'
            } else {
                Write-Ok '信任项包含 PublicKey。'
            }
        }
    }
    if (-not $found) {
        Write-Bad "未找到指向 $addInName 的信任项。"
        Write-Note '解决方案: 重新运行 scripts\Register-AddIn.ps1。'
    }
}

#endregion

#region ---- 3. Word 弹性机制 ----

Write-Section '3. Word 弹性机制（禁用项目 / 禁止自动禁用）'

$resiliencyVersions = @('14.0', '15.0', '16.0')
$disabledFound = $false

foreach ($version in $resiliencyVersions) {
    $resiliencyKey = "HKCU:\Software\Microsoft\Office\$version\Word\Resiliency"
    $hasAny = $false

    if (Test-Path -LiteralPath $resiliencyKey) {
        $resiliency = Get-ItemProperty -Path $resiliencyKey -ErrorAction SilentlyContinue

        if ($resiliency -and $null -ne $resiliency.DisabledItems) {
            $hasAny = $true
            $texts = Get-DisabledItemTexts -Data $resiliency.DisabledItems
            $hit = @($texts | Where-Object { $_ -like "*$addInName*" })

            if ($hit.Count -gt 0) {
                $disabledFound = $true
                Write-Bad ("Word {0} 的「禁用项目」中包含本加载项:" -f $version)
                foreach ($t in $hit) { Write-Host "      - $t" -ForegroundColor Red }
                Write-Note "解决方案: 运行 scripts\Register-AddIn.ps1 自动清理，或 Word 中 文件->选项->加载项->管理:禁用的项目 手动启用。"
            } else {
                Write-Ok ("Word {0} 的「禁用项目」中不含本加载项（共 {1} 条其他记录）。" -f $version, $texts.Count)
            }
        }

        $doNotDisableKey = Join-Path $resiliencyKey 'DoNotDisableAddinList'
        if (Test-Path -LiteralPath $doNotDisableKey) {
            $hasAny = $true
            $dnd = Get-ItemProperty -Path $doNotDisableKey -Name $addInName -ErrorAction SilentlyContinue
            if ($dnd) {
                Write-Ok ("Word {0} 的 DoNotDisableAddinList 已包含本加载项。" -f $version)
            } else {
                Write-Note ("Word {0} 的 DoNotDisableAddinList 未包含本加载项。" -f $version)
            }
        }
    }

    if (-not $hasAny) {
        Write-Note ("Word {0} 无弹性机制记录（说明该版本未禁用过加载项，或从未运行）。" -f $version)
    }
}

#endregion

#region ---- 4. VSTO 解决方案元数据 ----

Write-Section '4. VSTO 错误记录（SolutionMetadata）'

$vstoMeta = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
if (-not (Test-Path -LiteralPath $vstoMeta)) {
    Write-Note 'SolutionMetadata 不存在（通常表示 VSTO 没有记录到加载失败）。'
} else {
    $metaItems = @(Get-ChildItem -Path $vstoMeta -ErrorAction SilentlyContinue)
    if ($metaItems.Count -eq 0) {
        Write-Note 'SolutionMetadata 为空。'
    } else {
        # SolutionMetadata 只是 VSTO 记录的解决方案清单快照（含 CLR 版本、目标框架等），
        # 存在本身不代表失败；只有出现 Error / ErrorReport 之类的属性才是真正的错误记录。
        $errorPropertyNames = @('Error', 'ErrorReport', 'Exception', 'ErrorMessages', 'FailureReason')
        $matchedMeta = 0
        $realErrors = 0

        foreach ($item in $metaItems) {
            $props = Get-ItemProperty -Path $item.PSPath -ErrorAction SilentlyContinue
            if (-not $props) { continue }

            $isMine = ("$($props.addInName)" -eq $addInName) -or
                      ("$($props.description)" -like "*$addInName*") -or
                      ("$($item.PSChildName)" -like "*$addInName*")
            if (-not $isMine) { continue }

            $matchedMeta++

            $errorValues = @()
            foreach ($name in $errorPropertyNames) {
                $v = $null
                try { $v = $props.$name } catch { $v = $null }
                if ($v) { $errorValues += ("{0}={1}" -f $name, $v) }
            }

            $summary = "{0}（addInName={1}, loadBehavior={2}）" -f `
                $item.PSChildName, $props.addInName, $props.loadBehavior

            if ($errorValues.Count -gt 0) {
                $realErrors++
                Write-Bad $summary
                foreach ($e in $errorValues) { Write-Host "      $e" -ForegroundColor Red }
            } else {
                Write-Host "  $summary" -ForegroundColor DarkGray
            }
        }

        if ($matchedMeta -eq 0) {
            Write-Note 'SolutionMetadata 中没有本加载项的记录。'
        } elseif ($realErrors -eq 0) {
            Write-Ok 'SolutionMetadata 中的本加载项记录均为清单快照，未包含错误信息。'
        }
    }
}

#endregion

#region ---- 5. 环境信息 ----

Write-Section '5. 环境信息'

$ndpPaths = @(
    'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full'
)
$netRelease = Get-FirstPropertyValue -Paths $ndpPaths -Names @('Release')
$netVersion = Get-FirstPropertyValue -Paths $ndpPaths -Names @('Version')

if ($null -eq $netRelease) {
    Write-Bad '未检测到 .NET Framework 4.x。'
} else {
    $netText = ".NET Framework {0}（Release={1}）" -f $netVersion, $netRelease
    if ([int]$netRelease -ge 461808) {
        Write-Ok $netText
    } else {
        Write-Bad "$netText —— 低于 4.7.2，加载项无法启动。"
        Write-Note '下载: https://dotnet.microsoft.com/download/dotnet-framework/net472'
    }
}

$vstoRuntimePaths = @(
    'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
)
$vstoRuntimeVersion = Get-FirstPropertyValue -Paths $vstoRuntimePaths -Names @('Version')

$vstoeeCandidates = @(
    (Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\vstoee.dll')
    (Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared\VSTO\vstoee.dll')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if ($vstoRuntimeVersion) {
    Write-Ok ("VSTO 运行时: {0}" -f $vstoRuntimeVersion)
} elseif ($vstoeeCandidates.Count -gt 0) {
    Write-Ok '检测到 vstoee.dll，但注册表无版本记录。'
} else {
    Write-Bad '未检测到 VSTO 运行时，Word 无法加载 VSTO 加载项。'
    Write-Note '下载: https://www.microsoft.com/download/details.aspx?id=56961'
}

foreach ($dll in $vstoeeCandidates) {
    Write-Note ("vstoee.dll: {0}" -f $dll)
}

$c2rConfigPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
)
$officeVersion  = Get-FirstPropertyValue -Paths $c2rConfigPaths -Names @('VersionToReport', 'Version')
$officePlatform = Get-FirstPropertyValue -Paths $c2rConfigPaths -Names @('Platform')

$officeMajor = $null
if ($officeVersion) {
    $match = [regex]::Match($officeVersion, '^(\d+)\.')
    if ($match.Success) { $officeMajor = [int]$match.Groups[1].Value }
}

if ($null -ne $officeMajor) {
    if ($officeMajor -lt 14) {
        Write-Bad ("Office 主版本 {0}.x —— 低于 Word 2010，不受支持。" -f $officeMajor)
    } else {
        Write-Ok ("Office 主版本 {0}.x（{1}，位数: {2}）" -f $officeMajor, $officeVersion, $(if ($officePlatform) { $officePlatform } else { '未知' }))
    }
} else {
    Write-Note '未从 ClickToRun 配置读取到 Office 版本（零售/MSI 安装时属正常）。'
}

$wordPath = Get-FirstPropertyValue `
    -Paths @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
    ) `
    -Names @('(default)', 'Path')

if ($wordPath -and (Test-Path -LiteralPath $wordPath)) {
    $exeFile = Get-Item -LiteralPath $wordPath
    Write-Host ("  WINWORD.EXE = {0}" -f $wordPath)
    Write-Host ("  文件版本    = {0}" -f $exeFile.VersionInfo.FileVersion)
} else {
    Write-Note '未从 App Paths 找到 WINWORD.EXE 路径。'
}

$wordProcess = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
if ($wordProcess) {
    Write-Note '当前有 Word 进程在运行（检查前请完全退出 Word，否则看到的可能是旧状态）。'
}

#endregion

#region ---- 6. 输出目录文件与锁定状态 ----

Write-Section '6. 输出目录文件与网络来源锁定'

if (-not (Test-Path -LiteralPath $outDir)) {
    Write-Bad "输出目录不存在: $outDir"
} else {
    $requiredFiles = @(
        "$addInName.dll"
        "$addInName.dll.manifest"
        "$addInName.vsto"
        'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
    )

    foreach ($name in $requiredFiles) {
        $full = Join-Path $outDir $name
        if (Test-Path -LiteralPath $full) {
            Write-Ok ("存在: {0}" -f $name)
        } else {
            Write-Bad ("缺失: {0}" -f $name)
        }
    }

    $locked = @()
    Get-ChildItem -LiteralPath $outDir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $zone = Get-Content -LiteralPath $_.FullName -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
            if ($zone) { $locked += $_.FullName }
        } catch { }
    }

    if ($locked.Count -gt 0) {
        Write-Bad ("有 {0} 个文件带网络来源锁定(Zone.Identifier)，VSTO 会拒绝加载：" -f $locked.Count)
        foreach ($f in $locked) {
            Write-Host ("      - {0}" -f (Split-Path -Leaf $f)) -ForegroundColor Red
        }
        Write-Note '解决方案: 运行 scripts\Register-AddIn.ps1（会递归解除锁定），或对所有文件右键->属性->勾选「解除锁定」。'
    } else {
        Write-Ok '输出目录中没有带网络来源锁定的文件。'
    }
}

#endregion

#region ---- 7. 加载项自身日志 ----

Write-Section '7. 加载项自身日志'

$logPath = Join-Path $env:LOCALAPPDATA 'WordPictureAddIn\addin.log'
if (-not (Test-Path -LiteralPath $logPath)) {
    Write-Note "日志不存在: $logPath"
    Write-Note '说明加载项从未成功启动过；请重点检查上面 1~6 项。'
} else {
    $logInfo = Get-Item -LiteralPath $logPath
    Write-Host ("  路径      = {0}" -f $logPath)
    Write-Host ("  最后写入  = {0}" -f $logInfo.LastWriteTime)
    Write-Host ("  大小      = {0} 字节" -f $logInfo.Length)
    Write-Host ''
    Write-Host '  最近 40 行:' -ForegroundColor Gray
    Get-Content -LiteralPath $logPath -Tail 40 -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host ("    {0}" -f $_) -ForegroundColor DarkGray
    }
    Write-Note ('如出现「[ERROR]」行，即为加载失败的根因，可用记事本打开完整日志查看。')
}

#endregion

#region ---- 8. 可选：开启 VSTO 详细错误提示 ----

Write-Section '8. VSTO 错误提示开关'

if ($EnableAlerts -and $DisableAlerts) {
    Write-Bad '不能同时使用 -EnableAlerts 与 -DisableAlerts，已跳过。'
} elseif ($EnableAlerts) {
    try {
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', '0', 'User')
        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', '1', 'User')
        Write-Ok '已开启 VSTO 详细错误提示（写入当前用户环境变量）。'
        Write-Note 'VSTO_SUPPRESSDISPLAYALERTS = 0（启动失败时弹出错误框）'
        Write-Note 'VSTO_LOGALERTS = 1（在加载项目录生成 .vsto.log）'
        Write-Note '必须完全退出 Word 并重新打开，新环境变量才会生效。'
        Write-Note '排查完成后可用 -DisableAlerts 关闭。'
    } catch {
        Write-Bad ("设置环境变量失败: {0}" -f $_.Exception.Message)
    }
} elseif ($DisableAlerts) {
    try {
        [Environment]::SetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', '1', 'User')
        [Environment]::SetEnvironmentVariable('VSTO_LOGALERTS', '0', 'User')
        Write-Ok '已关闭 VSTO 详细错误提示。'
    } catch {
        Write-Bad ("设置环境变量失败: {0}" -f $_.Exception.Message)
    }
} else {
    $curSuppress = [Environment]::GetEnvironmentVariable('VSTO_SUPPRESSDISPLAYALERTS', 'User')
    $curLog      = [Environment]::GetEnvironmentVariable('VSTO_LOGALERTS', 'User')
    Write-Host ("  VSTO_SUPPRESSDISPLAYALERTS = {0}" -f $(if ($null -eq $curSuppress) { '(未设置，默认 1 即静默失败)' } else { $curSuppress }))
    Write-Host ("  VSTO_LOGALERTS             = {0}" -f $(if ($null -eq $curLog) { '(未设置，默认 0)' } else { $curLog }))
    Write-Note '想看到真实错误信息，可执行: .\Diagnose-AddIn.ps1 -EnableAlerts'
}

#endregion

Write-Host ''
Write-Host ('=' * 72) -ForegroundColor DarkGray
Write-Host ' 诊断完成。' -ForegroundColor Cyan
if ($disabledFound) {
    Write-Host ' 结论: 加载项曾被 Word 自动禁用，请先运行 Register-AddIn.ps1 清理禁用记录。' -ForegroundColor Yellow
}
Write-Host ('=' * 72) -ForegroundColor DarkGray
Write-Host ''

exit 0