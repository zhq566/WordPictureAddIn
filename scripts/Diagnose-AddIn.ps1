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
      6. 加载项启动耗时、NGen 本机映像状态、证书吊销检查状态
      7. 输出目录文件及其网络来源锁定（Zone.Identifier）状态
      8. 加载项自身日志（%LOCALAPPDATA%\WordPictureAddIn\addin.log）

    使用 -EnableAlerts 时会额外把 VSTO 的诊断环境变量写入当前用户，
    让下次启动 Word 时弹出真实错误框并生成 .vsto.log（属修改状态，可选）。

.PARAMETER Target
    Release（默认）、Debug，或任意输出目录的绝对/相对路径。

.PARAMETER EnableAlerts
    设置 VSTO_SUPPRESSDISPLAYALERTS=0、VSTO_LOGALERTS=1，使 VSTO 弹出详细错误。

.PARAMETER DisableAlerts
    还原为 VSTO_SUPPRESSDISPLAYALERTS=1、VSTO_LOGALERTS=0（关闭弹窗与日志）。

.EXAMPLE
    .\Diagnose-AddIn.ps1
    .\Diagnose-AddIn.ps1 Debug
    .\Diagnose-AddIn.ps1 Release -EnableAlerts
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Target = 'Release',

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

#region ---- 6. 加载项启动性能 ----

Write-Section '6. 加载项启动性能'

$dllPath = Join-Path $outDir "$addInName.dll"
$perfLogPath = Join-Path $env:LOCALAPPDATA 'WordPictureAddIn\addin.log'
$startupAverage = $null

# ---- 6.1 加载项日志中的计时记录 ----
$startupSamples  = New-Object System.Collections.Generic.List[double]
$prewarmSamples  = New-Object System.Collections.Generic.List[double]

if (-not (Test-Path -LiteralPath $perfLogPath)) {
    Write-Note '暂无 addin.log，无法统计启动耗时。'
} else {
    $perfLines = Get-Content -LiteralPath $perfLogPath -Tail 600 -ErrorAction SilentlyContinue
    foreach ($line in $perfLines) {
        $m = [regex]::Match($line, 'Startup 耗时\s*([0-9]+(?:\.[0-9]+)?)\s*ms')
        if ($m.Success) {
            $startupSamples.Add([double]$m.Groups[1].Value)
            continue
        }
        $m2 = [regex]::Match($line, '预热贴图窗口完成，耗时\s*([0-9]+(?:\.[0-9]+)?)\s*ms')
        if ($m2.Success) { $prewarmSamples.Add([double]$m2.Groups[1].Value) }
    }

    if ($startupSamples.Count -eq 0) {
        Write-Note '日志中没有计时记录（可能是升级前写入的旧日志，重启一次 Word 后再看）。'
    } else {
        $recent = @($startupSamples | Select-Object -Last 5)
        $startupAverage = ($recent | Measure-Object -Average).Average
        $startupMax     = ($recent | Measure-Object -Maximum).Maximum
        $perfText = '最近 {0} 次 Startup 耗时: 平均 {1:F1} ms，最大 {2:F1} ms' -f $recent.Count, $startupAverage, $startupMax

        if ($startupAverage -le 400) {
            Write-Ok $perfText
        } else {
            Write-Note $perfText
        }

        if ($startupAverage -gt 1500) {
            Write-Bad 'Startup 耗时偏高（平均 > 1500 ms）。请结合下面 NGen 与证书吊销检查结论定位原因。'
        }
    }

    if ($prewarmSamples.Count -gt 0) {
        $prewarmAverage = ($prewarmSamples | Measure-Object -Average).Average
        Write-Note ('贴图窗口预热耗时: 平均 {0:F1} ms（首次选中图片后才触发，不计入 Word 启动时间）' -f $prewarmAverage)
    } else {
        Write-Note '尚无贴图窗口预热记录（本次会话还没选中过图片）。'
    }
}

# ---- 6.2 Word 记录的加载项加载耗时 ----
$loadTimeKeys = @()
foreach ($ver in @('16.0', '15.0', '14.0')) {
    $loadTimeKeys += "HKCU:\Software\Microsoft\Office\$ver\Word\Addins\$addInName"
}
$loadTimeKeys += "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"

$loadTimeFound = $false
foreach ($key in $loadTimeKeys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    if ($null -eq $item) { continue }

    foreach ($prop in $item.PSObject.Properties) {
        if ($prop.Name -match '(?i)^loadtime') {
            Write-Host ("  Word 记录的加载耗时: {0} = {1}" -f $prop.Name, $prop.Value)
            $loadTimeFound = $true
        }
    }
}

if (-not $loadTimeFound) {
    Write-Note 'Word 尚未记录本加载项的 LoadTime（首次加载后才会写入）。'
}

# ---- 6.3 NGen 本机映像状态 ----
$ngenCommonPath = Join-Path $scriptDir 'NGen-Common.ps1'
$wordPlatform = $null

if (Test-Path -LiteralPath $ngenCommonPath) {
    try {
        . $ngenCommonPath
        $wordPlatform = Get-WordPlatform
    } catch {
        Write-Note ('加载 NGen-Common.ps1 失败: {0}' -f $_.Exception.Message)
    }
}

if (-not $wordPlatform) {
    Write-Note '无法确定 Word 的位数，跳过 NGen 状态检查。'
} else {
    Write-Host ("  Word 位数 = {0}" -f $wordPlatform)
    $ngenExe = Get-NGenPath -Platform $wordPlatform

    if (-not $ngenExe) {
        $frameworkDirName = if ($wordPlatform -eq 'x86') { 'Framework' } else { 'Framework64' }
        Write-Bad ('未找到 ngen.exe（期望: %WINDIR%\Microsoft.NET\{0}\v4.0.30319\ngen.exe）' -f $frameworkDirName)
    } else {
        Write-Host ("  ngen.exe  = {0}" -f $ngenExe)

        if (-not (Test-Path -LiteralPath $dllPath)) {
            Write-Note ("程序集不存在，无法检查 NGen 状态: {0}" -f $dllPath)
        } else {
            $ngenInstalled = $false
            $ngenChecked = $false

            $previousPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                $displayOutput = & $ngenExe display $dllPath /nologo 2>&1
                $ngenChecked = $true
                if (($displayOutput -join "`n") -match '(?i)WordPictureAddIn') { $ngenInstalled = $true }
            } catch {
            } finally {
                $ErrorActionPreference = $previousPreference
            }

            if (-not $ngenChecked) {
                Write-Note '无法查询 NGen 缓存（可能需要管理员权限）。'
            } elseif ($ngenInstalled) {
                Write-Ok '本机映像(NGen)已安装 —— CLR 加载时无需 JIT 编译。'
                Write-Note '程序集更新后需重新运行 Register-AddIn.ps1 才会刷新映像。'
            } else {
                Write-Bad '本机映像(NGen)未安装，Word 启动时需要 JIT 编译加载项。'
                Write-Note '解决: 以管理员身份运行 scripts\Register-AddIn.ps1（内部会执行 ngen install）。'
            }
        }
    }
}

if ($Target -eq 'Debug' -or $outDir -match '(?i)[\\/]bin[\\/]debug$') {
    Write-Bad '当前诊断的是 Debug 配置产物：JIT 不做优化，Word 启动会明显变慢。'
    Write-Note '日常使用请改用 Release：以管理员身份运行 scripts\Register-AddIn.ps1（默认 Release）。'
} else {
    Write-Ok '当前为 Release 配置产物。'
}

# ---- 6.4 证书签名与吊销检查 ----
# VSTO 校验的是 .vsto 清单的签名，所以优先检查清单，缺失时再退回检查 dll。
$signTargets = @(
    (Join-Path $outDir "$addInName.vsto")
    $dllPath
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if ($signTargets.Count -eq 0) {
    Write-Note '输出目录缺少 .vsto 与 dll，跳过签名检查。'
} else {
    $signTarget = $signTargets[0]
    Write-Host ("  签名对象 = {0}" -f (Split-Path -Leaf $signTarget))

    $signature = $null
    try { $signature = Get-AuthenticodeSignature -LiteralPath $signTarget -ErrorAction SilentlyContinue } catch { }

    if ($null -eq $signature) {
        Write-Note '无法读取 Authenticode 签名信息。'
    } elseif ($null -eq $signature.SignerCertificate) {
        Write-Ok '未签名 —— 不涉及证书吊销检查，启动不会被 CRL/OCSP 阻塞。'
    } else {
        Write-Host ("  签名主体 = {0}" -f $signature.SignerCertificate.Subject)
        Write-Host ("  签名状态 = {0}" -f $signature.Status)

        if ("$($signature.Status)" -eq 'Valid') {
            Write-Ok '签名有效，证书吊销检查已完成。'
        } else {
            Write-Bad ("签名状态为 {0}：VSTO 可能因证书校验失败或超时而拒绝加载。" -f $signature.Status)
            Write-Note '内网/离线环境下，CRL/OCSP 查询会一直等到超时，是 Word 启动卡顿十几秒的常见原因。'
        }
    }
}

$wintrustKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing'
$wintrustState = Get-FirstPropertyValue -Paths @($wintrustKey) -Names @('State')

if ($null -eq $wintrustState) {
    Write-Note '未读取到 WinTrust 吊销检查设置（使用系统默认值 0x23E00）。'
} else {
    $stateValue = 0
    [void][int]::TryParse("$wintrustState", [ref]$stateValue)

    if ($stateValue -eq 0x23C00) {
        Write-Bad '证书吊销检查已被关闭（State = 0x23C00）：启动会更快，但不再校验证书是否被吊销。'
        Write-Note ('恢复默认: Set-ItemProperty -Path "{0}" -Name State -Value 0x23E00' -f $wintrustKey)
    } else {
        Write-Note ('WinTrust State = 0x{0:X}（默认 0x23E00 表示启用证书吊销检查）。' -f $stateValue)
        Write-Note '若内网/离线环境启动卡顿十几秒，可临时设为 0x23C00 对比验证（属降低安全性，排查后请恢复）。'
    }
}

#endregion

#region ---- 7. 输出目录文件与锁定状态 ----

Write-Section '7. 输出目录文件与网络来源锁定'

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

#region ---- 8. 加载项自身日志 ----

Write-Section '8. 加载项自身日志'

$logPath = Join-Path $env:LOCALAPPDATA 'WordPictureAddIn\addin.log'
if (-not (Test-Path -LiteralPath $logPath)) {
    Write-Note "日志不存在: $logPath"
    Write-Note '说明加载项从未成功启动过；请重点检查上面 1~7 项。'
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

#region ---- 9. 可选：开启 VSTO 详细错误提示 ----

Write-Section '9. VSTO 错误提示开关'

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
if ($null -ne $startupAverage -and $startupAverage -gt 1500) {
    Write-Host (' 结论: Startup 平均耗时 {0:F0} ms，明显偏高；请按第 6 节提示安装 NGen 并检查证书吊销检查。' -f $startupAverage) -ForegroundColor Yellow
}
Write-Host ('=' * 72) -ForegroundColor DarkGray
Write-Host ''

exit 0