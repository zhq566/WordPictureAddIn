#requires -Version 3.0
<#
.SYNOPSIS
    注册 WordPictureAddIn（VSTO Word 加载项），无需管理员权限。

.DESCRIPTION
    1. 对输出目录下所有文件递归解除 Windows 文件锁定（Zone.Identifier）。
    2. 在 HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn 写入 VSTO 加载项注册项，
       Manifest 指向本地编译输出的 .vsto 文件并附加 |vstolocal 后缀，
       使 VSTO 运行时直接从该目录加载，而不经过 ClickOnce 缓存。
    3. 从 .vsto 清单中读取签名证书公钥，写入
       HKCU\Software\Microsoft\VSTO\Security\Inclusion\{GUID} 信任列表，
       避免 Word 因自签名证书不可信而拒绝加载。
    4. 清理 Office 14/15/16 中 Word 的「禁用项目」记录（Resiliency\DisabledItems），
       并写入 DoNotDisableAddinList，防止加载项被 Word 自动禁用。
    5. 检测 .NET Framework / VSTO 运行时 / Office 版本与位数，低于支持范围时明确提示。

.PARAMETER Target
    Debug（默认）、Release，或任意输出目录的绝对/相对路径。

.PARAMETER NoTrust
    跳过写入 VSTO 信任列表（HKCU\Software\Microsoft\VSTO\Security\Inclusion）。

.PARAMETER NoResiliency
    跳过清理 Word 弹性机制（DisabledItems / DoNotDisableAddinList）。

.EXAMPLE
    .\Register-AddIn.ps1
    .\Register-AddIn.ps1 Release
    .\Register-AddIn.ps1 "D:\build\bin\Debug"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Target = 'Debug',

    [switch]$NoTrust,

    [switch]$NoResiliency
)

$ErrorActionPreference = 'Stop'

# 保证在 cmd / 旧控制台窗口中中文提示不乱码
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$addInName = 'WordPictureAddIn'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir   = Split-Path -Parent $scriptDir

#region ---- 工具函数 ----

function Write-Step {
    param([string]$Text)
    Write-Host "[信息] $Text"
}

function Write-Warn {
    param([string]$Text)
    Write-Host "[警告] $Text" -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Text)
    Write-Host "[错误] $Text" -ForegroundColor Red
}

# 在指定注册表路径中按优先顺序查找第一个存在且非空的属性值
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

#endregion

#region ---- 解析输出目录 ----

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

Write-Step "输出目录: $outDir"

$vstoPath     = Join-Path $outDir "$addInName.vsto"
$manifestPath = Join-Path $outDir "$addInName.dll.manifest"
$dllPath      = Join-Path $outDir "$addInName.dll"

if (-not (Test-Path -LiteralPath $vstoPath)) {
    Write-Err "找不到 `"$vstoPath`""
    Write-Err '请先编译项目，或确认路径参数是否正确。'
    exit 1
}
if (-not (Test-Path -LiteralPath $manifestPath)) {
    Write-Err "找不到 `"$manifestPath`""
    exit 1
}

#endregion

#region ---- 环境检测 ----

# --- .NET Framework 4.x ---
$ndpPaths = @(
    'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full'
)
$netRelease = Get-FirstPropertyValue -Paths $ndpPaths -Names @('Release')
$netVersion = Get-FirstPropertyValue -Paths $ndpPaths -Names @('Version')

if ($null -eq $netRelease) {
    Write-Warn '未检测到 .NET Framework 4.x，本加载项要求 4.7.2 或更高版本。'
} else {
    Write-Step (".NET Framework: {0}（Release={1}）" -f $netVersion, $netRelease)
    if ([int]$netRelease -lt 461808) {
        Write-Warn '.NET Framework 低于 4.7.2，加载项可能无法启动。'
        Write-Warn '下载地址: https://dotnet.microsoft.com/download/dotnet-framework/net472'
    }
}

# --- VSTO 运行时 ---
$vstoRuntimePaths = @(
    'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
)
$vstoRuntimeVersion = Get-FirstPropertyValue -Paths $vstoRuntimePaths -Names @('Version')

$vstoeeCandidates = @(
    (Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\vstoee.dll')
    (Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared\VSTO\vstoee.dll')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $vstoRuntimeVersion -and $vstoeeCandidates.Count -eq 0) {
    Write-Warn '未检测到 VSTO 运行时（vstoee.dll），Word 无法加载此加载项。'
    Write-Warn '请安装 "Microsoft Visual Studio 2010 Tools for Office Runtime"：'
    Write-Warn 'https://www.microsoft.com/download/details.aspx?id=56961'
} else {
    if ($vstoRuntimeVersion) { Write-Step "VSTO 运行时: $vstoRuntimeVersion" }
    if ($vstoeeCandidates.Count -eq 0) {
        Write-Warn '注册表中存在 VSTO 运行时记录，但未找到 vstoee.dll，安装可能已损坏。'
    }
}

# --- Office 版本与位数 ---
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

$wordExe = Get-FirstPropertyValue `
    -Paths @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
    ) `
    -Names @('(default)', 'Path')

if (-not $officeMajor -and $wordExe) {
    $exeFile = Get-Item -LiteralPath $wordExe -ErrorAction SilentlyContinue
    if ($exeFile) {
        $fileVersion = $exeFile.VersionInfo.FileVersion
        $match = [regex]::Match("$fileVersion", '^(\d+)\.')
        if ($match.Success) {
            $officeMajor = [int]$match.Groups[1].Value
            if (-not $officeVersion) { $officeVersion = $fileVersion }
        }
    }
}

# 含 Office 主线版本 -> 产品名映射
$officeNameMap = @{
    12 = 'Office 2007'
    14 = 'Office 2010'
    15 = 'Office 2013'
    16 = 'Office 2016 / 2019 / 2021 / 365'
}

if ($null -ne $officeMajor) {
    $officeName = $officeNameMap[$officeMajor]
    if (-not $officeName) { $officeName = "Office $officeMajor.x" }
    $bitnessText = if ($officePlatform) { "$officePlatform" } else { '未知' }
    Write-Step ("Office 版本: {0}（{1}，位数: {2}）" -f $officeName, $officeVersion, $bitnessText)

    if ($officeMajor -lt 14) {
        Write-Warn '本加载项要求 Word 2010 或更高版本，当前 Office 版本不受支持。'
        Write-Warn 'Word 2007(12.0) 及更早版本不支持 VSTO 4.0 加载项。'
    }
} else {
    Write-Warn '未能确定 Office 版本，若加载失败请手工确认 Word 版本（需 2010 及以上）。'
}

# --- 说明：本加载项使用 EmbedInteropTypes，无需安装 Office PIA ---
Write-Host '[信息] 本加载项已内嵌互操作类型，无需单独安装 Office 主互操作程序集(PIA)。' -ForegroundColor Gray

#endregion

#region ---- 递归解除文件锁定 ----

$unblockedCount = 0
try {
    $outputFiles = Get-ChildItem -LiteralPath $outDir -Recurse -File -ErrorAction SilentlyContinue

    foreach ($file in $outputFiles) {
        try {
            # 仅处理带 Zone.Identifier 备用数据流的文件
            $zone = Get-Content -LiteralPath $file -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
            if ($zone) {
                Unblock-File -LiteralPath $file -ErrorAction SilentlyContinue
                $unblockedCount++
            }
        } catch { }
    }

    if ($unblockedCount -gt 0) {
        Write-Step "已解除 $unblockedCount 个文件的网络来源锁定(Mark of the Web)。"
    } else {
        Write-Step '输出目录中没有需要解锁的文件。'
    }
} catch {
    Write-Warn "解除文件锁定失败: $($_.Exception.Message)"
}

# 主 dll 兜底，防止上面按流检测有遗漏
if (Test-Path -LiteralPath $dllPath) {
    try { Unblock-File -LiteralPath $dllPath -ErrorAction SilentlyContinue } catch { }
}

#endregion

#region ---- 写入加载项注册表 ----

$regKey      = "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName"
$manifestUrl = 'file:///' + ($vstoPath -replace '\\', '/') + '|vstolocal'

if (-not (Test-Path -LiteralPath $regKey)) {
    New-Item -Path $regKey -Force | Out-Null
}

New-ItemProperty -Path $regKey -Name 'Description'  -Value $addInName    -PropertyType String   -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'FriendlyName' -Value $addInName    -PropertyType String   -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'LoadBehavior' -Value 3             -PropertyType DWord    -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'Manifest'     -Value $manifestUrl  -PropertyType String   -Force | Out-Null
New-ItemProperty -Path $regKey -Name 'CommandLineSafe' -Value 0          -PropertyType DWord    -Force | Out-Null

Write-Step "已写入注册表: $regKey"
Write-Step "Manifest = $manifestUrl"

#endregion

#region ---- 写入 VSTO 信任列表（免证书信任提示） ----

if ($NoTrust) {
    Write-Host '[信息] 已按 -NoTrust 跳过 VSTO 信任列表写入。' -ForegroundColor Gray
} else {
    $trustUrl = 'file:///' + ($vstoPath -replace '\\', '/')

    $publicKey = $null
    try {
        [xml]$vstoXml = Get-Content -LiteralPath $vstoPath -Raw
        $rsaNode = $vstoXml.SelectSingleNode("//*[local-name()='RSAKeyValue']")
        if ($rsaNode) { $publicKey = $rsaNode.OuterXml }
    } catch {
        $publicKey = $null
    }

    if ([string]::IsNullOrWhiteSpace($publicKey)) {
        Write-Warn '未能从 .vsto 清单读取签名公钥，跳过 VSTO 信任列表写入。'
        Write-Warn '若 Word 提示"此应用程序的发布者无法验证"，请手动将证书导入到"受信任的发布者"。'
    } else {
        $inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
        if (-not (Test-Path -LiteralPath $inclusionRoot)) {
            New-Item -Path $inclusionRoot -Force | Out-Null
        }

        # 清理指向同一 Url 的旧信任项，避免累积
        Get-ChildItem -Path $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and $entry.Url -eq $trustUrl) {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop
                    Write-Step "已清理旧的信任项: $($_.PSChildName)"
                } catch { }
            }
        }

        $trustKey = Join-Path $inclusionRoot ([guid]::NewGuid().ToString('B'))
        New-Item -Path $trustKey -Force | Out-Null
        New-ItemProperty -Path $trustKey -Name 'Url'       -Value $trustUrl -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $trustKey -Name 'PublicKey' -Value $publicKey -PropertyType String -Force | Out-Null

        Write-Step "已写入 VSTO 信任列表: $trustKey"
    }
}

#endregion

#region ---- 清理 Word 弹性机制（禁用项目 / 禁止自动禁用） ----

if ($NoResiliency) {
    Write-Host '[信息] 已按 -NoResiliency 跳过 Word 弹性机制清理。' -ForegroundColor Gray
} else {
    $resiliencyVersions = @('14.0', '15.0', '16.0')
    $removedDisabledCount = 0

    foreach ($version in $resiliencyVersions) {
        $resiliencyKey = "HKCU:\Software\Microsoft\Office\$version\Word\Resiliency"
        if (-not (Test-Path -LiteralPath $resiliencyKey)) { continue }

        # --- 从 DisabledItems 二进制值中剔除本加载项记录 ---
        $disabledItems = $null
        try {
            $disabledItems = (Get-ItemProperty -Path $resiliencyKey -Name 'DisabledItems' -ErrorAction Stop).DisabledItems
        } catch {
            $disabledItems = $null
        }

        if ($disabledItems -is [byte[]] -and $disabledItems.Length -gt 0) {
            $kept = New-Object System.Collections.Generic.List[byte]
            $offset = 0
            $changed = $false

            while ($offset + 4 -le $disabledItems.Length) {
                $entryLength = [System.BitConverter]::ToInt32($disabledItems, $offset)
                if ($entryLength -le 0 -or ($offset + 4 + $entryLength) -gt $disabledItems.Length) { break }

                $entryText = [System.Text.Encoding]::Unicode.GetString($disabledItems, $offset + 4, $entryLength)
                if ($entryText -like "*$addInName*") {
                    $changed = $true
                } else {
                    for ($i = 0; $i -lt (4 + $entryLength); $i++) {
                        $kept.Add($disabledItems[$offset + $i])
                    }
                }
                $offset += 4 + $entryLength
            }

            if ($offset -lt $disabledItems.Length) {
                for ($i = $offset; $i -lt $disabledItems.Length; $i++) { $kept.Add($disabledItems[$i]) }
            }

            if ($changed) {
                try {
                    if ($kept.Count -eq 0) {
                        Remove-ItemProperty -Path $resiliencyKey -Name 'DisabledItems' -Force -ErrorAction Stop
                    } else {
                        Set-ItemProperty -Path $resiliencyKey -Name 'DisabledItems' -Value $kept.ToArray() -Type Binary -Force -ErrorAction Stop
                    }
                    $removedDisabledCount++
                    Write-Step "已从 Word $version 的「禁用项目」中移除 $addInName"
                } catch {
                    Write-Warn "清理 Word $version 禁用项目失败: $($_.Exception.Message)"
                }
            }
        }

        # --- 写入 DoNotDisableAddinList ---
        $doNotDisableKey = Join-Path $resiliencyKey 'DoNotDisableAddinList'
        try {
            if (-not (Test-Path -LiteralPath $doNotDisableKey)) {
                New-Item -Path $doNotDisableKey -Force | Out-Null
            }
            New-ItemProperty -Path $doNotDisableKey -Name $addInName -Value 1 -PropertyType DWord -Force | Out-Null
            Write-Step "已写入 Word $version 的 DoNotDisableAddinList\$addInName"
        } catch {
            Write-Warn "写入 Word $version DoNotDisableAddinList 失败: $($_.Exception.Message)"
        }
    }

    if ($removedDisabledCount -eq 0) {
        Write-Step '未发现需要清理的「禁用项目」记录。'
    }
}

#endregion

#region ---- 回读校验 ----

$check = Get-ItemProperty -Path $regKey
Write-Host ''
Write-Host '当前注册项:'
Write-Host ("  LoadBehavior = {0}" -f $check.LoadBehavior)
Write-Host ("  Manifest     = {0}" -f $check.Manifest)

$wordProcess = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
if ($wordProcess) {
    Write-Host ''
    Write-Warn '检测到 Word 正在运行，请完全退出 Word 后重新打开，注册才会生效。'
}

Write-Host ''
Write-Host '[成功] 加载项已注册。' -ForegroundColor Green
Write-Host '       后续步骤:' -ForegroundColor Gray
Write-Host '         1) 完全关闭所有 Word 窗口（含后台残留 WINWORD.EXE），再重新打开 Word' -ForegroundColor Gray
Write-Host '         2) 若仍未加载，运行 scripts\Diagnose-AddIn.ps1 查看具体原因' -ForegroundColor Gray
Write-Host '         3) 手动排查: 文件 -> 选项 -> 加载项 -> 管理:禁用的项目' -ForegroundColor Gray

#endregion

exit 0