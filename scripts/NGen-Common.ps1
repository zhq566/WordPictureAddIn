#requires -Version 3.0
<#
.SYNOPSIS
    WordPictureAddIn 的 NGen 辅助函数（由 Register-AddIn.ps1 / Unregister-AddIn.ps1 点号引用）。

.DESCRIPTION
    NGen（Native Image Generator）会把 .NET 程序集预编译成本机代码并缓存到 GAC 旁，
    加载时不再需要 JIT 编译，可以明显缩短 Word 启动时 CLR 加载 VSTO 加载项的时间。

    关键点：NGen 缓存是「每 CLR 架构一份」的。加载项运行在 WINWORD.EXE 进程里，
    所以必须按 Word 的位数选择 Framework 或 Framework64 目录下的 ngen.exe，
    位数选错时 NGen 会报告成功但加载项仍然走 JIT，等于没装。

    这两个函数都需要管理员权限（写入 %WINDIR%\assembly 下的本机映像缓存）。
#>

<#
.SYNOPSIS
    检测 Word（WINWORD.EXE）的位数，返回 'x64'、'x86' 或 $null（未知）。
#>
function Get-WordPlatform {
    # 1) 优先读取 ClickToRun 配置，这是最可靠的来源
    $c2rPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )

    foreach ($path in $c2rPaths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $item = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }

        $platform = "$($item.Platform)"
        if ($platform -match '(?i)^x64$|^64') { return 'x64' }
        if ($platform -match '(?i)^x86$|^32') { return 'x86' }
    }

    # 2) 回退：从 App Paths 找到 winword.exe，读取 PE 头判断位数
    $appPathKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
    )

    $wordExe = $null
    foreach ($key in $appPathKeys) {
        if (-not (Test-Path -LiteralPath $key)) { continue }
        $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }
        $candidate = $item.'(default)'
        if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = $item.Path }
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            $wordExe = $candidate
            break
        }
    }

    if ($wordExe) {
        try {
            # 读 PE 头判断位数：Machine 值 0x8664 = x64，0x14C = x86
            # 偏移 0x3C 处是 PE 头偏移，PE 签名占 4 字节，紧随其后 2 字节即 Machine
            $stream = [System.IO.File]::OpenRead($wordExe)
            try {
                $reader = New-Object System.IO.BinaryReader($stream)
                $stream.Position = 0x3C
                $peOffset = $reader.ReadInt32()
                $stream.Position = $peOffset + 4
                $machine = $reader.ReadUInt16()
                if ($machine -eq 0x8664) { return 'x64' }
                if ($machine -eq 0x014C) { return 'x86' }
            } finally {
                $stream.Dispose()
            }
        } catch {
            return $null
        }
    }

    # 3) 最后回退：按操作系统位数猜（64 位 Office 远比 32 位常见，但不能确证）
    return $null
}

<#
.SYNOPSIS
    取得与加载项运行环境（Word 位数）匹配的 ngen.exe 路径。
#>
function Get-NGenPath {
    param([string]$Platform)

    $frameworkDirName = if ($Platform -eq 'x86') { 'Framework' } else { 'Framework64' }
    $base = Join-Path $env:WINDIR "Microsoft.NET\$frameworkDirName\v4.0.30319"
    $ngen = Join-Path $base 'ngen.exe'

    if (Test-Path -LiteralPath $ngen) { return $ngen }
    return $null
}

<#
.SYNOPSIS
    判断当前进程是否具备管理员权限。
#>
function Test-IsAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

<#
.SYNOPSIS
    把加载项程序集安装到 NGen 本机映像缓存。

.PARAMETER AssemblyPath
    要预编译的程序集完整路径（通常为输出目录下的 WordPictureAddIn.dll）。

.PARAMETER Platform
    'x64' 或 'x86'，决定使用哪一套 CLR 的 ngen.exe。

.OUTPUTS
    成功返回 $true，否则返回 $false。
#>
function Install-AddInNGen {
    param(
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][string]$Platform
    )

    $ngen = Get-NGenPath -Platform $Platform
    if (-not $ngen) {
        Write-Host "[警告] 未找到 .NET Framework v4.0.30319 的 ngen.exe，跳过预编译。" -ForegroundColor Yellow
        return $false
    }

    if (-not (Test-Path -LiteralPath $AssemblyPath)) {
        Write-Host "[警告] 找不到程序集，跳过预编译: $AssemblyPath" -ForegroundColor Yellow
        return $false
    }

    Write-Host "[信息] 正在预编译(NGen): $AssemblyPath"
    Write-Host "[信息] 使用: $ngen"

    # ngen.exe 是原生程序，失败时会往 stderr 写内容并返回非 0 退出码。
    # 调用方脚本可能设置了 $ErrorActionPreference = 'Stop'，这里必须局部放宽，
    # 否则 stderr 会被当成终止性错误，导致整个注册脚本中断。
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    try {
        # 先卸载旧映像，避免程序集更新后残留的旧版本映像被继续使用。
        # 首次安装时本来就没有旧映像，返回非 0 属正常，忽略即可。
        try {
            & $ngen uninstall $AssemblyPath /nologo 2>&1 | Out-Null
        } catch {
        }

        $output = & $ngen install $AssemblyPath /nologo 2>&1
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }

    if ($exitCode -ne 0) {
        Write-Host "[警告] NGen 预编译失败（退出码 $exitCode），加载项仍可正常运行，只是首次加载稍慢。" -ForegroundColor Yellow
        if ($output) { $output | ForEach-Object { Write-Host "       $_" -ForegroundColor DarkGray } }
        Write-Host "[提示] 常见原因：加载项 DLL 被 Word 占用，请先完全关闭 Word 后重试。" -ForegroundColor Gray
        return $false
    }

    Write-Host '[成功] NGen 预编译完成。' -ForegroundColor Green

    # WPF 程序集（PresentationFramework 等）通常已由系统预先 NGen，
    # 这里顺带确认一次，没有的话也不强制处理，因为首次选中图片时才用得到。
    return $true
}

<#
.SYNOPSIS
    从 NGen 本机映像缓存中移除加载项程序集。

.PARAMETER AssemblyPath
    程序集完整路径。

.PARAMETER Platform
    'x64' 或 'x86'。
#>
function Uninstall-AddInNGen {
    param(
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][string]$Platform
    )

    $ngen = Get-NGenPath -Platform $Platform
    if (-not $ngen) {
        return $false
    }

    if (-not (Test-Path -LiteralPath $AssemblyPath)) {
        Write-Host "[信息] 程序集已不存在，无需从 NGen 缓存移除: $AssemblyPath" -ForegroundColor Gray
        return $false
    }

    # 同 Install-AddInNGen：局部放宽错误处理，防止 stderr 中断调用方脚本。
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'

    try {
        $output = & $ngen uninstall $AssemblyPath /nologo 2>&1
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }

    if ($exitCode -ne 0) {
        # 程序集本来就没装过 NGen 时也会返回非 0，不必视为错误
        Write-Host '[信息] 该程序集不在 NGen 缓存中，或移除时被占用。' -ForegroundColor Gray
        return $false
    }

    Write-Host '[成功] 已从 NGen 缓存移除加载项。' -ForegroundColor Green
    return $true
}
