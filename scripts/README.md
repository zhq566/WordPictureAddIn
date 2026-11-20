# 注册 / 卸载 / 诊断脚本

`Register-AddIn.cmd`、`Unregister-AddIn.cmd`、`Diagnose-AddIn.cmd` 是对应 PowerShell 脚本的薄封装，
真正的逻辑在 `.ps1` 中（带 UTF-8 BOM，保证中文提示不乱码）。
注册与卸载只写 `HKCU`，**不需要管理员权限**；只有可选的 NGen 预编译步骤需要管理员权限
（不提升时脚本会自动跳过，不影响加载项功能）。

这些脚本用于在没有 Visual Studio 调试会话（F5）的情况下，手动把编译产物注册到 Word。

## 兼容性

| 项目 | 要求 |
|---|---|
| Word | **Word 2010 及以上**（2010 / 2013 / 2016 / 2019 / 2021 / 365），32 位与 64 位均可 |
| Word 2007 及更早 | **不支持**，无法加载 VSTO 4.0 加载项 |
| WPS Office | **不支持**，WPS 不加载 VSTO 加载项 |
| Windows | Windows 7 SP1 及以上 |
| .NET Framework | **4.7.2 及以上**（项目目标框架即为 v4.7.2，无需 4.8） |
| VSTO 运行时 | Microsoft Visual Studio 2010 Tools for Office Runtime |
| Office PIA | **不需要**，项目已设置 `EmbedInteropTypes=true` |

> 本加载项为 `AnyCPU`，同一份 `bin` 输出可直接用于 32 位和 64 位 Office。

## 前提

1. 已安装 **.NET Framework 4.7.2** 及以上。
2. 已安装 **Microsoft Visual Studio 2010 Tools for Office Runtime**（即 VSTO 运行时，提供 `vstoee.dll`）。
   脚本会检测 `vstoee.dll` 与注册表 `HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R`，缺失时给出警告。
3. 已编译项目（**分发请用 `Release`**），`bin\Release` 下存在
   `WordPictureAddIn.vsto`、`WordPictureAddIn.dll.manifest` 和 `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`。

## 注册

```cmd
scripts\Register-AddIn.cmd
```

默认注册 **`bin\Release`**（Release 产物已关闭调试符号与序列化程序集生成，启动更快）。
也可以指定配置或绝对路径：

```cmd
scripts\Register-AddIn.cmd Debug
scripts\Register-AddIn.cmd "D:\some\path\bin\Release"
```

脚本执行以下动作：

1. **递归解除文件锁定**：对输出目录下所有带 `Zone.Identifier` 的文件执行 `Unblock-File`。
   从网络/压缩包拷贝过来的文件会带「网络来源」标记，VSTO 会直接拒绝加载，这是换机器部署失败最常见的原因。
2. **写入加载项注册项**（见下表）。
3. **写入 VSTO 信任列表**：从 `WordPictureAddIn.vsto` 中读取签名证书公钥（`<RSAKeyValue>` 节点），
   写入 `HKCU\Software\Microsoft\VSTO\Security\Inclusion\{新GUID}` 的 `Url` 与 `PublicKey`。
   这样自签名证书无需导入「受信任的发布者」也不会被拦。会先清理指向同一 `Url` 的旧信任项。
4. **清理 Word 弹性机制**：从 Office 14.0 / 15.0 / 16.0 的
   `Word\Resiliency\DisabledItems` 二进制值中剔除本加载项记录，并在
   `Word\Resiliency\DoNotDisableAddinList` 写入 `WordPictureAddIn = 1`，防止 Word 再次自动禁用。
5. **环境检测**：打印 .NET Framework 版本、VSTO 运行时版本、Office 版本与位数；
   检测到 Office 低于 2010 时明确提示「不受支持」。
6. **NGen 预编译**（可选，需管理员权限）：按 Word 的位数选择 `Framework` / `Framework64`
   下的 `ngen.exe`，对 `WordPictureAddIn.dll` 执行 `ngen install`，生成的原生映像可省去
   Word 启动时的 JIT 编译。位数判断错误会导致「装了也没用」，脚本会先读
   `ClickToRun\Configuration\Platform`，再回退到解析 `WINWORD.EXE` 的 PE 头。

写入的加载项键值（`HKCU`，**不需要管理员权限**）：

| 值名 | 类型 | 内容 |
|---|---|---|
| `Description` | REG_SZ | `WordPictureAddIn` |
| `FriendlyName` | REG_SZ | `WordPictureAddIn` |
| `LoadBehavior` | REG_DWORD | `3`（启动时加载 + 按需加载） |
| `Manifest` | REG_SZ | `file:///<绝对路径>/WordPictureAddIn.vsto\|vstolocal` |
| `CommandLineSafe` | REG_DWORD | `0` |

位置：`HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn`

`|vstolocal` 后缀告诉 VSTO 运行时直接从本地目录加载，不经过 ClickOnce 缓存。

可用参数：

| 参数 | 作用 |
|---|---|
| `-NoTrust` | 跳过第 3 步（不写 VSTO 信任列表） |
| `-NoResiliency` | 跳过第 4 步（不动 Word 弹性机制） |
| `-NoNGen` | 跳过第 6 步（不做 NGen 预编译） |

> **重新编译后请重新运行一次注册脚本**：NGen 映像是按程序集版本绑定的，
> DLL 更新后旧映像会失效，需要重新 `ngen install` 才能继续生效。

注册完成后**完全退出 Word**（包括任务管理器里残留的 `WINWORD.EXE`）再重新打开。Word 中会出现「加载项」选项卡。

## 诊断

Word 里加载项显示为「已禁用」、或者根本没有出现时，先跑诊断脚本：

```cmd
scripts\Diagnose-AddIn.cmd Release
```

它只读不写（除 `-EnableAlerts` / `-DisableAlerts` 会写环境变量），会依次打印：

1. 加载项注册项（`LoadBehavior` / `Manifest`），并对 `LoadBehavior=2`（被 Word 置为不加载）给出提示
2. VSTO 信任列表匹配情况（`Url` 是否与当前目录一致、是否含 `PublicKey`）
3. Word 各版本的 `DisabledItems` 中是否含本加载项，`DoNotDisableAddinList` 是否已写入
4. `HKCU\Software\Microsoft\VSTO\SolutionMetadata` 中的记录，并区分「清单快照」与「真实错误」
5. 环境信息：.NET Framework、VSTO 运行时、Office 版本与位数、`WINWORD.EXE` 路径与文件版本
6. **启动性能**：`addin.log` 中最近几次 `Startup 耗时` 与贴图窗口预热耗时的均值/最大值、
   Word 记录的 `LoadTime`、NGen 本机映像是否已安装（并校验位数是否与 Word 一致）、
   程序集签名状态与 WinTrust 证书吊销检查设置
7. 输出目录必需文件是否存在、是否有文件带网络来源锁定
8. 加载项自身日志 `%LOCALAPPDATA%\WordPictureAddIn\addin.log` 的最近 40 行
   （该日志**只保留最近一次启动**的内容，每次打开 Word 都会覆盖上一次的记录）
9. VSTO 错误提示开关状态

> 默认目标目录已改为 `bin\Release`；要诊断 Debug 产物请显式传 `Debug`。

想看到 VSTO 的真实报错（默认是静默失败，什么都看不到）：

```cmd
scripts\Diagnose-AddIn.cmd Release -EnableAlerts
```

这会写入用户环境变量 `VSTO_SUPPRESSDISPLAYALERTS=0`、`VSTO_LOGALERTS=1`，
之后**完全重启 Word**，VSTO 会弹出详细错误框并在加载项目录生成 `.vsto.log`。
排查完用 `-DisableAlerts` 关掉即可。

## 卸载

```cmd
scripts\Unregister-AddIn.cmd
```

会删除加载项注册项、指向本加载项的 VSTO 信任项、各 Office 版本的
`DoNotDisableAddinList\WordPictureAddIn`，并从 NGen 本机映像缓存中移除
`WordPictureAddIn.dll`（需管理员权限，未提升时自动跳过）。脚本会先检查
`WINWORD.EXE` 是否在运行并提示确认。

| 参数 | 作用 |
|---|---|
| `-Force` | 跳过 Word 运行中的确认提示 |
| `-KeepTrust` | 保留 VSTO 信任项（换版本但复用同一证书时可指定） |
| `-NoNGen` | 保留 NGen 本机映像（不清理） |
| `-Target` | 指定要清理映像的输出目录，默认 `Release` |

## 启动慢排查

加载项功能正常、但打开 Word 明显变慢时，按下面顺序定位：

1. **先跑基线**：`scripts\Diagnose-AddIn.cmd` 看第 6 节。
   `addin.log` 中的 `Startup 耗时` 是加载项自身 `ThisAddIn_Startup` 的耗时，
   与 Word 整体启动时间分开统计，便于判断问题是否出在加载项内。
2. **确认用的是 Release**：Debug 产物带 `DebuggableAttribute`，JIT 不做优化，
   启动通常慢一倍以上。用 `scripts\Register-AddIn.cmd`（默认 Release）重新注册。
3. **安装 NGen 本机映像**：以管理员身份运行
   `scripts\Register-AddIn.cmd`，脚本会按 Word 位数选择正确的 `ngen.exe`。
   诊断报告第 6 节若显示「本机映像(NGen)未安装」，说明这一步没生效。
4. **检查证书吊销检查（CRL/OCSP）**：VSTO 在加载前会验证清单签名，
   离线或内网环境若无法访问吊销服务器，会一直等到 TCP 超时才继续，
   表现为「Word 启动卡十几秒到一分钟」。诊断报告会给出程序集签名状态与
   `WinTrust\Trust Providers\Software Publishing\State`。

   临时关闭吊销检查以验证（**属降低安全性，仅用于定位**）：

   ```powershell
   Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing' -Name State -Value 0x23C00
   ```

   验证完请恢复默认值：

   ```powershell
   Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing' -Name State -Value 0x23E00
   ```

   更推荐的做法是给证书链补齐 CRL 分发点，或改用内网可访问的证书，
   而不是长期关闭吊销检查。
5. **WPF 相关加载已延后**：贴图窗口（WPF）只在首次选中图片时才初始化，
   不再占用 Word 启动时间。日志中的「预热贴图窗口完成，耗时 xx ms」
   属于用户首次选中图片后的额外开销，**不计入 Word 启动**。

## 常见问题

- **Word 里看不到「加载项」选项卡，也没有任何提示**：这是 VSTO 的默认静默失败行为。按顺序排查：
  1. `scripts\Diagnose-AddIn.cmd Release` 看结论
  2. 检查 `LoadBehavior` 是否被 Word 改成了 `2`
  3. 检查输出目录文件是否有「网络来源」锁定（右键属性看是否有「解除锁定」勾选项）
  4. 检查 `VSTO\Security\Inclusion` 信任项 `Url` 是否与当前目录一致
  5. 检查 `Resiliency\DisabledItems` 是否含本加载项
  6. 用 `-EnableAlerts` 让 VSTO 弹出真实错误
- **加载项出现在 COM 加载项列表，但被 Word 自动禁用**：通常是加载过程抛异常。
  本项目已把 `ThisAddIn_Startup` 整体包在 `try/catch` 中，异常只记日志、不再向外抛，
  因此不会再被 Word 判定为「加载失败」而禁用。查看 `%LOCALAPPDATA%\WordPictureAddIn\addin.log` 定位原因。
- **日志里看不到上一次的问题记录**：日志只保留最近一次启动的内容，重新打开 Word 就会覆盖。
  排查时请在复现问题后**先看日志、再关 Word**，或者先把 `addin.log` 复制出来留存。
- **提示清单签名无效**：本项目使用自签名证书。若换机器，需要重新生成 `WordPictureAddIn_TemporaryKey.pfx`，
  并把证书指纹写入 [`WordPictureAddIn.csproj`](../WordPictureAddIn.csproj) 的 `ManifestCertificateThumbprint`，
  然后重新编译（`Rebuild`），否则清单与 dll 的哈希不匹配。
- **修改代码后**：卸载脚本并非必须，直接重新编译并重启 Word 即可；但如果换了输出目录，需要重新注册。
- **卸载后重启 Word，加载项又回来了**：说明 Word 在退出时把 `LoadBehavior` 又写回了 `3`（VSTO 的正常行为）。
  确认注册表项已删除即可，如果反复出现，检查是否还残留 `HKLM\Software\Microsoft\Office\Word\Addins\WordPictureAddIn`。
- **切换 Debug/Release 目录后加载项失效**：`bin\Debug` 与 `bin\Release` 是两份独立清单，
  信任项中会同时存在两条记录，属正常现象；只需对目标目录重新执行注册脚本。
