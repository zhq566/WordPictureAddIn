# 注册 / 卸载 / 诊断脚本

`Register-AddIn.cmd`、`Unregister-AddIn.cmd`、`Diagnose-AddIn.cmd` 是对应 PowerShell 脚本的薄封装，
真正的逻辑在 `.ps1` 中（带 UTF-8 BOM，保证中文提示不乱码）。
全部脚本只写 `HKCU`，**不需要管理员权限**。

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
scripts\Register-AddIn.cmd Release
```

默认注册 `bin\Debug`。也可以指定配置或绝对路径：

```cmd
scripts\Register-AddIn.cmd
scripts\Register-AddIn.cmd "D:\some\path\bin\Debug"
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

注册完成后**完全退出 Word**（包括任务管理器里残留的 `WINWORD.EXE`）再重新打开。Word 中会出现「加载项」选项卡。

## 诊断

Word 里加载项显示为「已禁用」、或者根本没有出现时，先跑诊断脚本：

```cmd
scripts\Diagnose-AddIn.cmd Release
```

它只读不写，会依次打印：

1. 加载项注册项（`LoadBehavior` / `Manifest`），并对 `LoadBehavior=2`（被 Word 置为不加载）给出提示
2. VSTO 信任列表匹配情况（`Url` 是否与当前目录一致、是否含 `PublicKey`）
3. Word 各版本的 `DisabledItems` 中是否含本加载项，`DoNotDisableAddinList` 是否已写入
4. `HKCU\Software\Microsoft\VSTO\SolutionMetadata` 中的记录，并区分「清单快照」与「真实错误」
5. 环境信息：.NET Framework、VSTO 运行时、Office 版本与位数、`WINWORD.EXE` 路径与文件版本
6. 输出目录必需文件是否存在、是否有文件带网络来源锁定
7. 加载项自身日志 `%LOCALAPPDATA%\WordPictureAddIn\addin.log` 的最近 40 行
8. VSTO 错误提示开关状态

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

会删除加载项注册项、指向本加载项的 VSTO 信任项，以及各 Office 版本的
`DoNotDisableAddinList\WordPictureAddIn`。脚本会先检查 `WINWORD.EXE` 是否在运行并提示确认。

| 参数 | 作用 |
|---|---|
| `-Force` | 跳过 Word 运行中的确认提示 |
| `-KeepTrust` | 保留 VSTO 信任项（换版本但复用同一证书时可指定） |

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
- **提示清单签名无效**：本项目使用自签名证书。若换机器，需要重新生成 `WordPictureAddIn_TemporaryKey.pfx`，
  并把证书指纹写入 [`WordPictureAddIn.csproj`](../WordPictureAddIn.csproj) 的 `ManifestCertificateThumbprint`，
  然后重新编译（`Rebuild`），否则清单与 dll 的哈希不匹配。
- **修改代码后**：卸载脚本并非必须，直接重新编译并重启 Word 即可；但如果换了输出目录，需要重新注册。
- **卸载后重启 Word，加载项又回来了**：说明 Word 在退出时把 `LoadBehavior` 又写回了 `3`（VSTO 的正常行为）。
  确认注册表项已删除即可，如果反复出现，检查是否还残留 `HKLM\Software\Microsoft\Office\Word\Addins\WordPictureAddIn`。
- **切换 Debug/Release 目录后加载项失效**：`bin\Debug` 与 `bin\Release` 是两份独立清单，
  信任项中会同时存在两条记录，属正常现象；只需对目标目录重新执行注册脚本。
