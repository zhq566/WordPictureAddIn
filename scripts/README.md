# 测试用注册 / 卸载脚本

`Register-AddIn.cmd` / `Unregister-AddIn.cmd` 是对 PowerShell 脚本的薄封装，
真正的逻辑在 `Register-AddIn.ps1` / `Unregister-AddIn.ps1` 中（带 UTF-8 BOM，保证中文提示不乱码）。
两者都只写 `HKCU`，**不需要管理员权限**。

这两个脚本用于在没有 Visual Studio 调试会话（F5）的情况下，手动把编译产物注册到 Word。

## 前提

1. 已安装 **.NET Framework 4.7.2**。
2. 已安装 **Microsoft Visual Studio 2010 Tools for Office Runtime**（即 VSTO 运行时，提供 `vstoee.dll`）。
   脚本会检测 `%CommonProgramFiles%\microsoft shared\VSTO\vstoee.dll`，缺失时给出警告。
3. 已编译项目，`bin\Debug` 下存在 `WordPictureAddIn.vsto` 和 `WordPictureAddIn.dll.manifest`。

## 注册

```cmd
scripts\Register-AddIn.cmd
```

默认注册 `bin\Debug`。也可以指定配置或绝对路径：

```cmd
scripts\Register-AddIn.cmd Release
scripts\Register-AddIn.cmd "D:\some\path\bin\Debug"
```

脚本会写入下面的键（`HKCU`，**不需要管理员权限**）：

| 值名 | 类型 | 内容 |
|---|---|---|
| `Description` | REG_SZ | `WordPictureAddIn` |
| `FriendlyName` | REG_SZ | `WordPictureAddIn` |
| `LoadBehavior` | REG_DWORD | `3`（启动时加载 + 按需加载） |
| `Manifest` | REG_SZ | `file:///<绝对路径>/WordPictureAddIn.vsto\|vstolocal` |

位置：`HKCU\Software\Microsoft\Office\Word\Addins\WordPictureAddIn`

`|vstolocal` 后缀告诉 VSTO 运行时直接从本地目录加载，不经过 ClickOnce 缓存。

注册完成后**重启 Word**。Word 中会出现「加载项」选项卡。

## 卸载

```cmd
scripts\Unregister-AddIn.cmd
```

会删除上面那个注册表项。脚本会先检查 `WINWORD.EXE` 是否在运行并提示确认。

## 常见问题

- **Word 里看不到「加载项」选项卡**：先看注册表项是否还在；再从下面两处查加载失败原因：
  - `HKCU\Software\Microsoft\VSTO\SolutionMetadata`
  - 事件查看器 → 应用程序日志 → `.NET Runtime` 错误
- **提示清单签名无效**：本项目使用自签名证书。若换机器，需要重新生成 `WordPictureAddIn_TemporaryKey.pfx`，
  并把证书指纹写入 [`WordPictureAddIn.csproj`](../WordPictureAddIn.csproj) 的 `ManifestCertificateThumbprint`，
  然后重新编译（`Rebuild`），否则清单与 dll 的哈希不匹配。
- **修改代码后**：卸载脚本并非必须，直接重新编译并重启 Word 即可；但如果换了输出目录，需要重新注册。
- **卸载后重启 Word，加载项又回来了**：说明 Word 在退出时把 `LoadBehavior` 又写回了 `3`（VSTO 的正常行为）。
  确认注册表项已删除即可，如果反复出现，检查是否还残留 `HKLM\Software\Microsoft\Office\Word\Addins\WordPictureAddIn`。
