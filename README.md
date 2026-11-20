# WordPictureAddIn

桌面版 Microsoft Word 图片双击预览插件，类似WPS的功能。

## 功能简介

- 在 Word 中双击图片，按原始像素尺寸贴在屏幕上（类似 Snipaste / PinPix 的贴图）；图片超过屏幕工作区时自动等比缩小
- 可以同时贴多张图，贴图置顶显示、失去焦点不会消失，用鼠标左键拖动移动位置
- 滚轮以光标为中心缩放贴图，工具栏「恢复原始尺寸」可回到 1:1
- 放大查看细节：贴图放大到超过屏幕时，窗口最大只占满当前显示器的工作区，图片不会被截断，可以在窗口内拖动查看其余部分（详见下方「放大与平移」）
- 鼠标悬停时在贴图下方显示工具栏，移开自动隐藏
- 支持临时标注(支持shift吸附)：
  - 直线
  - 矩形
  - 圆形
  - 箭头
  - 文字
- 选中标注工具后，工具栏下方自动展开一行参数设置（未选中工具时自动隐藏）：
  - 形状工具（直线/矩形/圆形/箭头）：8 个预设颜色 + 线宽（细/中/粗）
  - 文字工具：8 个预设颜色 + 字号（A− / A+ 逐档调整）
  - 颜色、线宽、字号的选择会被记住，下次打开贴图沿用上次设置
  - 已画好的标注保留各自的样式，修改参数只影响之后的新标注；文字编辑中改参数会立即应用到当前文字
- 支持撤销、切换预览底色（含透明棋盘格）、保存当前带标注的图片
- 标注内容只保留在对应贴图窗口中，关闭贴图后自动清空
- 关闭贴图：双击图片、按 Esc，或点工具栏的关闭按钮

### 放大与平移

| 操作 | 效果 |
|---|---|
| 滚轮 | 以光标所在位置为中心缩放，光标下的像素保持不动 |
| 左键拖动（图片未超出屏幕） | 移动整个贴图窗口，光标为 ✥ 移动样式 |
| 左键拖动（图片已超出屏幕） | 在窗口内平移图片，查看被遮住的部分，光标变为四向滚动样式；拖到图片边缘会自动停住，不会露出空白 |
| Ctrl + 左键拖动（图片已超出屏幕） | 仍然移动整个贴图窗口 |
| 工具栏「恢复原始尺寸」 | 回到 1:1，图片能完整显示时窗口自动恢复为与图片同样大小 |

- 横向、纵向分别判断：例如只有宽度超出屏幕时，只在横向平移，纵向窗口仍随图片大小变化
- 多显示器时按贴图当前所在的显示器计算可用区域
- 放大或平移状态下照常可以画标注、编辑文字；保存的图片始终是完整的原图加标注，不受当前缩放与平移位置影响
<img width="800" height="500" alt="Word界面" src="https://github.com/user-attachments/assets/c03625cd-b169-4075-bba3-9b32c0f556f5" />

<img width="800" height="500" alt="预览界面" src="https://github.com/user-attachments/assets/bcb44427-f1a3-42f5-baaa-94712c222dbd" />

## 兼容性

| 项目 | 要求 |
|---|---|
| Word | **Word 2010 及以上**（2010 / 2013 / 2016 / 2019 / 2021 / 365） |
| Office 位数 | 32 位与 64 位均可（加载项为 `AnyCPU`，同一份输出通用） |
| Word 2007 及更早 | **不支持**（无法加载 VSTO 4.0 加载项） |
| WPS Office | **不支持**（WPS 不加载 VSTO 加载项） |
| Windows | Windows 7 SP1 及以上 |
| .NET Framework | **4.7.2 及以上**（无需 4.8） |
| VSTO 运行时 | 需要 Microsoft Visual Studio 2010 Tools for Office Runtime |
| Office PIA | **不需要**（已内嵌互操作类型） |

## 安装和卸载
双击`setup.exe`即可安装，如遇到无法使用或者影响Word自身功能，通过设置里卸载。
<img width="611" height="234" alt="image" src="https://github.com/user-attachments/assets/659ad924-87c4-45ae-aeb3-22438ef77ed1" />

### 手动部署（无安装包时）
使用 [`scripts`](scripts) 目录下的脚本：

```cmd
scripts\Register-AddIn.cmd             :: 注册（默认 bin\Release）
scripts\Diagnose-AddIn.cmd             :: 加载失败或启动慢时先跑这个，只读诊断
scripts\Unregister-AddIn.cmd           :: 卸载
```

注册脚本会自动处理换机器部署最常见的三个坑：解除拷贝文件的「网络来源」锁定、
写入 VSTO 信任列表（避免自签名证书被拦）、清理 Word 的「禁用项目」记录。
脚本本身只写 `HKCU`，无需管理员权限；以管理员身份运行时还会额外执行
NGen 预编译（按 Word 位数自动选择 `Framework` / `Framework64` 的 `ngen.exe`），
不提升权限则自动跳过，不影响功能。

排查方法详见 [`scripts/README.md`](scripts/README.md)。

## 启动性能

加载项对 Word 启动时间的影响已做如下优化，功能行为不变：

| 优化点 | 说明 |
|---|---|
| WPF 延迟加载 | 贴图窗口（WPF）在**首次选中图片**时才初始化，Word 启动时不再加载 PresentationFramework / PresentationCore / WindowsBase |
| 类型隔离 | WPF 相关字段与方法全部移入 `PinManager`，`ThisAddIn` 不出现 WPF 类型，避免 CLR 在构建类型布局时连带加载 WPF 程序集 |
| 启动诊断延后 | `WriteStartupDiagnostics` 的若干 COM 调用推迟到启动后 2 秒执行，`ThisAddIn_Startup` 只写一行耗时日志 |
| 选区事件快速返回 | `WindowSelectionChange` 先用 `selection.Type` 判断，空白选区/光标状态下不再发起 `InlineShapes` / `ShapeRange` 跨进程调用 |
| 取消启动期预热 | 不再在 `ApplicationIdle` 建隐藏窗口；改为首次选中图片后延迟 400 ms 预热，贴图后与关闭后仍会再预热 |
| Release + NGen | 默认注册 Release 产物，并以管理员身份执行 `ngen install`，省去启动时的 JIT 编译 |

加载项自身耗时记录在 `%LOCALAPPDATA%\WordPictureAddIn\addin.log`：

- `加载项启动完成，Startup 耗时 xx ms` —— 每次打开 Word 一行
- `预热贴图窗口完成，耗时 xx ms` —— 首次选中图片后触发，不计入 Word 启动

如 Word 启动仍然很慢，请先运行 `scripts\Diagnose-AddIn.cmd`，其第 6 节会给出
NGen 本机映像状态、程序集签名状态与 WinTrust 证书吊销检查设置。
**离线/内网环境下 CRL/OCSP 查询超时是「Word 启动卡十几秒」的常见原因**，
具体处理方法见 [`scripts/README.md`](scripts/README.md) 的「启动慢排查」。
