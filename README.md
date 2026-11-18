# WordPictureAddIn

桌面版 Microsoft Word 图片双击预览插件，类似WPS的功能。

## 功能简介

- 在 Word 中双击图片，按原始像素尺寸贴在屏幕上（类似 Snipaste / PinPix 的贴图）；图片超过屏幕工作区时自动等比缩小
- 可以同时贴多张图，贴图置顶显示、失去焦点不会消失，用鼠标左键拖动移动位置
- 滚轮以光标为中心缩放贴图，工具栏「恢复原始尺寸」可回到 1:1
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
<img width="800" height="500" alt="Word界面" src="https://github.com/user-attachments/assets/c03625cd-b169-4075-bba3-9b32c0f556f5" />

<img width="800" height="500" alt="预览界面" src="https://github.com/user-attachments/assets/bcb44427-f1a3-42f5-baaa-94712c222dbd" />

## 安装和卸载
双击`setup.exe`即可安装，如遇到无法使用或者影响Word自身功能，通过设置里卸载。
<img width="611" height="234" alt="image" src="https://github.com/user-attachments/assets/659ad924-87c4-45ae-aeb3-22438ef77ed1" />
