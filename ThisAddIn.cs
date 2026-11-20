using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Interop.Word;

namespace WordPictureAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>选中图片后延迟抓图的防抖间隔，避免每次单击选中都阻塞 Word 界面线程。</summary>
        private const int SelectionCaptureDelayMilliseconds = 260;

        /// <summary>启动诊断信息的延迟写入时间，避开 Word 启动的关键路径。</summary>
        private const int StartupDiagnosticsDelayMilliseconds = 2000;

        /// <summary>
        /// 首次选中图片后延迟多久才去初始化 WPF 并预热贴图窗口。
        /// 留出这段间隔，用户的选中操作先完成，WPF 加载发生在随后的空闲时间。
        /// </summary>
        private const int PinPrewarmDelayMilliseconds = 400;

        /// <summary>
        /// 贴图窗口管理器。此处刻意使用延迟创建：
        /// 本类在 Word 启动时就会被加载，如果字段直接引用 WPF 类型（窗口、Dispatcher 等），
        /// CLR 在构建类型布局时就会连锁加载 PresentationFramework / PresentationCore / WindowsBase，
        /// 这些程序集加上 XAML 解析和贴图窗口类自身的 JIT 是启动慢的主要来源之一。
        /// 改为用户第一次真正选中图片时才创建管理器，启动路径上不触达任何 WPF 类型。
        /// </summary>
        private PinManager _pinManager;

        private string _cachedSelectionKey;
        private Image _cachedSelectionImage;
        private string _pendingSelectionKey;
        private bool _isShuttingDown;
        private bool _startupDiagnosticsWritten;

        /// <summary>是否已经请求过一次贴图窗口预热，避免每次选中图片都重新排队。</summary>
        private bool _pinPrewarmRequested;

        /// <summary>
        /// 防抖计时器使用 WinForms 定时器而不是 WPF 的 DispatcherTimer：
        /// 前者只依赖 Word 本来就有的消息循环，不需要先初始化 WPF Application。
        /// </summary>
        private System.Windows.Forms.Timer _selectionCaptureTimer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 整个启动过程包在 try 中：任何异常都只记录日志、绝不向外抛出，
            // 否则 VSTO 会把加载项标记为加载失败，Word 随后会静默禁用它。
            try
            {
                var stopwatch = Stopwatch.StartNew();

                _isShuttingDown = false;
                _startupDiagnosticsWritten = false;

                AddInLog.Write("==== 加载项启动 ====");
                AddInLog.Write("日志文件: {0}", AddInLog.LogFilePath);

                // 启动阶段只做最小必要工作：注册两个事件。
                // WPF 环境、贴图窗口预热、环境诊断全部推迟到首次使用 / 空闲时，避免拖慢 Word 启动。
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;

                ScheduleStartupDiagnostics();

                stopwatch.Stop();
                AddInLog.Write(
                    "加载项启动完成，Startup 耗时 {0:F1} ms。",
                    stopwatch.Elapsed.TotalMilliseconds);
            }
            catch (Exception ex)
            {
                AddInLog.Error("加载项启动失败", ex);
            }
        }

        /// <summary>
        /// 把环境诊断推迟到 Word 打开完成、界面空闲后再写，避免在启动关键路径上
        /// 反复跨 COM 读取 Version / Build 并同步写文件。
        /// </summary>
        private void ScheduleStartupDiagnostics()
        {
            try
            {
                var timer = new System.Windows.Forms.Timer
                {
                    Interval = StartupDiagnosticsDelayMilliseconds
                };

                timer.Tick += delegate
                {
                    timer.Stop();
                    timer.Dispose();

                    if (_isShuttingDown || _startupDiagnosticsWritten)
                    {
                        return;
                    }

                    _startupDiagnosticsWritten = true;
                    WriteStartupDiagnostics();
                };

                timer.Start();
            }
            catch (Exception ex)
            {
                AddInLog.Error("安排启动诊断写入失败", ex);
            }
        }

        /// <summary>把运行环境信息写入日志，便于在目标机上定位加载失败原因。</summary>
        private void WriteStartupDiagnostics()
        {
            try
            {
                AddInLog.Write("操作系统: {0}", Environment.OSVersion);
                AddInLog.Write(".NET 版本: {0}", Environment.Version);
                AddInLog.Write("进程位数: {0}", Environment.Is64BitProcess ? "64 位" : "32 位");
                AddInLog.Write("程序集版本: {0}", typeof(ThisAddIn).Assembly.GetName().Version);
            }
            catch (Exception ex)
            {
                AddInLog.Error("写入环境信息失败", ex);
            }

            try
            {
                AddInLog.Write("Word 版本: {0}", Application.Version);
                AddInLog.Write("Word Build: {0}", Application.Build);
                AddInLog.Write("Word 位数: {0}", Environment.Is64BitProcess ? "64 位" : "32 位");
            }
            catch (Exception ex)
            {
                AddInLog.Error("读取 Word 版本失败", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _isShuttingDown = true;

            AddInLog.Write("加载项开始卸载。");

            try
            {
                Application.WindowSelectionChange -= Application_WindowSelectionChange;
                Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
            }
            catch (Exception ex)
            {
                AddInLog.Error("注销事件失败", ex);
            }

            CancelPendingSelectionCapture();

            // 只有用过贴图功能才存在管理器，未使用时这里完全不会触达 WPF。
            if (_pinManager != null)
            {
                _pinManager.CloseAll();
            }

            ClearCachedSelectionImage();
        }

        private void Application_WindowSelectionChange(Selection selection)
        {
            try
            {
                string selectionKey;
                if (!TryGetSelectionKey(selection, out selectionKey))
                {
                    CancelPendingSelectionCapture();
                    ClearCachedSelectionImage();
                    return;
                }

                if (selectionKey == _cachedSelectionKey && _cachedSelectionImage != null)
                {
                    CancelPendingSelectionCapture();
                    return;
                }

                // 选中图片时不再同步截图：Word 序列化 WordOpenXML 的开销很大，
                // 改为延迟防抖后预取，避免每次单击选中都阻塞 Word 界面线程。
                ScheduleSelectionCapture(selectionKey);

                // 用户已经选中图片，说明很可能会双击贴图：此时才启动 WPF 与贴图窗口预热。
                // 这一步从 Word 启动阶段移到了这里，启动速度不再受影响。
                RequestPinManagerPrewarmOnce();
            }
            catch (Exception ex)
            {
                AddInLog.Error("Application_WindowSelectionChange 处理失败", ex);
            }
        }

        private void ScheduleSelectionCapture(string selectionKey)
        {
            _pendingSelectionKey = selectionKey;

            try
            {
                if (_selectionCaptureTimer == null)
                {
                    _selectionCaptureTimer = new System.Windows.Forms.Timer
                    {
                        Interval = SelectionCaptureDelayMilliseconds
                    };
                    _selectionCaptureTimer.Tick += SelectionCaptureTimer_Tick;
                }

                _selectionCaptureTimer.Stop();
                _selectionCaptureTimer.Start();
            }
            catch (Exception ex)
            {
                AddInLog.Error("安排选区抓图失败", ex);
            }
        }

        private void CancelPendingSelectionCapture()
        {
            _pendingSelectionKey = null;

            if (_selectionCaptureTimer != null)
            {
                _selectionCaptureTimer.Stop();
            }
        }

        private void SelectionCaptureTimer_Tick(object sender, EventArgs e)
        {
            var selectionKey = _pendingSelectionKey;
            CancelPendingSelectionCapture();

            if (_isShuttingDown)
            {
                return;
            }

            if (string.IsNullOrEmpty(selectionKey))
            {
                return;
            }

            if (selectionKey == _cachedSelectionKey && _cachedSelectionImage != null)
            {
                return;
            }

            try
            {
                var selection = Application.Selection;

                string currentKey;
                if (!TryGetSelectionKey(selection, out currentKey) || currentKey != selectionKey)
                {
                    return;
                }

                var image = CaptureSelectionImage(selection);
                if (image == null)
                {
                    return;
                }

                ReplaceCachedSelectionImage(selectionKey, image);
            }
            catch (Exception ex)
            {
                AddInLog.Error("SelectionCaptureTimer_Tick 处理失败", ex);
            }
        }

        private void Application_WindowBeforeDoubleClick(Selection selection, ref bool cancel)
        {
            try
            {
                string selectionKey;
                if (!TryGetSelectionKey(selection, out selectionKey))
                {
                    return;
                }

                // 即将贴图：取消可能仍在排队的预取，避免重复截图。
                CancelPendingSelectionCapture();

                var image = TryCloneCachedSelectionImage(selectionKey);
                if (image == null)
                {
                    image = CaptureSelectionImage(selection);
                    if (image == null)
                    {
                        return;
                    }

                    // 先存入缓存，之后同一选区再次贴图可以直接克隆，无需重新截图。
                    ReplaceCachedSelectionImage(selectionKey, image);
                    image = TryCloneCachedSelectionImage(selectionKey);
                    if (image == null)
                    {
                        return;
                    }
                }

                ShowPin(image);
            }
            catch (Exception ex)
            {
                AddInLog.Error("Application_WindowBeforeDoubleClick 处理失败", ex);
            }
        }

        /// <summary>
        /// 判断选区里是否有图片。
        ///
        /// 先用 Selection.Type 走快速路径：打开文档、点击正文、敲键盘都会频繁触发
        /// WindowSelectionChange，而 Selection.InlineShapes / ShapeRange 都是跨进程
        /// COM 调用，每次都要往返 Word。光标状态（wdSelectionIP）和空选区
        /// （wdNoSelection）下必然没有图片，直接返回，避免这些无谓的往返。
        /// 只有拿到类型信息失败时才退回完整检查，确保行为不变。
        /// </summary>
        private static bool HasSelectedImage(Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                var selectionType = selection.Type;
                if (selectionType == WdSelectionType.wdNoSelection ||
                    selectionType == WdSelectionType.wdSelectionIP)
                {
                    return false;
                }
            }
            catch
            {
                // 读取 Type 失败时继续走完整检查，绝不因此漏掉图片。
            }

            return selection.InlineShapes.Count > 0 ||
                   (selection.ShapeRange != null && selection.ShapeRange.Count > 0);
        }

        private static bool TryGetSelectionKey(Selection selection, out string selectionKey)
        {
            selectionKey = null;

            if (!HasSelectedImage(selection))
            {
                return false;
            }

            if (selection.InlineShapes.Count > 0)
            {
                var inlineShape = selection.InlineShapes[1];
                selectionKey = string.Format(
                    "inline:{0}:{1}:{2}:{3}",
                    selection.Range.Start,
                    selection.Range.End,
                    inlineShape.Width,
                    inlineShape.Height);
                return true;
            }

            if (selection.ShapeRange != null && selection.ShapeRange.Count > 0)
            {
                var shape = selection.ShapeRange[1];
                selectionKey = string.Format(
                    "shape:{0}:{1}:{2}:{3}",
                    shape.Left,
                    shape.Top,
                    shape.Width,
                    shape.Height);
                return true;
            }

            return false;
        }

        private const string PackageNamespaceUri = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private const string DrawingNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string WordprocessingDrawingNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private const string RelationshipsNamespaceUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string DocumentPartName = "/word/document.xml";
        private const string DocumentRelationshipsPartName = "/word/_rels/document.xml.rels";

        private static Image CaptureSelectionImage(Selection selection)
        {
            // 优先从 OpenXML 中读取原始图片数据，能完整保留 PNG 等格式的透明通道。
            var originalImage = TryCaptureOriginalImageViaWordOpenXml(selection);
            if (originalImage != null)
            {
                return originalImage;
            }

            // 回退到 EMF 截取：Word 会把图片绘制到白底上导致透明信息丢失，
            // 同时 EMF 的边界往往比图片大，需要裁掉外围的透明边。
            return TryCaptureSelectionImageViaEnhMetaFileBits(selection);
        }

        /// <summary>
        /// 从选区的 WordOpenXML（扁平 OPC）中还原图片原始数据。
        /// </summary>
        private static Image TryCaptureOriginalImageViaWordOpenXml(Selection selection)
        {
            try
            {
                if (selection == null)
                {
                    return null;
                }

                var flatOpcXml = (string)null;
                var expectedShapeName = (string)null;

                if (selection.InlineShapes.Count > 0)
                {
                    flatOpcXml = selection.InlineShapes[1].Range.WordOpenXML as string;
                }
                else if (selection.ShapeRange != null && selection.ShapeRange.Count > 0)
                {
                    var shape = selection.ShapeRange[1];
                    expectedShapeName = shape.Name;

                    // 浮动图片优先取选区范围（只含锚定段落，通常比整段 XML 小）。
                    flatOpcXml = selection.Range.WordOpenXML as string;

                    if (string.IsNullOrWhiteSpace(flatOpcXml))
                    {
                        var anchorRange = shape.Anchor;
                        if (anchorRange != null)
                        {
                            flatOpcXml = anchorRange.WordOpenXML as string;
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(flatOpcXml))
                {
                    return null;
                }

                return ParseOriginalImageFromFlatOpc(flatOpcXml, expectedShapeName);
            }
            catch (Exception ex)
            {
                AddInLog.Error("TryCaptureOriginalImageViaWordOpenXml 失败", ex);
                return null;
            }
        }

        private static Image ParseOriginalImageFromFlatOpc(string flatOpcXml, string expectedShapeName)
        {
            string embedId;
            double cropLeft, cropTop, cropRight, cropBottom;

            if (!TryReadPictureBlipFromFlatOpc(
                    flatOpcXml,
                    expectedShapeName,
                    out embedId,
                    out cropLeft,
                    out cropTop,
                    out cropRight,
                    out cropBottom))
            {
                return null;
            }

            var imagePartName = ResolveRelationshipTarget(flatOpcXml, embedId);
            if (string.IsNullOrWhiteSpace(imagePartName) || !IsSupportedBitmapPart(imagePartName))
            {
                return null;
            }

            var imageBytes = ReadBinaryPartData(flatOpcXml, imagePartName);
            if (imageBytes == null || imageBytes.Length == 0)
            {
                return null;
            }

            var bitmap = DecodeImagePreservingAlpha(imageBytes);
            if (bitmap == null)
            {
                return null;
            }

            return CropBitmapBySourceRectangle(bitmap, cropLeft, cropTop, cropRight, cropBottom);
        }

        private static XmlReader CreateFlatOpcReader(string flatOpcXml)
        {
            var settings = new XmlReaderSettings
            {
                CloseInput = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            };

            return XmlReader.Create(new StringReader(flatOpcXml), settings);
        }

        private static string ReadPartName(XmlReader reader)
        {
            var name = reader.GetAttribute("name", PackageNamespaceUri);
            if (string.IsNullOrEmpty(name))
            {
                name = reader.GetAttribute("name");
            }

            return name;
        }

        /// <summary>
        /// 流式扫描扁平 OPC，定位图片的 a:blip 引用及裁剪设置（a:srcRect）。
        /// 只深入 /word/document.xml，其余 part 整段跳过，不构建 DOM。
        /// </summary>
        private static bool TryReadPictureBlipFromFlatOpc(
            string flatOpcXml,
            string expectedShapeName,
            out string embedId,
            out double left,
            out double top,
            out double right,
            out double bottom)
        {
            embedId = null;
            left = 0d;
            top = 0d;
            right = 0d;
            bottom = 0d;

            var fallbackEmbedId = (string)null;
            double fallbackLeft = 0d, fallbackTop = 0d, fallbackRight = 0d, fallbackBottom = 0d;
            var hasFallback = false;

            using (var reader = CreateFlatOpcReader(flatOpcXml))
            {
                var inDocumentPart = false;
                var anchorDepth = -1;
                var currentDocPrName = (string)null;
                var currentEmbedId = (string)null;
                double currentLeft = 0d, currentTop = 0d, currentRight = 0d, currentBottom = 0d;
                var skipRead = false;

                while (skipRead || reader.Read())
                {
                    skipRead = false;

                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.LocalName == "part" &&
                        reader.NamespaceURI == PackageNamespaceUri)
                    {
                        inDocumentPart = string.Equals(
                            ReadPartName(reader),
                            DocumentPartName,
                            StringComparison.OrdinalIgnoreCase);

                        if (!inDocumentPart && !reader.IsEmptyElement)
                        {
                            // 与图片无关的 part（含大块 Base64 媒体数据）整段跳过。
                            reader.Skip();
                            skipRead = reader.NodeType != XmlNodeType.None;
                        }

                        continue;
                    }

                    if (!inDocumentPart)
                    {
                        continue;
                    }

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (anchorDepth < 0 &&
                            (reader.LocalName == "anchor" || reader.LocalName == "inline") &&
                            reader.NamespaceURI == WordprocessingDrawingNamespaceUri)
                        {
                            anchorDepth = reader.Depth;
                            currentDocPrName = null;
                            currentEmbedId = null;
                            currentLeft = 0d;
                            currentTop = 0d;
                            currentRight = 0d;
                            currentBottom = 0d;
                            continue;
                        }

                        if (anchorDepth < 0)
                        {
                            continue;
                        }

                        if (reader.LocalName == "docPr" &&
                            reader.NamespaceURI == WordprocessingDrawingNamespaceUri)
                        {
                            currentDocPrName = reader.GetAttribute("name");
                        }
                        else if (reader.LocalName == "blip" &&
                                 reader.NamespaceURI == DrawingNamespaceUri)
                        {
                            var candidate = reader.GetAttribute("embed", RelationshipsNamespaceUri);
                            if (!string.IsNullOrEmpty(candidate))
                            {
                                currentEmbedId = candidate;
                            }
                        }
                        else if (reader.LocalName == "srcRect" &&
                                 reader.NamespaceURI == DrawingNamespaceUri)
                        {
                            currentLeft = ReadPercentageAttribute(reader, "l");
                            currentTop = ReadPercentageAttribute(reader, "t");
                            currentRight = ReadPercentageAttribute(reader, "r");
                            currentBottom = ReadPercentageAttribute(reader, "b");
                        }

                        continue;
                    }

                    if (reader.NodeType == XmlNodeType.EndElement &&
                        anchorDepth >= 0 &&
                        reader.Depth == anchorDepth &&
                        (reader.LocalName == "anchor" || reader.LocalName == "inline"))
                    {
                        anchorDepth = -1;

                        if (string.IsNullOrEmpty(currentEmbedId))
                        {
                            continue;
                        }

                        if (!hasFallback)
                        {
                            fallbackEmbedId = currentEmbedId;
                            fallbackLeft = currentLeft;
                            fallbackTop = currentTop;
                            fallbackRight = currentRight;
                            fallbackBottom = currentBottom;
                            hasFallback = true;
                        }

                        if (string.IsNullOrEmpty(expectedShapeName) ||
                            string.Equals(currentDocPrName, expectedShapeName, StringComparison.Ordinal))
                        {
                            embedId = currentEmbedId;
                            left = currentLeft;
                            top = currentTop;
                            right = currentRight;
                            bottom = currentBottom;
                            return true;
                        }
                    }
                }
            }

            if (!hasFallback)
            {
                return false;
            }

            // 名称对不上时（同一段落锚定多张浮动图片），退回第一张图片。
            embedId = fallbackEmbedId;
            left = fallbackLeft;
            top = fallbackTop;
            right = fallbackRight;
            bottom = fallbackBottom;
            return true;
        }

        private static string ResolveRelationshipTarget(string flatOpcXml, string relationshipId)
        {
            if (string.IsNullOrWhiteSpace(flatOpcXml) || string.IsNullOrWhiteSpace(relationshipId))
            {
                return null;
            }

            using (var reader = CreateFlatOpcReader(flatOpcXml))
            {
                var inRelationshipsPart = false;
                var skipRead = false;

                while (skipRead || reader.Read())
                {
                    skipRead = false;

                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.LocalName == "part" &&
                        reader.NamespaceURI == PackageNamespaceUri)
                    {
                        inRelationshipsPart = string.Equals(
                            ReadPartName(reader),
                            DocumentRelationshipsPartName,
                            StringComparison.OrdinalIgnoreCase);

                        if (!inRelationshipsPart && !reader.IsEmptyElement)
                        {
                            reader.Skip();
                            skipRead = reader.NodeType != XmlNodeType.None;
                        }

                        continue;
                    }

                    if (!inRelationshipsPart ||
                        reader.NodeType != XmlNodeType.Element ||
                        !string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!string.Equals(reader.GetAttribute("Id"), relationshipId, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }

                    var target = reader.GetAttribute("Target");
                    if (string.IsNullOrWhiteSpace(target))
                    {
                        return null;
                    }

                    return NormalizePartName("/word", target);
                }
            }

            return null;
        }

        private static byte[] ReadBinaryPartData(string flatOpcXml, string partName)
        {
            if (string.IsNullOrWhiteSpace(flatOpcXml) || string.IsNullOrWhiteSpace(partName))
            {
                return null;
            }

            using (var reader = CreateFlatOpcReader(flatOpcXml))
            {
                var isTargetPart = false;
                var skipRead = false;

                while (skipRead || reader.Read())
                {
                    skipRead = false;

                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.LocalName == "part" &&
                        reader.NamespaceURI == PackageNamespaceUri)
                    {
                        isTargetPart = string.Equals(
                            ReadPartName(reader),
                            partName,
                            StringComparison.OrdinalIgnoreCase);

                        if (!isTargetPart && !reader.IsEmptyElement)
                        {
                            reader.Skip();
                            skipRead = reader.NodeType != XmlNodeType.None;
                        }

                        continue;
                    }

                    if (!isTargetPart ||
                        reader.NodeType != XmlNodeType.Element ||
                        !string.Equals(reader.LocalName, "binaryData", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var base64Text = reader.ReadElementContentAsString();
                    if (string.IsNullOrWhiteSpace(base64Text))
                    {
                        return null;
                    }

                    return Convert.FromBase64String(base64Text.Trim());
                }
            }

            return null;
        }

        private static string NormalizePartName(string basePartName, string target)
        {
            target = target.Replace('\\', '/').Trim();

            var combined = target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : basePartName.TrimEnd('/') + "/" + target;

            var segments = combined.Split('/');
            var stack = new List<string>();

            foreach (var segment in segments)
            {
                if (segment.Length == 0 || segment == ".")
                {
                    continue;
                }

                if (segment == "..")
                {
                    if (stack.Count > 0)
                    {
                        stack.RemoveAt(stack.Count - 1);
                    }

                    continue;
                }

                stack.Add(segment);
            }

            return "/" + string.Join("/", stack.ToArray());
        }

        private static bool IsSupportedBitmapPart(string partName)
        {
            if (string.IsNullOrWhiteSpace(partName))
            {
                return false;
            }

            switch (Path.GetExtension(partName).ToLowerInvariant())
            {
                case ".png":
                case ".jpg":
                case ".jpeg":
                case ".gif":
                case ".bmp":
                case ".tif":
                case ".tiff":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 读取 a:srcRect 的百分比属性（单位 1/100000），换算为 0~1 的比例。
        /// </summary>
        private static double ReadPercentageAttribute(XmlReader reader, string attributeName)
        {
            double value;
            if (reader == null ||
                !double.TryParse(
                    reader.GetAttribute(attributeName),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value))
            {
                return 0d;
            }

            return Math.Max(0d, Math.Min(value / 100000d, 1d));
        }

        private static Bitmap DecodeImagePreservingAlpha(byte[] imageBytes)
        {
            try
            {
                using (var stream = new MemoryStream(imageBytes))
                using (var source = Image.FromStream(stream, true, true))
                {
                    return ConvertToArgb(source);
                }
            }
            catch
            {
                return null;
            }
        }

        private static Bitmap ConvertToArgb(Image source)
        {
            if (source == null || source.Width <= 0 || source.Height <= 0)
            {
                return null;
            }

            var bitmap = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            bitmap.SetResolution(96f, 96f);

            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphics.PixelOffsetMode = PixelOffsetMode.Half;
                graphics.DrawImage(
                    source,
                    new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    0,
                    0,
                    source.Width,
                    source.Height,
                    GraphicsUnit.Pixel);
            }

            return bitmap;
        }

        private static Image CropBitmapBySourceRectangle(
            Bitmap bitmap,
            double left,
            double top,
            double right,
            double bottom)
        {
            if (bitmap == null)
            {
                return null;
            }

            if (left <= 0d && top <= 0d && right <= 0d && bottom <= 0d)
            {
                return bitmap;
            }

            var width = bitmap.Width;
            var height = bitmap.Height;

            var leftPixels = (int)Math.Round(width * left);
            var topPixels = (int)Math.Round(height * top);
            var cropWidth = width - leftPixels - (int)Math.Round(width * right);
            var cropHeight = height - topPixels - (int)Math.Round(height * bottom);

            leftPixels = Math.Max(0, Math.Min(leftPixels, width - 1));
            topPixels = Math.Max(0, Math.Min(topPixels, height - 1));
            cropWidth = Math.Max(1, Math.Min(cropWidth, width - leftPixels));
            cropHeight = Math.Max(1, Math.Min(cropHeight, height - topPixels));

            if (cropWidth >= width && cropHeight >= height)
            {
                return bitmap;
            }

            var cropped = new Bitmap(cropWidth, cropHeight, PixelFormat.Format32bppArgb);
            cropped.SetResolution(96f, 96f);

            using (var graphics = Graphics.FromImage(cropped))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphics.PixelOffsetMode = PixelOffsetMode.Half;
                graphics.DrawImage(
                    bitmap,
                    new System.Drawing.Rectangle(0, 0, cropWidth, cropHeight),
                    leftPixels,
                    topPixels,
                    cropWidth,
                    cropHeight,
                    GraphicsUnit.Pixel);
            }

            bitmap.Dispose();
            return cropped;
        }

        /// <summary>
        /// 裁掉图片外围的全透明边，避免 EMF 边界比图片大时出现多余区域。
        /// </summary>
        private static Bitmap TrimTransparentMargins(Bitmap sourceBitmap)
        {
            if (sourceBitmap == null)
            {
                return null;
            }

            Bitmap working = null;
            try
            {
                working = sourceBitmap.PixelFormat == PixelFormat.Format32bppArgb
                    ? sourceBitmap
                    : ConvertToArgb(sourceBitmap);

                if (working == null)
                {
                    return sourceBitmap;
                }

                var width = working.Width;
                var height = working.Height;
                var bounds = new System.Drawing.Rectangle(0, 0, width, height);

                var minX = width;
                var minY = height;
                var maxX = -1;
                var maxY = -1;

                BitmapData bitmapData = null;
                try
                {
                    bitmapData = working.LockBits(bounds, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

                    var stride = bitmapData.Stride;
                    var buffer = new byte[stride * height];
                    Marshal.Copy(bitmapData.Scan0, buffer, 0, buffer.Length);

                    for (var y = 0; y < height; y++)
                    {
                        var rowOffset = y * stride;
                        for (var x = 0; x < width; x++)
                        {
                            if (buffer[rowOffset + (x * 4) + 3] == 0)
                            {
                                continue;
                            }

                            if (x < minX)
                            {
                                minX = x;
                            }

                            if (x > maxX)
                            {
                                maxX = x;
                            }

                            if (y < minY)
                            {
                                minY = y;
                            }

                            if (y > maxY)
                            {
                                maxY = y;
                            }
                        }
                    }
                }
                finally
                {
                    if (bitmapData != null)
                    {
                        working.UnlockBits(bitmapData);
                    }
                }

                if (maxX < 0 || maxY < 0)
                {
                    return sourceBitmap;
                }

                if (minX == 0 && minY == 0 && maxX == width - 1 && maxY == height - 1)
                {
                    return sourceBitmap;
                }

                var cropRectangle = new System.Drawing.Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
                var trimmed = new Bitmap(cropRectangle.Width, cropRectangle.Height, PixelFormat.Format32bppArgb);
                trimmed.SetResolution(96f, 96f);

                using (var graphics = Graphics.FromImage(trimmed))
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                    graphics.PixelOffsetMode = PixelOffsetMode.Half;
                    graphics.DrawImage(
                        working,
                        new System.Drawing.Rectangle(0, 0, cropRectangle.Width, cropRectangle.Height),
                        cropRectangle,
                        GraphicsUnit.Pixel);
                }

                return trimmed;
            }
            catch
            {
                return sourceBitmap;
            }
            finally
            {
                if (working != null && !ReferenceEquals(working, sourceBitmap))
                {
                    working.Dispose();
                }
            }
        }

        private static Image TryCaptureSelectionImageViaEnhMetaFileBits(Selection selection)
        {
            try
            {
                if (selection == null)
                {
                    return null;
                }

                var bits = selection.EnhMetaFileBits as byte[];
                if (bits == null || bits.Length == 0)
                {
                    return null;
                }

                using (var stream = new MemoryStream(bits))
                using (var bitmap = new Bitmap(stream))
                {
                    var trimmed = TrimTransparentMargins(bitmap) ?? bitmap;

                    Image result = null;
                    if (selection.Type == WdSelectionType.wdSelectionInlineShape &&
                        selection.InlineShapes.Count > 0)
                    {
                        var inlineShape = selection.InlineShapes[1];
                        result = CropInlineShapeBitmap(trimmed, inlineShape);
                    }

                    if (result == null)
                    {
                        result = new Bitmap(trimmed);
                    }

                    if (!ReferenceEquals(trimmed, bitmap))
                    {
                        trimmed.Dispose();
                    }

                    return result;
                }
            }
            catch
            {
                return null;
            }
        }

        private static Image CropInlineShapeBitmap(Bitmap sourceBitmap, InlineShape inlineShape)
        {
            try
            {
                if (sourceBitmap == null || inlineShape == null)
                {
                    return null;
                }

                if (inlineShape.Height <= 0 || inlineShape.Width <= 0)
                {
                    return new Bitmap(sourceBitmap);
                }

                var targetWidth = (int)Math.Round(sourceBitmap.Height * (inlineShape.Width / inlineShape.Height));
                targetWidth = Math.Max(1, Math.Min(targetWidth, sourceBitmap.Width));

                var cropRectangle = new System.Drawing.Rectangle(0, 0, targetWidth, sourceBitmap.Height);
                var croppedBitmap = new Bitmap(cropRectangle.Width, cropRectangle.Height, PixelFormat.Format32bppArgb);
                croppedBitmap.SetResolution(96f, 96f);

                using (var graphics = Graphics.FromImage(croppedBitmap))
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                    graphics.PixelOffsetMode = PixelOffsetMode.Half;
                    graphics.DrawImage(
                        sourceBitmap,
                        new System.Drawing.Rectangle(0, 0, croppedBitmap.Width, croppedBitmap.Height),
                        cropRectangle,
                        System.Drawing.GraphicsUnit.Pixel);
                }

                return croppedBitmap;
            }
            catch (Exception ex)
            {
                AddInLog.Error("CropBitmapBySourceRectangle 失败，回退为整图", ex);
                return new Bitmap(sourceBitmap);
            }
        }

        /// <summary>
        /// 把图片贴到屏幕上。位图所有权移交给新窗口，窗口关闭时自行释放。
        /// 窗口与 WPF 相关的管理工作全部交给 <see cref="PinManager"/>。
        /// </summary>
        private void ShowPin(Image image)
        {
            GetOrCreatePinManager().ShowPin(image);
        }

        /// <summary>
        /// 取贴图窗口管理器，不存在时创建。
        /// 这是加载项中唯一会触达 WPF 类型（ImagePreviewWindow）的位置，
        /// 因此 WPF 程序集的加载被推迟到用户第一次选中图片时。
        /// </summary>
        private PinManager GetOrCreatePinManager()
        {
            if (_pinManager == null)
            {
                _pinManager = new PinManager();
            }

            return _pinManager;
        }

        /// <summary>
        /// 命中缓存时克隆一份位图交给贴图窗口，缓存继续保留以便重复贴同一张图。
        /// </summary>
        private Image TryCloneCachedSelectionImage(string selectionKey)
        {
            if (_cachedSelectionImage == null || selectionKey != _cachedSelectionKey)
            {
                return null;
            }

            try
            {
                var source = _cachedSelectionImage as Bitmap;
                var copy = source != null
                    ? source.Clone(new System.Drawing.Rectangle(0, 0, source.Width, source.Height), PixelFormat.Format32bppArgb)
                    : new Bitmap(_cachedSelectionImage);

                copy.SetResolution(96f, 96f);
                return copy;
            }
            catch (Exception ex)
            {
                AddInLog.Error("TryCloneCachedSelectionImage 失败", ex);
                return null;
            }
        }

        private void ReplaceCachedSelectionImage(string selectionKey, Image image)
        {
            ClearCachedSelectionImage();
            _cachedSelectionKey = selectionKey;
            _cachedSelectionImage = image;
        }

        private void ClearCachedSelectionImage()
        {
            if (_cachedSelectionImage != null)
            {
                _cachedSelectionImage.Dispose();
                _cachedSelectionImage = null;
            }

            _cachedSelectionKey = null;
        }

        /// <summary>
        /// 用户第一次选中图片后，在界面空闲时预热贴图窗口并初始化 WPF 环境。
        /// 不在 Word 启动阶段做这件事，避免拖慢启动。
        ///
        /// 用独立定时器延后触发：WindowSelectionChange 在用户点选图片时同步执行，
        /// 此刻创建 WPF 环境会让选中操作本身出现卡顿，推迟到消息循环空闲再加载，
        /// 用户感知不到，而双击贴图时窗口通常已经就绪。
        /// </summary>
        private void RequestPinManagerPrewarmOnce()
        {
            if (_isShuttingDown || _pinPrewarmRequested)
            {
                return;
            }

            _pinPrewarmRequested = true;

            try
            {
                var timer = new System.Windows.Forms.Timer
                {
                    Interval = PinPrewarmDelayMilliseconds
                };

                timer.Tick += delegate
                {
                    timer.Stop();
                    timer.Dispose();

                    if (_isShuttingDown)
                    {
                        return;
                    }

                    SchedulePreviewWindowPrewarm();
                };

                timer.Start();
            }
            catch (Exception ex)
            {
                AddInLog.Error("安排贴图窗口预热失败", ex);
            }
        }

        /// <summary>把预热请求转交给 <see cref="PinManager"/>，由它在 UI 空闲时创建窗口。</summary>
        private void SchedulePreviewWindowPrewarm()
        {
            if (_isShuttingDown)
            {
                return;
            }

            GetOrCreatePinManager().SchedulePreviewWindowPrewarm();
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
