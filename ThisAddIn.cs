using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace WordPictureAddIn
{
    public partial class ThisAddIn
    {
        private const double SelectionCaptureDelayMilliseconds = 260.0d;

        private ImagePreviewWindow _previewWindow;
        private string _cachedSelectionKey;
        private Image _cachedSelectionImage;
        private System.Windows.Threading.DispatcherTimer _selectionCaptureTimer;
        private string _pendingSelectionKey;
        private string _previewSelectionKey;
        private bool _isShuttingDown;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _isShuttingDown = false;
            EnsureWpfApplication();
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
            SchedulePreviewWindowPrewarm();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _isShuttingDown = true;

            Application.WindowSelectionChange -= Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;

            CancelPendingSelectionCapture();

            if (_previewWindow != null)
            {
                _previewWindow.Close();
                _previewWindow = null;
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
            }
            catch
            {
            }
        }

        private void ScheduleSelectionCapture(string selectionKey)
        {
            _pendingSelectionKey = selectionKey;

            var application = System.Windows.Application.Current;
            if (application == null)
            {
                return;
            }

            if (_selectionCaptureTimer == null)
            {
                _selectionCaptureTimer = new System.Windows.Threading.DispatcherTimer(
                    System.Windows.Threading.DispatcherPriority.Background,
                    application.Dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(SelectionCaptureDelayMilliseconds)
                };
                _selectionCaptureTimer.Tick += SelectionCaptureTimer_Tick;
            }

            _selectionCaptureTimer.Stop();
            _selectionCaptureTimer.Start();
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
            catch
            {
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

                // 即将打开预览：取消可能仍在排队的预取，避免重复截图。
                CancelPendingSelectionCapture();

                var image = TryTakeCachedSelectionImage(selectionKey) ?? CaptureSelectionImage(selection);
                if (image == null)
                {
                    return;
                }

                ShowPreview(image, selectionKey);
            }
            catch
            {
            }
        }

        private static bool HasSelectedImage(Selection selection)
        {
            if (selection == null)
            {
                return false;
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
            catch
            {
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
            catch
            {
                return new Bitmap(sourceBitmap);
            }
        }

        private void ShowPreview(Image image, string selectionKey)
        {
            EnsurePreviewWindow();

            _previewSelectionKey = selectionKey;
            _previewWindow.SetImage(image);

            if (!_previewWindow.IsVisible)
            {
                _previewWindow.Show();
            }
            else
            {
                _previewWindow.Activate();
            }

            _previewWindow.Activate();
        }

        /// <summary>
        /// 命中缓存时直接把位图所有权交给预览窗口，避免再复制一份整图数据。
        /// 交出后缓存置空，防止后续 Dispose 到已被窗口使用的位图。
        /// </summary>
        private Image TryTakeCachedSelectionImage(string selectionKey)
        {
            if (_cachedSelectionImage == null || selectionKey != _cachedSelectionKey)
            {
                return null;
            }

            var image = _cachedSelectionImage;
            _cachedSelectionImage = null;
            _cachedSelectionKey = null;
            return image;
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

        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };
            }
        }

        /// <summary>
        /// Word 启动时同步创建 WPF 窗口会拖慢启动速度，改为界面空闲时预热。
        /// </summary>
        private void SchedulePreviewWindowPrewarm()
        {
            var application = System.Windows.Application.Current;
            if (application == null)
            {
                return;
            }

            application.Dispatcher.BeginInvoke(
            System.Windows.Threading.DispatcherPriority.ApplicationIdle,
            new Action(delegate
            {
                if (_isShuttingDown)
                {
                    return;
                }

                try
                {
                    EnsurePreviewWindow();
                }
                catch
                {
                }
            }));
        }

        private bool TryGetCurrentSelectionKey(out string selectionKey)
        {
            selectionKey = null;

            if (_isShuttingDown)
            {
                return false;
            }

            try
            {
                return TryGetSelectionKey(Application.Selection, out selectionKey);
            }
            catch
            {
                return false;
            }
        }

        private void EnsurePreviewWindow()
        {
            if (_previewWindow != null)
            {
                return;
            }

            var window = new ImagePreviewWindow();
            _previewWindow = window;
            window.Closed += delegate
            {
                // 预览窗口里的位图由 SetImage 移交所有权，关闭时取回。
                // 若当前选区仍是刚才那张图，直接放回缓存，再次双击无需重新截图。
                var bitmap = window.ReleaseBitmap();
                if (bitmap != null)
                {
                    string currentKey;
                    if (!_isShuttingDown &&
                        _previewSelectionKey != null &&
                        TryGetCurrentSelectionKey(out currentKey) &&
                        currentKey == _previewSelectionKey)
                    {
                        ReplaceCachedSelectionImage(currentKey, bitmap);
                    }
                    else
                    {
                        bitmap.Dispose();
                    }
                }

                _previewSelectionKey = null;

                if (ReferenceEquals(_previewWindow, window))
                {
                    _previewWindow = null;
                }

                if (_isShuttingDown)
                {
                    return;
                }

                // 关闭后再次空闲时预热新窗口，保证下一次双击依然能快速弹出。
                SchedulePreviewWindowPrewarm();
            };
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
