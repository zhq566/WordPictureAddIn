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
        private ImagePreviewWindow _previewWindow;
        private string _cachedSelectionKey;
        private Image _cachedSelectionImage;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            EnsureWpfApplication();
            EnsurePreviewWindow();
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WindowSelectionChange -= Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;

            if (_previewWindow != null)
            {
                _previewWindow.Close();
                _previewWindow = null;
            }

            ClearCachedSelectionImage();
        }

        private void Application_WindowSelectionChange(Selection selection)
        {
            string selectionKey;
            if (!TryGetSelectionKey(selection, out selectionKey))
            {
                ClearCachedSelectionImage();
                return;
            }

            if (selectionKey == _cachedSelectionKey && _cachedSelectionImage != null)
            {
                return;
            }

            var image = CaptureSelectionImage(selection);
            if (image == null)
            {
                ClearCachedSelectionImage();
                return;
            }

            ReplaceCachedSelectionImage(selectionKey, image);
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

                var image = TryGetCachedSelectionImage(selectionKey) ?? CaptureSelectionImage(selection);
                if (image == null)
                {
                    return;
                }

                ShowPreview(image);
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

                if (string.IsNullOrWhiteSpace(flatOpcXml) &&
                    selection.ShapeRange != null &&
                    selection.ShapeRange.Count > 0)
                {
                    var shape = selection.ShapeRange[1];
                    expectedShapeName = shape.Name;

                    var anchorRange = shape.Anchor;
                    if (anchorRange != null)
                    {
                        flatOpcXml = anchorRange.WordOpenXML as string;
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
            var packageDocument = new XmlDocument();
            packageDocument.PreserveWhitespace = false;
            packageDocument.XmlResolver = null;
            packageDocument.LoadXml(flatOpcXml);

            var packageNamespaceManager = CreatePackageNamespaceManager(packageDocument.NameTable);

            var documentPart = FindPackagePart(packageDocument, packageNamespaceManager, DocumentPartName);
            var documentRoot = GetPartXmlRoot(documentPart);
            if (documentRoot == null)
            {
                return null;
            }

            var drawingNamespaceManager = CreateDrawingNamespaceManager(documentRoot.OwnerDocument.NameTable);

            var blip = FindPictureBlipElement(documentRoot, drawingNamespaceManager, expectedShapeName);
            if (blip == null)
            {
                return null;
            }

            var embedId = blip.GetAttribute("embed", RelationshipsNamespaceUri);
            if (string.IsNullOrWhiteSpace(embedId))
            {
                return null;
            }

            var imagePartName = ResolveRelationshipTarget(packageDocument, packageNamespaceManager, embedId);
            if (string.IsNullOrWhiteSpace(imagePartName) || !IsSupportedBitmapPart(imagePartName))
            {
                return null;
            }

            var imageBytes = ReadBinaryPartData(
                FindPackagePart(packageDocument, packageNamespaceManager, imagePartName));

            if (imageBytes == null || imageBytes.Length == 0)
            {
                return null;
            }

            var bitmap = DecodeImagePreservingAlpha(imageBytes);
            if (bitmap == null)
            {
                return null;
            }

            double cropLeft, cropTop, cropRight, cropBottom;
            ReadSourceRectangle(blip, out cropLeft, out cropTop, out cropRight, out cropBottom);

            return CropBitmapBySourceRectangle(bitmap, cropLeft, cropTop, cropRight, cropBottom);
        }

        private static XmlNamespaceManager CreatePackageNamespaceManager(XmlNameTable nameTable)
        {
            var namespaceManager = new XmlNamespaceManager(nameTable);
            namespaceManager.AddNamespace("pkg", PackageNamespaceUri);
            return namespaceManager;
        }

        private static XmlNamespaceManager CreateDrawingNamespaceManager(XmlNameTable nameTable)
        {
            var namespaceManager = new XmlNamespaceManager(nameTable);
            namespaceManager.AddNamespace("a", DrawingNamespaceUri);
            namespaceManager.AddNamespace("r", RelationshipsNamespaceUri);
            namespaceManager.AddNamespace("wp", WordprocessingDrawingNamespaceUri);
            namespaceManager.AddNamespace("pkg", PackageNamespaceUri);
            return namespaceManager;
        }

        private static XmlElement FindPackagePart(
            XmlDocument packageDocument,
            XmlNamespaceManager packageNamespaceManager,
            string partName)
        {
            if (packageDocument == null || string.IsNullOrWhiteSpace(partName))
            {
                return null;
            }

            var parts = packageDocument.SelectNodes("/pkg:package/pkg:part", packageNamespaceManager);
            if (parts == null)
            {
                return null;
            }

            foreach (XmlNode node in parts)
            {
                var part = node as XmlElement;
                if (part == null)
                {
                    continue;
                }

                var name = part.GetAttribute("name", PackageNamespaceUri);
                if (string.IsNullOrEmpty(name))
                {
                    name = part.GetAttribute("name");
                }

                if (string.Equals(name, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return part;
                }
            }

            return null;
        }

        private static XmlElement GetPartXmlRoot(XmlElement partElement)
        {
            if (partElement == null)
            {
                return null;
            }

            foreach (XmlNode child in partElement.ChildNodes)
            {
                var xmlData = child as XmlElement;
                if (xmlData == null || !string.Equals(xmlData.LocalName, "xmlData", StringComparison.Ordinal))
                {
                    continue;
                }

                foreach (XmlNode dataChild in xmlData.ChildNodes)
                {
                    var dataElement = dataChild as XmlElement;
                    if (dataElement != null)
                    {
                        return dataElement;
                    }
                }
            }

            return null;
        }

        private static byte[] ReadBinaryPartData(XmlElement partElement)
        {
            if (partElement == null)
            {
                return null;
            }

            foreach (XmlNode child in partElement.ChildNodes)
            {
                var binaryData = child as XmlElement;
                if (binaryData == null ||
                    !string.Equals(binaryData.LocalName, "binaryData", StringComparison.Ordinal))
                {
                    continue;
                }

                var base64Text = binaryData.InnerText;
                if (string.IsNullOrWhiteSpace(base64Text))
                {
                    return null;
                }

                return Convert.FromBase64String(base64Text.Trim());
            }

            return null;
        }

        private static XmlElement FindPictureBlipElement(
            XmlElement documentRoot,
            XmlNamespaceManager namespaceManager,
            string expectedShapeName)
        {
            if (documentRoot == null)
            {
                return null;
            }

            var anchors = documentRoot.SelectNodes(".//wp:anchor | .//wp:inline", namespaceManager);
            if (anchors != null)
            {
                foreach (XmlNode node in anchors)
                {
                    var anchor = node as XmlElement;
                    if (anchor == null)
                    {
                        continue;
                    }

                    var blip = FindBlipInside(anchor, namespaceManager);
                    if (blip == null)
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(expectedShapeName))
                    {
                        return blip;
                    }

                    var docPr = anchor.SelectSingleNode(".//wp:docPr", namespaceManager) as XmlElement;
                    if (docPr == null)
                    {
                        continue;
                    }

                    if (string.Equals(docPr.GetAttribute("name"), expectedShapeName, StringComparison.Ordinal))
                    {
                        return blip;
                    }
                }
            }

            // 多个浮动图片锚定在同一段落时，名称可能对不上，退回到文档中第一个图片引用。
            return documentRoot.SelectSingleNode(".//a:blip[@r:embed]", namespaceManager) as XmlElement;
        }

        private static XmlElement FindBlipInside(XmlElement container, XmlNamespaceManager namespaceManager)
        {
            if (container == null)
            {
                return null;
            }

            return container.SelectSingleNode(".//a:blip[@r:embed]", namespaceManager) as XmlElement;
        }

        private static string ResolveRelationshipTarget(
            XmlDocument packageDocument,
            XmlNamespaceManager packageNamespaceManager,
            string relationshipId)
        {
            var relationshipsRoot = GetPartXmlRoot(
                FindPackagePart(packageDocument, packageNamespaceManager, DocumentRelationshipsPartName));

            if (relationshipsRoot == null)
            {
                return null;
            }

            foreach (XmlNode node in relationshipsRoot.ChildNodes)
            {
                var relationship = node as XmlElement;
                if (relationship == null ||
                    !string.Equals(relationship.LocalName, "Relationship", StringComparison.Ordinal))
                {
                    continue;
                }

                if (!string.Equals(relationship.GetAttribute("Id"), relationshipId, StringComparison.Ordinal))
                {
                    continue;
                }

                if (string.Equals(relationship.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                var target = relationship.GetAttribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    return null;
                }

                return NormalizePartName("/word", target);
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
        /// 读取图片裁剪设置（a:srcRect），返回值已换算为 0~1 的比例。
        /// </summary>
        private static void ReadSourceRectangle(
            XmlElement blip,
            out double left,
            out double top,
            out double right,
            out double bottom)
        {
            left = 0d;
            top = 0d;
            right = 0d;
            bottom = 0d;

            var blipFill = blip == null ? null : blip.ParentNode as XmlElement;
            if (blipFill == null)
            {
                return;
            }

            XmlElement sourceRectangle = null;
            foreach (XmlNode child in blipFill.ChildNodes)
            {
                var element = child as XmlElement;
                if (element != null && string.Equals(element.LocalName, "srcRect", StringComparison.Ordinal))
                {
                    sourceRectangle = element;
                    break;
                }
            }

            if (sourceRectangle == null)
            {
                return;
            }

            left = ReadPercentageAttribute(sourceRectangle, "l");
            top = ReadPercentageAttribute(sourceRectangle, "t");
            right = ReadPercentageAttribute(sourceRectangle, "r");
            bottom = ReadPercentageAttribute(sourceRectangle, "b");
        }

        private static double ReadPercentageAttribute(XmlElement element, string attributeName)
        {
            double value;
            if (!double.TryParse(
                    element.GetAttribute(attributeName),
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

        private void ShowPreview(Image image)
        {
            EnsurePreviewWindow();

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

        private Image TryGetCachedSelectionImage(string selectionKey)
        {
            if (_cachedSelectionImage == null || selectionKey != _cachedSelectionKey)
            {
                return null;
            }

            return new Bitmap(_cachedSelectionImage);
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

        private void EnsurePreviewWindow()
        {
            if (_previewWindow == null)
            {
                _previewWindow = new ImagePreviewWindow();
                _previewWindow.Closed += delegate
                {
                    _previewWindow = null;
                };
            }
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
