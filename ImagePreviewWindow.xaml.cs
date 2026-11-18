using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;
using WordPictureAddIn.Properties;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingImage = System.Drawing.Image;

namespace WordPictureAddIn
{
    /// <summary>
    /// 贴图窗口：按图片原始像素尺寸贴在屏幕上（超过工作区则等比缩小），
    /// 可拖动移动、滚轮缩放，失去焦点不关闭，悬停显示工具栏。
    /// </summary>
    public partial class ImagePreviewWindow : Window
    {
        private const double ZoomStepFactor = 1.1d;
        private const double MinZoom = 0.1d;
        private const double MaxZoom = 8.0d;
        private const string DefaultAnnotationColor = "#FF4500";
        private const double DefaultAnnotationStrokeWidth = 4.0d;
        private const double DefaultAnnotationFontSize = 20.0d;
        private const double MinAnnotationStrokeWidth = 1.0d;
        private const double MaxAnnotationStrokeWidth = 12.0d;
        private const double MinAnnotationFontSize = 12.0d;
        private const double MaxAnnotationFontSize = 72.0d;
        private const double TextBoxHorizontalPadding = 12.0d;
        private const double TextBoxVerticalPadding = 8.0d;
        private const string DefaultBackgroundColor = "#C7EDCC";
        private const string TransparentBackgroundColor = "Transparent";

        /// <summary>文字工具可选的字号档位。</summary>
        private static readonly double[] AnnotationFontSizeSteps =
        {
            12.0d, 14.0d, 16.0d, 18.0d, 20.0d, 24.0d, 28.0d, 32.0d, 40.0d, 48.0d, 64.0d, 72.0d
        };

        /// <summary>图片外围留给阴影的边距（DIP）。</summary>
        private const double ShadowMargin = 10.0d;

        /// <summary>贴图本体边框的描边宽度（DIP），与 XAML 中 RootBorder 保持一致。</summary>
        private const double PinBorderThickness = 1.0d;

        /// <summary>窗口左上角到图片左上角的偏移（阴影边距 + 描边）。</summary>
        private const double ContentOffset = ShadowMargin + PinBorderThickness;

        /// <summary>贴图最大占工作区的比例，超出时等比缩小。</summary>
        private const double MaxWorkspaceRatio = 0.9d;

        private const int ToolbarHideDelayMilliseconds = 320;

        /// <summary>显示用位图的最长边像素上限，超大图片会先降采样再交给 WPF。</summary>
        private const int MaxDisplayPixelDimension = 4096;

        private DrawingBitmap _bitmap;
        private bool _ownsBitmap;

        /// <summary>缩放比例 1 时的图片尺寸（DIP），按物理像素 / DPI 缩放比换算。</summary>
        private double _baseImageWidth;
        private double _baseImageHeight;
        private double _zoom = 1.0d;

        /// <summary>内容超出视口后的平移偏移（DIP，负值表示内容向左上移），始终限制在 [视口 - 内容, 0]。</summary>
        private double _panX;
        private double _panY;

        /// <summary>是否正在用左键拖动平移内容（而不是拖动窗口）。</summary>
        private bool _isPanning;
        private Point _panStartMouse;
        private double _panStartOffsetX;
        private double _panStartOffsetY;

        private bool _drawingAnnotation;
        private Point _annotationStart;
        private Shape _previewShape;
        private Path _previewArrow;
        private AnnotationTool _currentTool = AnnotationTool.None;
        private readonly List<UIElement> _annotationHistory = new List<UIElement>();
        private Grid _activeTextEditorContainer;
        private TextBox _activeTextEditor;
        private bool _isClosing;
        private int _interactionDepth;
        private bool _suppressToolbarRefresh;
        private DispatcherTimer _interactionEndTimer;
        private DispatcherTimer _toolbarHideTimer;

        /// <summary>当前标注颜色，从设置读取，修改后立即持久化。</summary>
        private Color _annotationColor;
        private double _annotationStrokeWidth = DefaultAnnotationStrokeWidth;
        private double _annotationFontSize = DefaultAnnotationFontSize;

        public ImagePreviewWindow()
        {
            InitializeComponent();

            _annotationColor = LoadSavedAnnotationColor();
            _annotationStrokeWidth = ClampStrokeWidth(LoadSavedAnnotationStrokeWidth());
            _annotationFontSize = NormalizeFontSize(LoadSavedAnnotationFontSize());

            UpdateCursorVisual();
            ApplyBackgroundColor(LoadSavedBackgroundColor());
            UpdateToolOptionsPanel();
        }

        /// <summary>
        /// 交给窗口的位图所有权归窗口所有，关闭时自动释放。
        /// </summary>
        public void SetImage(DrawingImage image)
        {
            ReleaseImage();

            _bitmap = EnsureArgbBitmap(image, out _ownsBitmap);

            // 传入的位图所有权已交给贴图窗口；若转换产生了新位图，原始数据可立即释放。
            if (image != null && !ReferenceEquals(image, _bitmap))
            {
                image.Dispose();
            }

            // P1：不再做 PNG 编码再解码的往返，直接从像素缓冲区生成 BitmapSource。
            // P2：超大图片按显示上限降采样，导出时仍使用 _bitmap 的全分辨率数据。
            PreviewImage.Source = CreateDisplaySource(_bitmap);

            ClearAnnotations();
            ApplyPinLayout();
        }

        /// <summary>
        /// 统一转换为 32bppArgb 位图，确保 PNG 等格式的透明通道不会在复制过程中丢失。
        /// </summary>
        private static DrawingBitmap EnsureArgbBitmap(DrawingImage image, out bool ownsBitmap)
        {
            ownsBitmap = true;

            if (image == null)
            {
                return null;
            }

            var source = image as DrawingBitmap;
            if (source != null &&
                source.PixelFormat == System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            {
                // 已经是目标格式：直接接管所有权，避免整图复制。
                return source;
            }

            var bitmap = new DrawingBitmap(
                image.Width,
                image.Height,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            bitmap.SetResolution(96f, 96f);

            using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                graphics.DrawImage(
                    image,
                    new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    0,
                    0,
                    image.Width,
                    image.Height,
                    System.Drawing.GraphicsUnit.Pixel);
            }

            return bitmap;
        }

        /// <summary>
        /// 显示用位图源：超过显示上限的大图先降采样，减少 WPF 渲染与内存压力。
        /// </summary>
        private static BitmapSource CreateDisplaySource(DrawingBitmap bitmap)
        {
            if (bitmap == null)
            {
                return null;
            }

            var sourceBitmap = bitmap;
            var ownsSourceBitmap = false;

            try
            {
                var maxSide = Math.Max(bitmap.Width, bitmap.Height);
                if (maxSide > MaxDisplayPixelDimension)
                {
                    var ratio = (double)MaxDisplayPixelDimension / maxSide;
                    var narrowWidth = Math.Max(1, (int)Math.Round(bitmap.Width * ratio));
                    var narrowHeight = Math.Max(1, (int)Math.Round(bitmap.Height * ratio));

                    var scaled = new DrawingBitmap(
                        narrowWidth,
                        narrowHeight,
                        System.Drawing.Imaging.PixelFormat.Format32bppArgb);

                    try
                    {
                        using (var graphics = System.Drawing.Graphics.FromImage(scaled))
                        {
                            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            graphics.DrawImage(
                                bitmap,
                                new System.Drawing.Rectangle(0, 0, narrowWidth, narrowHeight),
                                0,
                                0,
                                bitmap.Width,
                                bitmap.Height,
                                System.Drawing.GraphicsUnit.Pixel);
                        }
                    }
                    catch
                    {
                        scaled.Dispose();
                        throw;
                    }

                    sourceBitmap = scaled;
                    ownsSourceBitmap = true;
                }

                return CreateBitmapSource(sourceBitmap);
            }
            finally
            {
                if (ownsSourceBitmap)
                {
                    sourceBitmap.Dispose();
                }
            }
        }

        /// <summary>
        /// 直接用 LockBits 的像素缓冲区生成 Bgra32 的 BitmapSource，省掉 PNG 编解码。
        /// </summary>
        private static BitmapSource CreateBitmapSource(DrawingBitmap bitmap)
        {
            var bounds = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(
                bounds,
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            try
            {
                var source = BitmapSource.Create(
                    bitmap.Width,
                    bitmap.Height,
                    96.0d,
                    96.0d,
                    PixelFormats.Bgra32,
                    null,
                    data.Scan0,
                    data.Stride * bitmap.Height,
                    data.Stride);

                source.Freeze();
                return source;
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        #region 贴图布局与缩放

        /// <summary>
        /// 按原始像素尺寸贴图：1 图像像素对应 1 物理像素（已按 DPI 换算为 DIP），
        /// 超过工作区限制时等比缩小，并居中贴到鼠标所在显示器的工作区。
        /// </summary>
        private void ApplyPinLayout()
        {
            if (_bitmap == null)
            {
                return;
            }

            var dpi = GetDpiScale();
            _baseImageWidth = _bitmap.Width / dpi;
            _baseImageHeight = _bitmap.Height / dpi;

            if (_baseImageWidth <= 0d || _baseImageHeight <= 0d)
            {
                return;
            }

            var workspace = GetTargetWorkspace();

            _zoom = 1.0d;
            var fitScale = Math.Min(
                workspace.Width * MaxWorkspaceRatio / _baseImageWidth,
                workspace.Height * MaxWorkspaceRatio / _baseImageHeight);
            if (fitScale < 1.0d)
            {
                _zoom = Math.Max(MinZoom, fitScale);
            }

            _panX = 0d;
            _panY = 0d;

            ApplyVisualSize();
            ApplyWindowSize(workspace);

            Left = workspace.Left + ((workspace.Width - Width) / 2.0d);
            Top = workspace.Top + ((workspace.Height - Height) / 2.0d);

            UpdateCursorVisual();
        }

        /// <summary>内容层尺寸与缩放：始终以图片左上角为原点，平移偏移叠加在缩放之后。</summary>
        private void ApplyVisualSize()
        {
            PreviewImage.Width = _baseImageWidth;
            PreviewImage.Height = _baseImageHeight;
            ImageBackdrop.Width = _baseImageWidth;
            ImageBackdrop.Height = _baseImageHeight;
            AnnotationCanvas.Width = _baseImageWidth;
            AnnotationCanvas.Height = _baseImageHeight;
            ContentCanvas.Width = _baseImageWidth;
            ContentCanvas.Height = _baseImageHeight;
            ContentScaleTransform.ScaleX = _zoom;
            ContentScaleTransform.ScaleY = _zoom;
            ContentPanTransform.X = _panX;
            ContentPanTransform.Y = _panY;
        }

        /// <summary>
        /// 视口内可用区域的宽高（DIP）：窗口尺寸上限就是当前显示器工作区，
        /// 内容超出视口后窗口不再变大，靠平移继续查看图片。
        /// </summary>
        private static double GetMaxViewportWidth(Rect workspace)
        {
            return Math.Max(1.0d, workspace.Width - (ContentOffset * 2.0d));
        }

        private static double GetMaxViewportHeight(Rect workspace)
        {
            return Math.Max(1.0d, workspace.Height - (ContentOffset * 2.0d));
        }

        /// <summary>窗口尺寸跟随缩放变化，最大不超过所在显示器的工作区。</summary>
        private void ApplyWindowSize(Rect workspace)
        {
            var contentWidth = Math.Max(1.0d, _baseImageWidth * _zoom);
            var contentHeight = Math.Max(1.0d, _baseImageHeight * _zoom);

            _suppressToolbarRefresh = true;
            try
            {
                Width = Math.Min(contentWidth, GetMaxViewportWidth(workspace)) + (ContentOffset * 2.0d);
                Height = Math.Min(contentHeight, GetMaxViewportHeight(workspace)) + (ContentOffset * 2.0d);
            }
            finally
            {
                _suppressToolbarRefresh = false;
            }
        }

        /// <summary>
        /// 当前 DIP 与物理像素的换算比（125% 缩放时为 1.25）。
        /// </summary>
        private double GetDpiScale()
        {
            var source = PresentationSource.FromVisual(this);
            if (source == null)
            {
                // 首次 SetImage 时窗口可能还没创建句柄，先确保句柄存在才能读到准确的 DPI。
                try
                {
                    new System.Windows.Interop.WindowInteropHelper(this).EnsureHandle();
                }
                catch
                {
                }

                source = PresentationSource.FromVisual(this);
            }

            if (source != null && source.CompositionTarget != null)
            {
                var matrix = source.CompositionTarget.TransformToDevice;
                if (matrix.M11 > 0d)
                {
                    return matrix.M11;
                }
            }

            try
            {
                using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    if (graphics.DpiX > 0f)
                    {
                        return graphics.DpiX / 96.0d;
                    }
                }
            }
            catch
            {
            }

            return 1.0d;
        }

        /// <summary>
        /// 鼠标所在显示器的工作区（DIP），即去掉任务栏后的可用区域。
        /// </summary>
        private Rect GetTargetWorkspace()
        {
            var fallback = new Rect(
                SystemParameters.WorkArea.Left,
                SystemParameters.WorkArea.Top,
                SystemParameters.WorkArea.Width,
                SystemParameters.WorkArea.Height);

            try
            {
                var screen = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position);
                var workingArea = screen.WorkingArea;
                if (workingArea.Width <= 0 || workingArea.Height <= 0)
                {
                    return fallback;
                }

                var dpi = GetDpiScale();
                return new Rect(
                    workingArea.Left / dpi,
                    workingArea.Top / dpi,
                    workingArea.Width / dpi,
                    workingArea.Height / dpi);
            }
            catch
            {
                return fallback;
            }
        }

        /// <summary>
        /// 贴图窗口当前所在显示器的工作区（DIP）。
        /// </summary>
        private Rect GetWindowWorkspace()
        {
            var fallback = GetTargetWorkspace();

            try
            {
                var handle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
                if (handle == IntPtr.Zero)
                {
                    return fallback;
                }

                var screen = System.Windows.Forms.Screen.FromHandle(handle);
                var workingArea = screen.WorkingArea;
                if (workingArea.Width <= 0 || workingArea.Height <= 0)
                {
                    return fallback;
                }

                var dpi = GetDpiScale();
                return new Rect(
                    workingArea.Left / dpi,
                    workingArea.Top / dpi,
                    workingArea.Width / dpi,
                    workingArea.Height / dpi);
            }
            catch
            {
                return fallback;
            }
        }

        /// <summary>
        /// 把平移偏移限制在 [视口 - 内容, 0]，保证内容既不会被拖出视口，也不会露出空白。
        /// </summary>
        private static double ClampPan(double offset, double viewportLength, double contentLength)
        {
            var minOffset = Math.Min(0d, viewportLength - contentLength);
            return Math.Max(minOffset, Math.Min(0d, offset));
        }

        /// <summary>
        /// 以窗口内某点为锚点缩放：锚点下的图像位置保持不动。
        /// 内容超出工作区时窗口不再变大，改为贴齐工作区并用平移保持锚点。
        /// </summary>
        private void ChangeZoom(double factor, Point anchorInWindow)
        {
            if (_baseImageWidth <= 0d || _baseImageHeight <= 0d)
            {
                return;
            }

            var oldZoom = _zoom;
            var newZoom = Math.Max(MinZoom, Math.Min(MaxZoom, oldZoom * factor));
            if (Math.Abs(newZoom - oldZoom) < 0.0001d)
            {
                return;
            }

            // 锚点在图片基准坐标系（未缩放）中的位置，旧的平移偏移要计入。
            var imageX = (anchorInWindow.X - Left - ContentOffset - _panX) / oldZoom;
            var imageY = (anchorInWindow.Y - Top - ContentOffset - _panY) / oldZoom;

            var workspace = GetWindowWorkspace();
            var contentWidth = Math.Max(1.0d, _baseImageWidth * newZoom);
            var contentHeight = Math.Max(1.0d, _baseImageHeight * newZoom);

            _zoom = newZoom;
            ApplyWindowSize(workspace);

            // 视口按窗口的实际尺寸取，系统对窗口尺寸的限制也能被正确考虑。
            var viewportWidth = Math.Max(1.0d, Width - (ContentOffset * 2.0d));
            var viewportHeight = Math.Max(1.0d, Height - (ContentOffset * 2.0d));

            if (contentWidth > viewportWidth + 0.01d)
            {
                // 横向超出：窗口已撑到工作区宽度，位置尽量保持不动，用平移让锚点下的像素停在原处。
                Left = ClampWindowOrigin(Left, workspace.Left, workspace.Width, Width);
                _panX = ClampPan(
                    anchorInWindow.X - Left - ContentOffset - (imageX * newZoom),
                    viewportWidth,
                    contentWidth);
            }
            else
            {
                // 完整放得下：回到"窗口等于整张图"的原始行为，并留在工作区内。
                _panX = 0d;
                Left = ClampWindowOrigin(
                    anchorInWindow.X - ContentOffset - (imageX * newZoom),
                    workspace.Left,
                    workspace.Width,
                    Width);
            }

            if (contentHeight > viewportHeight + 0.01d)
            {
                Top = ClampWindowOrigin(Top, workspace.Top, workspace.Height, Height);
                _panY = ClampPan(
                    anchorInWindow.Y - Top - ContentOffset - (imageY * newZoom),
                    viewportHeight,
                    contentHeight);
            }
            else
            {
                _panY = 0d;
                Top = ClampWindowOrigin(
                    anchorInWindow.Y - ContentOffset - (imageY * newZoom),
                    workspace.Top,
                    workspace.Height,
                    Height);
            }

            ApplyVisualSize();

            // 缩放到超出/回落到视口内时，左键的含义会变化（平移图片 / 拖动窗口），光标要跟着切换。
            UpdateCursorVisual();
        }

        /// <summary>窗口原点限制在工作区内，避免缩小后贴图跑到屏幕外。</summary>
        private static double ClampWindowOrigin(double origin, double workspaceOrigin, double workspaceLength, double windowLength)
        {
            var maxOrigin = workspaceOrigin + workspaceLength - windowLength;
            if (maxOrigin < workspaceOrigin)
            {
                return workspaceOrigin;
            }

            return Math.Max(workspaceOrigin, Math.Min(maxOrigin, origin));
        }

        private void RestoreOriginalSize()
        {
            if (_baseImageWidth <= 0d || _baseImageHeight <= 0d)
            {
                return;
            }

            // 以窗口中心为锚点回到 1:1 原尺寸。
            var center = new Point(Left + (Width / 2.0d), Top + (Height / 2.0d));
            var factor = 1.0d / _zoom;
            ChangeZoom(factor, center);
        }

        #endregion

        #region 工具栏显示控制

        private void ShowToolbar()
        {
            StopToolbarHideTimer();

            if (!ToolbarPopup.IsOpen)
            {
                ToolbarPopup.IsOpen = true;
            }

            RefreshToolbarPlacement();
        }

        private void ScheduleToolbarHide()
        {
            if (_toolbarHideTimer == null)
            {
                _toolbarHideTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(ToolbarHideDelayMilliseconds)
                };
                _toolbarHideTimer.Tick += ToolbarHideTimer_Tick;
            }

            _toolbarHideTimer.Stop();
            _toolbarHideTimer.Start();
        }

        private void StopToolbarHideTimer()
        {
            if (_toolbarHideTimer != null)
            {
                _toolbarHideTimer.Stop();
            }
        }

        private void ToolbarHideTimer_Tick(object sender, EventArgs e)
        {
            StopToolbarHideTimer();

            // 鼠标正停在工具栏上时保持显示（点击工具栏按钮会让主窗口失活，不能据此关闭）。
            if (IsMouseOverToolbar())
            {
                return;
            }

            // 色板弹层依附在工具栏上，工具栏关闭会连带关闭色板。
            if (BackgroundColorPopup.IsOpen)
            {
                return;
            }

            // 切到 Word 或其他程序后，工具栏不应继续浮在最上层。
            if (!IsActive && _activeTextEditor == null)
            {
                ToolbarPopup.IsOpen = false;
                return;
            }

            // 标注中或正在输入文字时保持显示，避免操作途中工具栏消失。
            if (ShouldKeepToolbarVisible())
            {
                return;
            }

            ToolbarPopup.IsOpen = false;
        }

        /// <summary>
        /// 正在标注、编辑文字或选择底色时，工具栏保持可见，避免操作中途消失。
        /// </summary>
        private bool ShouldKeepToolbarVisible()
        {
            return _currentTool != AnnotationTool.None ||
                   _activeTextEditor != null ||
                   BackgroundColorPopup.IsOpen;
        }

        private bool IsMouseOverToolbar()
        {
            if (!ToolbarPopup.IsOpen)
            {
                return false;
            }

            return ToolbarPanel.IsMouseOver;
        }

        /// <summary>
        /// 窗口移动或缩放后，微调 Popup 偏移强制重新计算位置，让工具栏跟随贴图。
        /// </summary>
        private void RefreshToolbarPlacement()
        {
            if (!ToolbarPopup.IsOpen || _suppressToolbarRefresh)
            {
                return;
            }

            var offset = ToolbarPopup.HorizontalOffset;
            ToolbarPopup.HorizontalOffset = offset + 0.1d;
            ToolbarPopup.HorizontalOffset = offset;
        }

        private void Window_MouseEnter(object sender, MouseEventArgs e)
        {
            ShowToolbar();
        }

        private void Window_MouseLeave(object sender, MouseEventArgs e)
        {
            // 点击工具栏按钮时鼠标会离开主窗口，但工具栏仍处于鼠标悬停状态；
            // 这种情况下不隐藏，交给工具栏自身的 MouseLeave 决定。
            if (IsMouseOverToolbar())
            {
                return;
            }

            if (ShouldKeepToolbarVisible())
            {
                return;
            }

            ScheduleToolbarHide();
        }

        private void ToolbarPanel_MouseLeave(object sender, MouseEventArgs e)
        {
            if (ShouldKeepToolbarVisible())
            {
                return;
            }

            ScheduleToolbarHide();
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            // 点击工具栏按钮（尤其是参数行的色块 / 字号按钮）同样会触发失活。
            // 此时不能提交文字编辑，否则"编辑中改样式"就没机会生效了。
            if (IsMouseOverToolbar())
            {
                return;
            }

            CommitOrCancelTextEditing();

            // 切换到 Word 或其他程序时工具栏不应继续浮在最上层；
            // 但要留出时间给工具栏按钮的点击处理。
            if (ShouldKeepToolbarVisible())
            {
                return;
            }

            ScheduleToolbarHide();
        }

        /// <summary>
        /// 把焦点交还给贴图本体，避免工具栏按钮抢走焦点后滚轮/键盘失效。
        /// </summary>
        private void RefocusPin()
        {
            try
            {
                if (!IsVisible)
                {
                    return;
                }

                Activate();

                // 正在编辑文字时，把焦点还给输入框，否则用户改完参数就无法继续打字。
                if (_activeTextEditor != null)
                {
                    _activeTextEditor.Focus();
                }
            }
            catch
            {
            }
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            RefreshToolbarPlacement();
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            RefreshToolbarPlacement();
        }

        #endregion

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.Z)
            {
                CommitOrCancelTextEditing();
                UndoLastAnnotation();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                if (BackgroundColorPopup.IsOpen)
                {
                    BackgroundColorPopup.IsOpen = false;
                    e.Handled = true;
                    return;
                }

                if (_currentTool != AnnotationTool.None)
                {
                    CancelTextEditing();
                    ClearToolSelection();
                    e.Handled = true;
                    return;
                }

                ClosePin();
                e.Handled = true;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var point = e.GetPosition(RootGrid);
            if (IsPointInsideActiveTextEditor(point))
            {
                return;
            }

            CommitOrCancelTextEditing();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            ClosePin();
        }

        /// <summary>
        /// 工具栏按钮点击后把焦点交还给贴图窗口，保证 Esc / Ctrl+Z 仍然可用。
        /// 注意：Popup 中的按钮点击会让主窗口失活，因此每次点击后都重新激活。
        /// </summary>
        private void ToolbarActionCompleted()
        {
            RefocusPin();
            ShowToolbar();
        }

        private void ClosePin()
        {
            if (_isClosing)
            {
                return;
            }

            CommitOrCancelTextEditing();
            ClearAnnotations();
            _isClosing = true;
            Close();
        }

        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();
            UndoLastAnnotation();
            ToolbarActionCompleted();
        }

        private void RestoreSizeButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();
            RestoreOriginalSize();
            ToolbarActionCompleted();
        }

        private void BackgroundColorButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            BackgroundColorPopup.IsOpen = true;
        }

        private void BackgroundSwatch_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null)
            {
                return;
            }

            var color = button.Tag as string;
            if (string.IsNullOrWhiteSpace(color))
            {
                return;
            }

            ApplyBackgroundColor(color);
            SaveBackgroundColor(color);
            BackgroundColorPopup.IsOpen = false;
            ToolbarActionCompleted();
        }

        private void CustomBackgroundColorButton_Click(object sender, RoutedEventArgs e)
        {
            BackgroundColorPopup.IsOpen = false;
            CommitOrCancelTextEditing();

            try
            {
                using (var dialog = new System.Windows.Forms.ColorDialog
                {
                    AllowFullOpen = true,
                    FullOpen = true,
                    AnyColor = true,
                    Color = ToDrawingColor(ImageBackdrop.Fill)
                })
                {
                    var helper = new System.Windows.Interop.WindowInteropHelper(this);
                    var result = dialog.ShowDialog(new NativeWindowOwner(helper.Handle));
                    if (result != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    var color = string.Format("#{0:X2}{1:X2}{2:X2}", dialog.Color.R, dialog.Color.G, dialog.Color.B);
                    ApplyBackgroundColor(color);
                    SaveBackgroundColor(color);
                }
            }
            catch
            {
            }
            finally
            {
                ToolbarActionCompleted();
            }
        }

        private void ApplyBackgroundColor(string colorText)
        {
            if (string.IsNullOrWhiteSpace(colorText))
            {
                colorText = DefaultBackgroundColor;
            }

            var fill = CreateBackgroundBrush(colorText);
            ImageBackdrop.Fill = fill;
            UpdateBackgroundColorPreview(fill);
        }

        private static Brush CreateBackgroundBrush(string colorText)
        {
            if (string.Equals(colorText, TransparentBackgroundColor, StringComparison.OrdinalIgnoreCase))
            {
                // 完全透明的画刷在分层窗口里不参与命中测试，用 1/255 的白色保证整张图可拖动、可双击。
                return CreateFrozenBrush(Color.FromArgb(1, 255, 255, 255));
            }

            var color = ParseColor(colorText) ?? ParseColor(DefaultBackgroundColor).Value;
            return CreateFrozenBrush(color);
        }

        private static Brush CreateFrozenBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        private static Color? ParseColor(string colorText)
        {
            if (string.IsNullOrWhiteSpace(colorText))
            {
                return null;
            }

            try
            {
                var converted = ColorConverter.ConvertFromString(colorText.Trim());
                if (converted is Color)
                {
                    return (Color)converted;
                }
            }
            catch (FormatException)
            {
            }
            catch (NotSupportedException)
            {
            }

            return null;
        }

        private void UpdateBackgroundColorPreview(Brush fill)
        {
            var solid = fill as SolidColorBrush;
            if (solid == null || solid.Color.A <= 1)
            {
                // 透明时用棋盘格图标表示
                BackgroundColorPreview.Fill = TryFindResource("CheckerboardBrush") as Brush;
                return;
            }

            BackgroundColorPreview.Fill = CreateFrozenBrush(solid.Color);
        }

        private static System.Drawing.Color ToDrawingColor(Brush brush)
        {
            var solid = brush as SolidColorBrush;
            if (solid == null || solid.Color.A <= 1)
            {
                return System.Drawing.Color.Empty;
            }

            var color = solid.Color;
            return System.Drawing.Color.FromArgb(color.R, color.G, color.B);
        }

        private static string LoadSavedBackgroundColor()
        {
            try
            {
                var saved = Settings.Default.PreviewBackgroundColor;
                return string.IsNullOrWhiteSpace(saved) ? DefaultBackgroundColor : saved;
            }
            catch
            {
                return DefaultBackgroundColor;
            }
        }

        private static void SaveBackgroundColor(string colorText)
        {
            try
            {
                Settings.Default.PreviewBackgroundColor = colorText;
                Settings.Default.Save();
            }
            catch
            {
            }
        }

        #region 标注参数（颜色 / 线宽 / 字号）

        private static Color LoadSavedAnnotationColor()
        {
            try
            {
                var parsed = ParseColor(Settings.Default.AnnotationColor);
                if (parsed.HasValue)
                {
                    return parsed.Value;
                }
            }
            catch
            {
            }

            return ParseColor(DefaultAnnotationColor).Value;
        }

        private static double LoadSavedAnnotationStrokeWidth()
        {
            try
            {
                return Settings.Default.AnnotationStrokeWidth;
            }
            catch
            {
                return DefaultAnnotationStrokeWidth;
            }
        }

        private static double LoadSavedAnnotationFontSize()
        {
            try
            {
                return Settings.Default.AnnotationFontSize;
            }
            catch
            {
                return DefaultAnnotationFontSize;
            }
        }

        private static double ClampStrokeWidth(double width)
        {
            if (double.IsNaN(width) || width <= 0d)
            {
                return DefaultAnnotationStrokeWidth;
            }

            return Math.Max(MinAnnotationStrokeWidth, Math.Min(MaxAnnotationStrokeWidth, width));
        }

        /// <summary>把任意字号对齐到最接近的预设档位。</summary>
        private static double NormalizeFontSize(double fontSize)
        {
            if (double.IsNaN(fontSize) || fontSize <= 0d)
            {
                return DefaultAnnotationFontSize;
            }

            var clamped = Math.Max(MinAnnotationFontSize, Math.Min(MaxAnnotationFontSize, fontSize));
            var best = AnnotationFontSizeSteps[0];
            var bestDelta = Math.Abs(clamped - best);

            foreach (var step in AnnotationFontSizeSteps)
            {
                var delta = Math.Abs(clamped - step);
                if (delta < bestDelta)
                {
                    best = step;
                    bestDelta = delta;
                }
            }

            return best;
        }

        private void SaveAnnotationOptions()
        {
            try
            {
                Settings.Default.AnnotationColor = ToHexString(_annotationColor);
                Settings.Default.AnnotationStrokeWidth = _annotationStrokeWidth;
                Settings.Default.AnnotationFontSize = _annotationFontSize;
                Settings.Default.Save();
            }
            catch
            {
            }
        }

        private static string ToHexString(Color color)
        {
            return string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
        }

        /// <summary>按当前工具切换参数行以及线宽/字号子面板的可见性。</summary>
        private void UpdateToolOptionsPanel()
        {
            if (ToolOptionsPanel == null)
            {
                return;
            }

            var hasTool = _currentTool != AnnotationTool.None;
            ToolOptionsPanel.Visibility = hasTool ? Visibility.Visible : Visibility.Collapsed;

            var isText = _currentTool == AnnotationTool.Text;
            StrokeWidthPanel.Visibility = hasTool && !isText ? Visibility.Visible : Visibility.Collapsed;
            FontSizePanel.Visibility = hasTool && isText ? Visibility.Visible : Visibility.Collapsed;

            if (hasTool)
            {
                UpdateAnnotationColorSelection();
                UpdateStrokeWidthSelection();
                UpdateFontSizeSelection();
            }

            RefreshToolbarPlacement();
        }

        private void UpdateAnnotationColorSelection()
        {
            if (AnnotationColorPanel == null)
            {
                return;
            }

            var current = ToHexString(_annotationColor);

            foreach (var child in AnnotationColorPanel.Children)
            {
                var swatch = child as Button;
                if (swatch == null)
                {
                    continue;
                }

                var tag = swatch.Tag as string;
                var isCurrent = tag != null && string.Equals(tag, current, StringComparison.OrdinalIgnoreCase);
                swatch.BorderBrush = isCurrent
                    ? CreateFrozenBrush(Color.FromRgb(0x37, 0x41, 0x51))
                    : CreateFrozenBrush(Color.FromRgb(0xD0, 0xD4, 0xDA));
                swatch.BorderThickness = new Thickness(isCurrent ? 2 : 1);
            }
        }

        private void UpdateStrokeWidthSelection()
        {
            var width = (int)Math.Round(_annotationStrokeWidth);
            SetStrokeWidthButtonState(ThinStrokeButton, width == 2);
            SetStrokeWidthButtonState(MediumStrokeButton, width == 4);
            SetStrokeWidthButtonState(ThickStrokeButton, width == 6);
        }

        private static void SetStrokeWidthButtonState(ToggleButton button, bool isChecked)
        {
            if (button == null)
            {
                return;
            }

            button.IsChecked = isChecked;
        }

        private void UpdateFontSizeSelection()
        {
            if (FontSizeValueText != null)
            {
                FontSizeValueText.Text = ((int)Math.Round(_annotationFontSize)).ToString();
            }

            var atMin = _annotationFontSize <= AnnotationFontSizeSteps[0] + 0.01d;
            var atMax = _annotationFontSize >= AnnotationFontSizeSteps[AnnotationFontSizeSteps.Length - 1] - 0.01d;
            if (FontSizeDecreaseButton != null)
            {
                FontSizeDecreaseButton.IsEnabled = !atMin;
                FontSizeDecreaseButton.Opacity = atMin ? 0.35d : 1.0d;
            }

            if (FontSizeIncreaseButton != null)
            {
                FontSizeIncreaseButton.IsEnabled = !atMax;
                FontSizeIncreaseButton.Opacity = atMax ? 0.35d : 1.0d;
            }
        }

        private void AnnotationColorSwatch_Click(object sender, RoutedEventArgs e)
        {
            var swatch = sender as Button;
            var tag = swatch == null ? null : swatch.Tag as string;
            var color = ParseColor(tag);
            if (!color.HasValue)
            {
                return;
            }

            _annotationColor = color.Value;
            SaveAnnotationOptions();
            UpdateAnnotationColorSelection();
            ApplyAnnotationStyleToActiveTextEditor();
            ToolbarActionCompleted();
        }

        private void StrokeWidthButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as ToggleButton;
            if (button == null)
            {
                return;
            }

            double width;
            if (!double.TryParse((button.Tag as string) ?? string.Empty, out width))
            {
                return;
            }

            _annotationStrokeWidth = ClampStrokeWidth(width);
            SaveAnnotationOptions();
            UpdateStrokeWidthSelection();
            ToolbarActionCompleted();
        }

        private void FontSizeDecreaseButton_Click(object sender, RoutedEventArgs e)
        {
            StepAnnotationFontSize(-1);
        }

        private void FontSizeIncreaseButton_Click(object sender, RoutedEventArgs e)
        {
            StepAnnotationFontSize(1);
        }

        private void StepAnnotationFontSize(int direction)
        {
            var index = Array.FindIndex(AnnotationFontSizeSteps, s => Math.Abs(s - _annotationFontSize) < 0.01d);
            if (index < 0)
            {
                index = Array.FindIndex(AnnotationFontSizeSteps, s => s >= _annotationFontSize);
                if (index < 0)
                {
                    index = AnnotationFontSizeSteps.Length - 1;
                }

                if (direction < 0)
                {
                    index = Math.Max(0, index - 1);
                }
            }
            else
            {
                index = Math.Max(0, Math.Min(AnnotationFontSizeSteps.Length - 1, index + direction));
            }

            _annotationFontSize = AnnotationFontSizeSteps[index];
            SaveAnnotationOptions();
            UpdateFontSizeSelection();
            ApplyAnnotationStyleToActiveTextEditor();
            ToolbarActionCompleted();
        }

        /// <summary>正在编辑的文字立即应用当前颜色与字号，并同步调整输入框尺寸。</summary>
        private void ApplyAnnotationStyleToActiveTextEditor()
        {
            if (_activeTextEditor == null || _activeTextEditorContainer == null)
            {
                return;
            }

            var brush = CreateFrozenBrush(_annotationColor);
            _activeTextEditor.Foreground = brush;
            _activeTextEditor.CaretBrush = brush;
            _activeTextEditor.FontSize = _annotationFontSize;

            _activeTextEditorContainer.MinHeight = GetTextContainerMinHeight(_annotationFontSize);
            _activeTextEditorContainer.Width = GetTextContainerWidth(_annotationFontSize);
        }

        private static double GetTextContainerMinHeight(double fontSize)
        {
            return (fontSize * 1.6d) + TextBoxVerticalPadding;
        }

        private static double GetTextContainerWidth(double fontSize)
        {
            return Math.Max(200.0d, fontSize * 10.0d);
        }

        /// <summary>新建标注时使用的冻结画刷。</summary>
        private Brush CreateAnnotationBrush()
        {
            return CreateFrozenBrush(_annotationColor);
        }

        #endregion

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();

            if (_bitmap == null)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "PNG Image|*.png|JPEG Image|*.jpg;*.jpeg|BMP Image|*.bmp",
                AddExtension = true,
                DefaultExt = ".png",
                FileName = "word-picture.png"
            };

            try
            {
                if (dialog.ShowDialog(this) != true)
                {
                    return;
                }

                using (var savedBitmap = BuildSavedBitmap())
                {
                    if (savedBitmap == null)
                    {
                        return;
                    }

                    var extension = System.IO.Path.GetExtension(dialog.FileName).ToLowerInvariant();
                    var imageFormat = System.Drawing.Imaging.ImageFormat.Png;
                    switch (extension)
                    {
                        case ".jpg":
                        case ".jpeg":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                            break;
                        case ".bmp":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
                            break;
                    }

                    savedBitmap.Save(dialog.FileName, imageFormat);
                }
            }
            catch
            {
            }
            finally
            {
                ToolbarActionCompleted();
            }
        }

        private DrawingBitmap BuildSavedBitmap()
        {
            if (_bitmap == null || _baseImageWidth <= 0 || _baseImageHeight <= 0)
            {
                return null;
            }

            var output = new DrawingBitmap(_bitmap);
            var scaleX = output.Width / _baseImageWidth;
            var scaleY = output.Height / _baseImageHeight;

            using (var graphics = System.Drawing.Graphics.FromImage(output))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                foreach (var element in _annotationHistory)
                {
                    DrawAnnotationElement(graphics, element, scaleX, scaleY);
                }
            }

            return output;
        }

        private void DrawAnnotationElement(System.Drawing.Graphics graphics, UIElement element, double scaleX, double scaleY)
        {
            if (element is Line line)
            {
                using (var pen = CreateDrawingPen(line.Stroke, line.StrokeThickness, scaleX))
                {
                    graphics.DrawLine(
                        pen,
                        (float)(line.X1 * scaleX),
                        (float)(line.Y1 * scaleY),
                        (float)(line.X2 * scaleX),
                        (float)(line.Y2 * scaleY));
                }

                return;
            }

            if (element is Rectangle rectangle)
            {
                using (var pen = CreateDrawingPen(rectangle.Stroke, rectangle.StrokeThickness, scaleX))
                {
                    var left = Canvas.GetLeft(rectangle) * scaleX;
                    var top = Canvas.GetTop(rectangle) * scaleY;
                    var width = rectangle.Width * scaleX;
                    var height = rectangle.Height * scaleY;
                    graphics.DrawRectangle(pen, (float)left, (float)top, (float)width, (float)height);
                }

                return;
            }

            if (element is Ellipse ellipse)
            {
                using (var pen = CreateDrawingPen(ellipse.Stroke, ellipse.StrokeThickness, scaleX))
                {
                    var left = Canvas.GetLeft(ellipse) * scaleX;
                    var top = Canvas.GetTop(ellipse) * scaleY;
                    var width = ellipse.Width * scaleX;
                    var height = ellipse.Height * scaleY;
                    graphics.DrawEllipse(pen, (float)left, (float)top, (float)width, (float)height);
                }

                return;
            }

            if (element is Path path)
            {
                DrawPathAnnotation(graphics, path, scaleX, scaleY);
                return;
            }

            if (element is Grid grid)
            {
                DrawTextAnnotation(graphics, grid, scaleX, scaleY);
            }
        }

        private void DrawPathAnnotation(System.Drawing.Graphics graphics, Path path, double scaleX, double scaleY)
        {
            var geometry = path.Data as GeometryGroup;
            if (geometry == null)
            {
                return;
            }

            using (var pen = CreateDrawingPen(path.Stroke, path.StrokeThickness, scaleX))
            {
                foreach (var child in geometry.Children)
                {
                    if (child is LineGeometry line)
                    {
                        graphics.DrawLine(
                            pen,
                            (float)(line.StartPoint.X * scaleX),
                            (float)(line.StartPoint.Y * scaleY),
                            (float)(line.EndPoint.X * scaleX),
                            (float)(line.EndPoint.Y * scaleY));
                    }
                }
            }
        }

        private void DrawTextAnnotation(System.Drawing.Graphics graphics, Grid grid, double scaleX, double scaleY)
        {
            if (grid.Children.Count == 0 || !(grid.Children[0] is TextBlock textBlock))
            {
                return;
            }

            var left = Canvas.GetLeft(grid);
            var top = Canvas.GetTop(grid);
            var width = grid.Width;

            // 颜色和字号都取自 TextBlock 自身，界面与导出保持一致。
            var textColor = textBlock.Foreground is SolidColorBrush solid
                ? System.Drawing.Color.FromArgb(
                    textBlock.Opacity >= 1.0d ? solid.Color.A : (byte)(solid.Color.A * textBlock.Opacity),
                    solid.Color.R,
                    solid.Color.G,
                    solid.Color.B)
                : System.Drawing.Color.OrangeRed;

            var fontSize = (float)(Math.Max(1.0d, textBlock.FontSize) * scaleY);

            using (var brush = new System.Drawing.SolidBrush(textColor))
            using (var font = new System.Drawing.Font("Microsoft YaHei UI", fontSize, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel))
            {
                var layoutRect = new System.Drawing.RectangleF(
                    (float)((left + TextBoxHorizontalPadding / 2.0d) * scaleX),
                    (float)((top + TextBoxVerticalPadding / 2.0d) * scaleY),
                    (float)Math.Max(1, (width - TextBoxHorizontalPadding) * scaleX),
                    10000f);

                var format = new System.Drawing.StringFormat
                {
                    Alignment = System.Drawing.StringAlignment.Near,
                    LineAlignment = System.Drawing.StringAlignment.Near
                };

                graphics.DrawString(textBlock.Text, font, brush, layoutRect, format);
            }
        }

        /// <summary>按标注元素自身的颜色与线宽创建 GDI+ 画笔，导出时与界面保持一致。</summary>
        private static System.Drawing.Pen CreateDrawingPen(Brush stroke, double thickness, double scale)
        {
            var drawingColor = stroke is SolidColorBrush solid
                ? System.Drawing.Color.FromArgb(solid.Color.A, solid.Color.R, solid.Color.G, solid.Color.B)
                : System.Drawing.Color.OrangeRed;

            var penWidth = Math.Max(1.0d, thickness * scale);
            return new System.Drawing.Pen(drawingColor, (float)penWidth)
            {
                StartCap = System.Drawing.Drawing2D.LineCap.Round,
                EndCap = System.Drawing.Drawing2D.LineCap.Round,
                LineJoin = System.Drawing.Drawing2D.LineJoin.Round
            };
        }

        private void ToolButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();

            var clicked = sender as ToggleButton;
            if (clicked == null)
            {
                return;
            }

            var isChecked = clicked.IsChecked == true;
            SetToolButtonState(LineToolButton, clicked == LineToolButton && isChecked);
            SetToolButtonState(RectangleToolButton, clicked == RectangleToolButton && isChecked);
            SetToolButtonState(EllipseToolButton, clicked == EllipseToolButton && isChecked);
            SetToolButtonState(ArrowToolButton, clicked == ArrowToolButton && isChecked);
            SetToolButtonState(TextToolButton, clicked == TextToolButton && isChecked);

            if (!isChecked)
            {
                _currentTool = AnnotationTool.None;
                AnnotationCanvas.IsHitTestVisible = false;
                UpdateCursorVisual();
                UpdateToolOptionsPanel();
                ScheduleToolbarHide();
                return;
            }

            switch ((string)clicked.Tag)
            {
                case "Line":
                    _currentTool = AnnotationTool.Line;
                    break;
                case "Rectangle":
                    _currentTool = AnnotationTool.Rectangle;
                    break;
                case "Ellipse":
                    _currentTool = AnnotationTool.Ellipse;
                    break;
                case "Arrow":
                    _currentTool = AnnotationTool.Arrow;
                    break;
                case "Text":
                    _currentTool = AnnotationTool.Text;
                    break;
                default:
                    _currentTool = AnnotationTool.None;
                    break;
            }

            AnnotationCanvas.IsHitTestVisible = _currentTool != AnnotationTool.None;
            UpdateCursorVisual();
            UpdateToolOptionsPanel();
            ShowToolbar();
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_bitmap == null)
            {
                return;
            }

            var factor = e.Delta > 0 ? ZoomStepFactor : 1.0d / ZoomStepFactor;
            BeginImageInteraction();
            ChangeZoom(factor, e.GetPosition(this));
            EndImageInteraction();
            e.Handled = true;
        }

        private void PreviewImage_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            Window_MouseWheel(sender, e);
        }

        /// <summary>
        /// 缩放/拖动期间临时降低位图缩放质量，交互结束后恢复，减轻渲染压力。
        /// </summary>
        private void BeginImageInteraction()
        {
            _interactionDepth++;
            if (_interactionDepth == 1)
            {
                RenderOptions.SetBitmapScalingMode(PreviewImage, BitmapScalingMode.LowQuality);
            }

            if (_interactionEndTimer == null)
            {
                _interactionEndTimer = new DispatcherTimer(
                    DispatcherPriority.Background,
                    Dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(140)
                };
                _interactionEndTimer.Tick += InteractionEndTimer_Tick;
            }

            _interactionEndTimer.Stop();
            _interactionEndTimer.Start();
        }

        private void EndImageInteraction()
        {
            if (_interactionDepth > 0)
            {
                _interactionDepth--;
            }
        }

        private void InteractionEndTimer_Tick(object sender, EventArgs e)
        {
            if (_interactionEndTimer != null)
            {
                _interactionEndTimer.Stop();
            }

            _interactionDepth = 0;
            RenderOptions.SetBitmapScalingMode(PreviewImage, BitmapScalingMode.HighQuality);
            RefreshToolbarPlacement();
        }

        private void PreviewImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2 && _currentTool == AnnotationTool.None)
            {
                // 双击图片直接关闭贴图（Snipaste 习惯）
                e.Handled = true;
                ClosePin();
                return;
            }

            if (_currentTool != AnnotationTool.None)
            {
                return;
            }

            CommitOrCancelTextEditing();

            if (e.ButtonState == MouseButtonState.Pressed)
            {
                if (HasOverflowingContent() && (Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                {
                    // 图片超出视口：左键改为平移内容，避免窗口无法移动导致看不到其余部分。
                    // 仍想移动窗口时按住 Ctrl。
                    StartPanning(e);
                    e.Handled = true;
                    return;
                }

                try
                {
                    // 交给系统拖动，移动过程平滑且能跨屏
                    DragMove();
                }
                catch (InvalidOperationException)
                {
                }
            }

            RefreshToolbarPlacement();
            e.Handled = true;
        }

        /// <summary>当前缩放下的内容是否大于视口的任一方向。</summary>
        private bool HasOverflowingContent()
        {
            if (_baseImageWidth <= 0d || _baseImageHeight <= 0d)
            {
                return false;
            }

            var viewportWidth = Width - (ContentOffset * 2.0d);
            var viewportHeight = Height - (ContentOffset * 2.0d);

            // 窗口尺寸尚未确定（首次布局）时回退到工作区尺寸。
            if (double.IsNaN(viewportWidth) || double.IsNaN(viewportHeight) ||
                viewportWidth <= 0d || viewportHeight <= 0d)
            {
                var workspace = GetWindowWorkspace();
                viewportWidth = GetMaxViewportWidth(workspace);
                viewportHeight = GetMaxViewportHeight(workspace);
            }

            return (_baseImageWidth * _zoom) > viewportWidth + 0.01d ||
                   (_baseImageHeight * _zoom) > viewportHeight + 0.01d;
        }

        private void StartPanning(MouseButtonEventArgs e)
        {
            _isPanning = true;
            _panStartMouse = e.GetPosition(this);
            _panStartOffsetX = _panX;
            _panStartOffsetY = _panY;

            BeginImageInteraction();
            CaptureMouse();
            ImageCanvas.Cursor = Cursors.ScrollAll;
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isPanning)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndPanning();
                return;
            }

            var current = e.GetPosition(this);
            var viewportWidth = Width - (ContentOffset * 2.0d);
            var viewportHeight = Height - (ContentOffset * 2.0d);

            _panX = ClampPan(
                _panStartOffsetX + (current.X - _panStartMouse.X),
                viewportWidth,
                _baseImageWidth * _zoom);
            _panY = ClampPan(
                _panStartOffsetY + (current.Y - _panStartMouse.Y),
                viewportHeight,
                _baseImageHeight * _zoom);

            ContentPanTransform.X = _panX;
            ContentPanTransform.Y = _panY;
            e.Handled = true;
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isPanning)
            {
                return;
            }

            EndPanning();
            e.Handled = true;
        }

        private void Window_LostMouseCapture(object sender, MouseEventArgs e)
        {
            if (!_isPanning)
            {
                return;
            }

            EndPanning();
        }

        private void EndPanning()
        {
            _isPanning = false;
            ReleaseMouseCapture();
            EndImageInteraction();
            UpdateCursorVisual();
            RefreshToolbarPlacement();
        }

        private void AnnotationCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool == AnnotationTool.None)
            {
                return;
            }

            var point = ToContentCoordinates(e.GetPosition(ContentCanvas));

            if (_currentTool == AnnotationTool.Text)
            {
                AddTextAnnotation(point);
                e.Handled = true;
                return;
            }

            _drawingAnnotation = true;
            _annotationStart = point;
            AnnotationCanvas.CaptureMouse();
            CreatePreviewAnnotation(point);
            e.Handled = true;
        }

        private void AnnotationCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_drawingAnnotation)
            {
                return;
            }

            var point = ToContentCoordinates(e.GetPosition(ContentCanvas));
            UpdatePreviewAnnotation(_annotationStart, point);
        }

        private void AnnotationCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_drawingAnnotation)
            {
                return;
            }

            _drawingAnnotation = false;
            AnnotationCanvas.ReleaseMouseCapture();

            var point = ToContentCoordinates(e.GetPosition(ContentCanvas));
            CommitPreviewAnnotation(_annotationStart, point);
            RemovePreviewAnnotation();
            e.Handled = true;
        }

        /// <summary>标注坐标直接存于图片基准坐标系，无需再反算窗口平移。</summary>
        private static Point ToContentCoordinates(Point canvasPoint)
        {
            return canvasPoint;
        }

        private void SetToolButtonState(ToggleButton button, bool isChecked)
        {
            button.IsChecked = isChecked;
        }

        private void ClearToolSelection()
        {
            _drawingAnnotation = false;
            _currentTool = AnnotationTool.None;
            AnnotationCanvas.IsHitTestVisible = false;
            AnnotationCanvas.ReleaseMouseCapture();
            RemovePreviewAnnotation();
            SetToolButtonState(LineToolButton, false);
            SetToolButtonState(RectangleToolButton, false);
            SetToolButtonState(EllipseToolButton, false);
            SetToolButtonState(ArrowToolButton, false);
            SetToolButtonState(TextToolButton, false);
            UpdateCursorVisual();
            UpdateToolOptionsPanel();
        }

        private void CreatePreviewAnnotation(Point point)
        {
            RemovePreviewAnnotation();

            var stroke = CreateAnnotationBrush();
            var thickness = _annotationStrokeWidth;

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    _previewShape = new Line
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        X1 = point.X,
                        Y1 = point.Y,
                        X2 = point.X,
                        Y2 = point.Y
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Rectangle:
                    _previewShape = new Rectangle
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Ellipse:
                    _previewShape = new Ellipse
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Arrow:
                    _previewArrow = new Path
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewArrow);
                    break;
            }
        }

        private void UpdatePreviewAnnotation(Point start, Point end)
        {
            end = ApplyShiftConstraint(start, end, _currentTool);

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    var line = _previewShape as Line;
                    if (line != null)
                    {
                        line.X1 = start.X;
                        line.Y1 = start.Y;
                        line.X2 = end.X;
                        line.Y2 = end.Y;
                    }
                    break;
                case AnnotationTool.Rectangle:
                    UpdateSizedShape(_previewShape, start, end);
                    break;
                case AnnotationTool.Ellipse:
                    UpdateSizedShape(_previewShape, start, end);
                    break;
                case AnnotationTool.Arrow:
                    if (_previewArrow != null)
                    {
                        _previewArrow.Data = BuildArrowGeometry(start, end, _annotationStrokeWidth);
                    }
                    break;
            }
        }

        private void CommitPreviewAnnotation(Point start, Point end)
        {
            end = ApplyShiftConstraint(start, end, _currentTool);

            if (Math.Abs(end.X - start.X) < 2 && Math.Abs(end.Y - start.Y) < 2)
            {
                return;
            }

            var stroke = CreateAnnotationBrush();
            var thickness = _annotationStrokeWidth;

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    var line = new Line
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        X1 = start.X,
                        Y1 = start.Y,
                        X2 = end.X,
                        Y2 = end.Y
                    };
                    AnnotationCanvas.Children.Add(line);
                    _annotationHistory.Add(line);
                    break;
                case AnnotationTool.Rectangle:
                    var rect = new Rectangle
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(rect);
                    UpdateSizedShape(rect, start, end);
                    _annotationHistory.Add(rect);
                    break;
                case AnnotationTool.Ellipse:
                    var ellipse = new Ellipse
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(ellipse);
                    UpdateSizedShape(ellipse, start, end);
                    _annotationHistory.Add(ellipse);
                    break;
                case AnnotationTool.Arrow:
                    var arrow = new Path
                    {
                        Stroke = stroke,
                        StrokeThickness = thickness,
                        Fill = Brushes.Transparent,
                        Data = BuildArrowGeometry(start, end, thickness)
                    };
                    AnnotationCanvas.Children.Add(arrow);
                    _annotationHistory.Add(arrow);
                    break;
            }
        }

        private void UpdateSizedShape(Shape shape, Point start, Point end)
        {
            if (shape == null)
            {
                return;
            }

            var left = Math.Min(start.X, end.X);
            var top = Math.Min(start.Y, end.Y);
            var width = Math.Abs(end.X - start.X);
            var height = Math.Abs(end.Y - start.Y);

            Canvas.SetLeft(shape, left);
            Canvas.SetTop(shape, top);
            shape.Width = width;
            shape.Height = height;
        }

        private Geometry BuildArrowGeometry(Point start, Point end, double strokeWidth)
        {
            var group = new GeometryGroup();
            group.Children.Add(new LineGeometry(start, end));

            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 0.1d)
            {
                return group;
            }

            var ux = dx / length;
            var uy = dy / length;

            // 箭头头部随线宽放大，避免粗线配小箭头。
            var arrowLength = Math.Max(14.0d, strokeWidth * 4.0d);
            var arrowWidth = Math.Max(6.0d, strokeWidth * 2.0d);

            var left = new Point(
                end.X - arrowLength * ux + arrowWidth * uy,
                end.Y - arrowLength * uy - arrowWidth * ux);
            var right = new Point(
                end.X - arrowLength * ux - arrowWidth * uy,
                end.Y - arrowLength * uy + arrowWidth * ux);

            group.Children.Add(new LineGeometry(end, left));
            group.Children.Add(new LineGeometry(end, right));
            return group;
        }

        private Point ApplyShiftConstraint(Point start, Point end, AnnotationTool tool)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) != ModifierKeys.Shift)
            {
                return end;
            }

            switch (tool)
            {
                case AnnotationTool.Line:
                case AnnotationTool.Arrow:
                    return SnapLineToEightDirections(start, end);
                case AnnotationTool.Rectangle:
                case AnnotationTool.Ellipse:
                    return SnapToSquare(start, end);
                default:
                    return end;
            }
        }

        private static Point SnapLineToEightDirections(Point start, Point end)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 0.001d)
            {
                return end;
            }

            var angle = Math.Atan2(dy, dx);
            var step = Math.PI / 4.0d;
            var snappedAngle = Math.Round(angle / step) * step;

            return new Point(
                start.X + Math.Cos(snappedAngle) * length,
                start.Y + Math.Sin(snappedAngle) * length);
        }

        private static Point SnapToSquare(Point start, Point end)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var size = Math.Max(Math.Abs(dx), Math.Abs(dy));

            return new Point(
                start.X + Math.Sign(dx == 0 ? 1 : dx) * size,
                start.Y + Math.Sign(dy == 0 ? 1 : dy) * size);
        }

        private void AddTextAnnotation(Point point)
        {
            CommitOrCancelTextEditing();

            var brush = CreateAnnotationBrush();
            var fontSize = _annotationFontSize;

            var container = new Grid
            {
                Width = GetTextContainerWidth(fontSize),
                MinHeight = GetTextContainerMinHeight(fontSize)
            };

            var rectangle = new Rectangle
            {
                Stroke = Brushes.Gray,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(28, 255, 255, 255)),
                StrokeDashArray = new DoubleCollection { 4, 3 },
                RadiusX = 4,
                RadiusY = 4
            };

            var textBox = new TextBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Foreground = brush,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontSize = fontSize,
                FontWeight = FontWeights.Bold,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 4, 6, 4),
                CaretBrush = brush
            };

            container.Children.Add(rectangle);
            container.Children.Add(textBox);
            AnnotationCanvas.Children.Add(container);
            Canvas.SetLeft(container, point.X);
            Canvas.SetTop(container, point.Y);

            _activeTextEditorContainer = container;
            _activeTextEditor = textBox;

            textBox.Loaded += delegate
            {
                textBox.Focus();
                textBox.Select(textBox.Text.Length, 0);
            };
            textBox.Focus();
            textBox.Select(textBox.Text.Length, 0);
        }

        private void CommitOrCancelTextEditing()
        {
            if (_activeTextEditorContainer == null || _activeTextEditor == null)
            {
                return;
            }

            var text = _activeTextEditor.Text;

            if (!string.IsNullOrWhiteSpace(text))
            {
                // 沿用输入框当前的样式，避免提交瞬间颜色或字号跳变。
                // 不设 Width，让 TextBlock 自动填满容器；Margin 与 TextBox 的 Padding 对齐即可保证换行宽度一致。
                var textBlock = new TextBlock
                {
                    Text = text,
                    Foreground = _activeTextEditor.Foreground,
                    FontFamily = _activeTextEditor.FontFamily,
                    FontSize = _activeTextEditor.FontSize,
                    FontWeight = _activeTextEditor.FontWeight,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(6, 4, 6, 4)
                };

                _activeTextEditorContainer.Children.Clear();
                _activeTextEditorContainer.Children.Add(textBlock);
                _annotationHistory.Add(_activeTextEditorContainer);
            }
            else
            {
                AnnotationCanvas.Children.Remove(_activeTextEditorContainer);
            }

            _activeTextEditor = null;
            _activeTextEditorContainer = null;
        }

        private void CancelTextEditing()
        {
            if (_activeTextEditorContainer == null)
            {
                return;
            }

            AnnotationCanvas.Children.Remove(_activeTextEditorContainer);
            _activeTextEditor = null;
            _activeTextEditorContainer = null;
        }

        private void RemovePreviewAnnotation()
        {
            if (_previewShape != null)
            {
                AnnotationCanvas.Children.Remove(_previewShape);
                _previewShape = null;
            }

            if (_previewArrow != null)
            {
                AnnotationCanvas.Children.Remove(_previewArrow);
                _previewArrow = null;
            }
        }

        private void ClearAnnotations()
        {
            CancelTextEditing();
            AnnotationCanvas.Children.Clear();
            _annotationHistory.Clear();
            _previewShape = null;
            _previewArrow = null;
            _drawingAnnotation = false;
            ClearToolSelection();
        }

        private void UndoLastAnnotation()
        {
            if (_annotationHistory.Count == 0)
            {
                return;
            }

            var lastIndex = _annotationHistory.Count - 1;
            var element = _annotationHistory[lastIndex];
            _annotationHistory.RemoveAt(lastIndex);
            AnnotationCanvas.Children.Remove(element);
        }

        private bool IsPointInsideActiveTextEditor(Point rootPoint)
        {
            if (_activeTextEditorContainer == null)
            {
                return false;
            }

            var topLeft = _activeTextEditorContainer.TranslatePoint(new Point(0, 0), RootGrid);
            var rect = new Rect(topLeft.X, topLeft.Y, _activeTextEditorContainer.ActualWidth, _activeTextEditorContainer.ActualHeight);
            return rect.Contains(rootPoint);
        }

        private void UpdateCursorVisual()
        {
            Cursor cursor;
            switch (_currentTool)
            {
                case AnnotationTool.None:
                    // 内容超出视口时左键拖动的是图片，用滚动光标区分于拖窗口。
                    cursor = HasOverflowingContent() ? Cursors.ScrollAll : Cursors.SizeAll;
                    break;
                case AnnotationTool.Text:
                    cursor = Cursors.IBeam;
                    break;
                default:
                    cursor = Cursors.Cross;
                    break;
            }

            // 标注模式下鼠标实际落在 AnnotationCanvas 上，两层都要设置光标。
            ImageCanvas.Cursor = cursor;
            AnnotationCanvas.Cursor = cursor;
        }

        /// <summary>释放本窗口持有的位图。</summary>
        private void ReleaseImage()
        {
            PreviewImage.Source = null;

            if (_bitmap != null && _ownsBitmap)
            {
                _bitmap.Dispose();
            }

            _bitmap = null;
            _ownsBitmap = false;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_interactionEndTimer != null)
            {
                _interactionEndTimer.Stop();
                _interactionEndTimer.Tick -= InteractionEndTimer_Tick;
                _interactionEndTimer = null;
            }

            if (_toolbarHideTimer != null)
            {
                _toolbarHideTimer.Stop();
                _toolbarHideTimer.Tick -= ToolbarHideTimer_Tick;
                _toolbarHideTimer = null;
            }

            ToolbarPopup.IsOpen = false;
            BackgroundColorPopup.IsOpen = false;
            ReleaseImage();

            base.OnClosed(e);
        }

        private enum AnnotationTool
        {
            None,
            Line,
            Rectangle,
            Ellipse,
            Arrow,
            Text
        }

        /// <summary>
        /// 把贴图窗口句柄包装成 IWin32Window，使颜色对话框以贴图窗口为 owner，
        /// 避免被 Topmost 的贴图窗口遮挡。
        /// </summary>
        private sealed class NativeWindowOwner : System.Windows.Forms.IWin32Window
        {
            private readonly IntPtr _handle;

            public NativeWindowOwner(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle
            {
                get { return _handle; }
            }
        }
    }
}
