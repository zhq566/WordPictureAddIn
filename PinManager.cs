using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

namespace WordPictureAddIn
{
    /// <summary>
    /// 贴图窗口的运行时管理器：初始化 WPF 环境、预热备用窗口、维护窗口生命周期。
    ///
    /// 单独拆成独立类是性能设计的一部分：让 <see cref="ThisAddIn"/> 在 Word 启动阶段
    /// 完全不引用 WPF 类型（Application / Window / DispatcherTimer）。
    /// 加载 PresentationFramework、PresentationCore、WindowsBase 以及 JIT 两千多行
    /// 贴图窗口代码的开销，如果放在 Word 启动路径上会明显拖慢启动；改为用户第一次
    /// 真的要用贴图功能（选中图片）时才按需创建，功能与体验保持一致。
    /// </summary>
    internal sealed class PinManager
    {
        /// <summary>已贴出的贴图窗口，可同时存在多个。</summary>
        private readonly List<ImagePreviewWindow> _pinWindows = new List<ImagePreviewWindow>();

        /// <summary>预先创建好的空闲贴图窗口，下次双击时直接复用，省去构造开销。</summary>
        private ImagePreviewWindow _sparePinWindow;

        private bool _isShuttingDown;
        private bool _wpfInitialized;

        /// <summary>
        /// 把图片贴到屏幕上。位图所有权移交给新窗口，窗口关闭时自行释放。
        /// </summary>
        public void ShowPin(Image image)
        {
            if (image == null)
            {
                return;
            }

            if (!EnsureWpfApplication())
            {
                image.Dispose();
                return;
            }

            var window = TakeSparePinWindow();

            try
            {
                window.SetImage(image);
            }
            catch (Exception ex)
            {
                AddInLog.Error("ShowPin 设置图片失败", ex);

                try
                {
                    window.Close();
                }
                catch (Exception closeEx)
                {
                    AddInLog.Error("ShowPin 关闭窗口失败", closeEx);
                }

                image.Dispose();
                throw;
            }

            _pinWindows.Add(window);

            window.Show();
            window.Activate();

            // 空闲时再预热一个备用窗口，保证下一次双击依然能立刻弹出。
            SchedulePreviewWindowPrewarm();
        }

        /// <summary>
        /// 空闲时预热备用贴图窗口。首次调用会顺带初始化 WPF 环境，
        /// 因此这里是整个加载项唯一的 WPF 加载入口——由第一次选中图片触发。
        /// </summary>
        public void SchedulePreviewWindowPrewarm()
        {
            if (_isShuttingDown)
            {
                return;
            }

            if (!EnsureWpfApplication())
            {
                return;
            }

            var application = System.Windows.Application.Current;
            if (application == null)
            {
                return;
            }

            try
            {
                application.Dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.ApplicationIdle,
                    new Action(delegate
                    {
                        CreateSparePinWindowIfNeeded();
                    }));
            }
            catch (Exception ex)
            {
                AddInLog.Error("安排贴图窗口预热失败", ex);
            }
        }

        /// <summary>关闭所有贴图窗口（含备用窗口），加载项卸载时调用。</summary>
        public void CloseAll()
        {
            _isShuttingDown = true;

            while (_pinWindows.Count > 0)
            {
                var window = _pinWindows[_pinWindows.Count - 1];
                _pinWindows.RemoveAt(_pinWindows.Count - 1);

                try
                {
                    window.Closed -= PinWindow_Closed;
                    window.Close();
                }
                catch (Exception ex)
                {
                    AddInLog.Error("关闭贴图窗口失败", ex);
                }
            }

            var spare = _sparePinWindow;
            _sparePinWindow = null;

            if (spare != null)
            {
                try
                {
                    spare.Closed -= PinWindow_Closed;
                    spare.Close();
                }
                catch (Exception ex)
                {
                    AddInLog.Error("关闭备用贴图窗口失败", ex);
                }
            }
        }

        /// <summary>
        /// 初始化 WPF 环境。必须在 Word 主线程（STA）上调用。
        /// </summary>
        private bool EnsureWpfApplication()
        {
            if (_wpfInitialized)
            {
                return true;
            }

            try
            {
                if (System.Windows.Application.Current == null)
                {
                    new System.Windows.Application
                    {
                        ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                    };
                }

                _wpfInitialized = true;
                return true;
            }
            catch (Exception ex)
            {
                AddInLog.Error("初始化 WPF 环境失败", ex);
                return false;
            }
        }

        private void CreateSparePinWindowIfNeeded()
        {
            if (_isShuttingDown || _sparePinWindow != null)
            {
                return;
            }

            try
            {
                var stopwatch = Stopwatch.StartNew();
                _sparePinWindow = CreatePinWindow();
                stopwatch.Stop();

                AddInLog.Write("预热贴图窗口完成，耗时 {0:F1} ms。", stopwatch.Elapsed.TotalMilliseconds);
            }
            catch (Exception ex)
            {
                AddInLog.Error("预热贴图窗口失败", ex);
            }
        }

        private ImagePreviewWindow CreatePinWindow()
        {
            var window = new ImagePreviewWindow();
            window.Closed += PinWindow_Closed;
            return window;
        }

        private ImagePreviewWindow TakeSparePinWindow()
        {
            var window = _sparePinWindow;
            if (window != null)
            {
                _sparePinWindow = null;
                return window;
            }

            return CreatePinWindow();
        }

        private void PinWindow_Closed(object sender, EventArgs e)
        {
            var window = sender as ImagePreviewWindow;
            if (window == null)
            {
                return;
            }

            window.Closed -= PinWindow_Closed;
            _pinWindows.Remove(window);

            if (ReferenceEquals(_sparePinWindow, window))
            {
                _sparePinWindow = null;
            }

            if (_isShuttingDown)
            {
                return;
            }

            // 关闭后再次空闲时预热新的备用窗口。
            SchedulePreviewWindowPrewarm();
        }
    }
}
