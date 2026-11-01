using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace OfficeImageViewer
{
    // Coordinates are DIPs. Zoom=1 maps one source pixel to one screen pixel.
    public sealed class ImageViewport
    {
        public double SourceWidth { get; private set; }
        public double SourceHeight { get; private set; }
        public double ViewWidth { get; private set; }
        public double ViewHeight { get; private set; }
        public double Zoom { get; private set; }
        public double X { get; private set; }
        public double Y { get; private set; }
        public bool IsFit { get; private set; }
        public double FitZoom { get { return Math.Min(1, Math.Min(ViewWidth / SourceWidth, ViewHeight / SourceHeight)); } }

        public ImageViewport() { SourceWidth = SourceHeight = ViewWidth = ViewHeight = Zoom = 1; IsFit = true; }
        public void SetImage(double width, double height)
        {
            ValidatePositive(width); ValidatePositive(height);
            SourceWidth = width; SourceHeight = height;
            Fit();
        }
        public void Resize(double width, double height)
        {
            ValidatePositive(width); ValidatePositive(height);
            double imageX = (ViewWidth / 2 - X) / Zoom;
            double imageY = (ViewHeight / 2 - Y) / Zoom;
            ViewWidth = width; ViewHeight = height;
            if (IsFit) { Fit(); return; }
            X = ViewWidth / 2 - imageX * Zoom;
            Y = ViewHeight / 2 - imageY * Zoom;
            Constrain();
        }
        public void Fit()
        {
            IsFit = true; Zoom = FitZoom;
            X = (ViewWidth - SourceWidth * Zoom) / 2;
            Y = (ViewHeight - SourceHeight * Zoom) / 2;
        }
        public void ZoomAt(double requestedZoom, double anchorX, double anchorY)
        {
            ValidatePositive(requestedZoom);
            // Fit may be below the manual lower bound. A zoom-out request
            // must never jump upward to that bound or leave fit mode.
            if (Zoom < 0.1 && requestedZoom <= Zoom) return;
            double next = Math.Max(0.1, Math.Min(16, requestedZoom));
            double imageX = (anchorX - X) / Zoom;
            double imageY = (anchorY - Y) / Zoom;
            Zoom = next; IsFit = false;
            X = anchorX - imageX * Zoom; Y = anchorY - imageY * Zoom;
            Constrain();
        }
        public void Pan(double dx, double dy)
        {
            if (Double.IsNaN(dx) || Double.IsInfinity(dx) || Double.IsNaN(dy) || Double.IsInfinity(dy)) return;
            X += dx; Y += dy; Constrain();
        }
        private void Constrain()
        {
            double width = SourceWidth * Zoom, height = SourceHeight * Zoom;
            X = width <= ViewWidth ? (ViewWidth - width) / 2 : Math.Max(ViewWidth - width, Math.Min(0, X));
            Y = height <= ViewHeight ? (ViewHeight - height) / 2 : Math.Max(ViewHeight - height, Math.Min(0, Y));
        }
        private static void ValidatePositive(double value)
        {
            if (value <= 0 || Double.IsNaN(value) || Double.IsInfinity(value)) throw new ArgumentOutOfRangeException("value");
        }
    }

    public sealed class PreviewWindowService : IImageViewer, IDisposable
    {
        private PreviewWindow window;
        private bool disposed;
        public void ShowImage(ImagePayload image, IntPtr owner)
        {
            if (disposed) throw new ObjectDisposedException("PreviewWindowService");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("图片查看器需要在 Office 窗口线程中打开。");
            BitmapSource bitmap = DecodeImage(image);
            bool created = window == null;
            if (created)
            {
                PreviewWindow createdWindow = new PreviewWindow();
                createdWindow.WindowStartupLocation = owner == IntPtr.Zero ? WindowStartupLocation.CenterScreen : WindowStartupLocation.CenterOwner;

                createdWindow.Closed += delegate { if (Object.ReferenceEquals(window, createdWindow)) window = null; };
                window = createdWindow;
            }
            try
            {
                // Reused preview follows the current document window, not the first one.
                WindowInteropHelper interop = new WindowInteropHelper(window);
                bool ownerChanged = interop.Owner != owner;
                interop.Owner = owner;
                window.SetImage(bitmap, image.Title, image.Detail);
                if (!window.IsVisible) window.Show();
                if (window.WindowState == WindowState.Minimized) window.WindowState = WindowState.Normal;
                if (created || ownerChanged) PreviewPlacement.KeepVisible(window, owner);
                window.Activate();
                window.FocusCanvas();
            }
            catch
            {
                if (created && window != null) { window.Close(); window = null; }
                throw;
            }
        }

        public static BitmapSource DecodeImage(ImagePayload image)
        {
            if (image == null || image.Bytes == null || image.Bytes.Length == 0)
                throw new ImageReadException("没有可查看的图片，请先选择一张图片。");
            if (image.Bytes.Length > 64 * 1024 * 1024)
                throw new ImageReadException("图片文件超过 64 MB，无法在查看器中安全打开。");
            try
            {
                using (MemoryStream stream = new MemoryStream(image.Bytes, false))
                {
                    // OnDemand + DelayCreation reads metadata first. Limit expanded pixels before OnLoad decode.
                    BitmapDecoder header = BitmapDecoder.Create(stream, BitmapCreateOptions.DelayCreation,
                                                                 BitmapCacheOption.OnDemand);
                    BitmapFrame frame = header.Frames[0];
                    if (frame.PixelWidth < 1 || frame.PixelHeight < 1 ||
                        frame.PixelWidth > 65535 || frame.PixelHeight > 65535 ||
                        (long)frame.PixelWidth * frame.PixelHeight > 32000000)
                        throw new ImageReadException("图片超过 3200 万像素或边长过大，无法在查看器中安全打开。");
                    stream.Position = 0;
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.CreateOptions = BitmapCreateOptions.PreservePixelFormat;
                    bitmap.StreamSource = stream;
                    bitmap.EndInit();
                    bitmap.Freeze();
                    return bitmap;
                }
            }
            catch (ImageReadException) { throw; }
            catch (OutOfMemoryException)
            {
                throw new ImageReadException("内存不足，无法打开这张图片。请关闭其他大文档后重试。");
            }
            catch (Exception ex)
            {
                if (ex is NotSupportedException || ex is FileFormatException ||
                    ex is ArgumentException || ex is IOException || ex is System.Runtime.InteropServices.COMException)
                    throw new ImageReadException("图片格式不受支持或内容已损坏。首版支持 PNG、JPEG、BMP、GIF 和 TIFF。", ex);
                throw;
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (window != null)
            {
                if (window.Dispatcher.CheckAccess()) window.Close();
                else window.Dispatcher.Invoke(new Action(window.Close));
                window = null;
            }
        }
    }

    public sealed class PreviewWindow : Window
    {
        private readonly Canvas canvas;
        private readonly Image image;
        private readonly TextBlock zoomLabel, detailLabel;
        private readonly ImageViewport viewport = new ImageViewport();
        private BitmapSource bitmap;
        private Point lastDrag;
        private bool dragging;
        private double dpiX = 1, dpiY = 1;
        private HwndSource source;
        public ImageViewport Viewport { get { return viewport; } }
        public Canvas ImageCanvas { get { return canvas; } }

        public PreviewWindow()
        {
            Title = "图片查看器";
            Width = Math.Min(1000, Math.Max(520, SystemParameters.WorkArea.Width - 40));
            Height = Math.Min(740, Math.Max(360, SystemParameters.WorkArea.Height - 40));
            MinWidth = Math.Min(520, SystemParameters.WorkArea.Width);
            MinHeight = Math.Min(360, SystemParameters.WorkArea.Height);
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = Brush("#181A1D"); Foreground = Brush("#F3F4F6");
            FontFamily = new FontFamily("Segoe UI, Microsoft YaHei UI"); FontSize = 14;
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;
            Border toolbar = new Border { Background = Brush("#25282D"), Padding = new Thickness(14, 10, 14, 10) };
            StackPanel controls = new StackPanel { Orientation = Orientation.Horizontal };
            toolbar.Child = controls; root.Children.Add(toolbar);
            controls.Children.Add(MakeButton("适应窗口", "适应窗口 (Home)", delegate { FitImage(); }));
            controls.Children.Add(MakeButton("100%", "原始像素 (1)", delegate { SetZoom(1); }));
            Border separator = new Border { Width = 1, Background = Brush("#434850"), Margin = new Thickness(9, 4, 13, 4) };
            controls.Children.Add(separator);
            controls.Children.Add(MakeButton("缩小", "缩小 (-)", delegate { SetZoom(viewport.Zoom / 1.25); }));
            zoomLabel = new TextBlock { Width = 66, TextAlignment = TextAlignment.Center, VerticalAlignment = VerticalAlignment.Center, FontSize = 14 };
            AutomationProperties.SetName(zoomLabel, "当前缩放比例");
            controls.Children.Add(zoomLabel);
            controls.Children.Add(MakeButton("放大", "放大 (+)", delegate { SetZoom(viewport.Zoom * 1.25); }));

            canvas = new Canvas { Background = Brush("#181A1D"), ClipToBounds = true, Focusable = true, Cursor = Cursors.Hand };
            AutomationProperties.SetName(canvas, "图片画布，滚轮缩放，拖动平移");
            Grid.SetRow(canvas, 1); root.Children.Add(canvas);
            image = new Image { Stretch = Stretch.Fill, IsHitTestVisible = false, SnapsToDevicePixels = true };
            RenderOptions.SetBitmapScalingMode(image, BitmapScalingMode.HighQuality);
            canvas.Children.Add(image);

            Border footer = new Border { Background = Brush("#25282D"), Padding = new Thickness(16, 9, 16, 9) };
            StackPanel footerText = new StackPanel();
            detailLabel = new TextBlock { FontSize = 12, Foreground = Brush("#D5D9DF"), TextTrimming = TextTrimming.CharacterEllipsis };
            footerText.Children.Add(detailLabel);
            footerText.Children.Add(new TextBlock { Text = "滚轮缩放  ·  拖动平移  ·  Home 适应窗口  ·  1 原始像素  ·  Esc 关闭",
                Margin = new Thickness(0, 4, 0, 0), FontSize = 12, Foreground = Brush("#B6BDC8"), TextWrapping = TextWrapping.Wrap });
            footer.Child = footerText; Grid.SetRow(footer, 2); root.Children.Add(footer);

            canvas.SizeChanged += delegate
            {
                if (canvas.ActualWidth > 0 && canvas.ActualHeight > 0)
                { viewport.Resize(canvas.ActualWidth, canvas.ActualHeight); ApplyViewport(); }
            };
            canvas.MouseWheel += delegate(object sender, MouseWheelEventArgs e)
            {
                Point p = e.GetPosition(canvas);
                viewport.ZoomAt(viewport.Zoom * Math.Pow(1.25, e.Delta / 120.0), p.X, p.Y);
                ApplyViewport(); e.Handled = true;
            };
            canvas.MouseLeftButtonDown += delegate(object sender, MouseButtonEventArgs e)
            {
                canvas.Focus(); lastDrag = e.GetPosition(canvas); dragging = canvas.CaptureMouse();
                if (dragging) canvas.Cursor = Cursors.SizeAll;
                e.Handled = true;
            };
            canvas.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                if (!dragging) return;
                Point p = e.GetPosition(canvas);
                viewport.Pan(p.X - lastDrag.X, p.Y - lastDrag.Y); lastDrag = p; ApplyViewport();
            };
            canvas.MouseLeftButtonUp += delegate { StopDragging(); };
            canvas.LostMouseCapture += delegate { dragging = false; canvas.Cursor = Cursors.Hand; };
            PreviewKeyDown += HandleKey;
            SourceInitialized += delegate
            {
                source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
                if (source != null) source.AddHook(WindowMessage);
                UpdateDpi();
            };
            Closed += delegate
            {
                StopDragging();
                if (source != null) { source.RemoveHook(WindowMessage); source = null; }
                image.Source = null; bitmap = null;
            };
        }

        public void SetImage(BitmapSource value, string title, string detail)
        {
            if (value == null) throw new ArgumentNullException("value");
            bitmap = value; image.Source = value;
            Title = (String.IsNullOrWhiteSpace(title) ? "图片" : title) + " — 图片查看器";
            detailLabel.Text = value.PixelWidth + " × " + value.PixelHeight + " 像素  ·  " +
                (String.IsNullOrWhiteSpace(detail) ? "原图，可能包含文档中被裁剪的区域" : detail);
            detailLabel.ToolTip = detailLabel.Text;
            viewport.SetImage(value.PixelWidth / dpiX, value.PixelHeight / dpiY);
            ApplyViewport();
        }
        public void FocusCanvas() { canvas.Focus(); }
        public void FitImage() { viewport.Fit(); ApplyViewport(); }
        public void SetZoom(double value)
        {
            viewport.ZoomAt(value, viewport.ViewWidth / 2, viewport.ViewHeight / 2); ApplyViewport();
        }
        public void PanImage(double dx, double dy) { viewport.Pan(dx, dy); ApplyViewport(); }
        public void ZoomAt(double value, Point anchor) { viewport.ZoomAt(value, anchor.X, anchor.Y); ApplyViewport(); }

        private void ApplyViewport()
        {
            image.Width = viewport.SourceWidth * viewport.Zoom;
            image.Height = viewport.SourceHeight * viewport.Zoom;
            Canvas.SetLeft(image, viewport.X); Canvas.SetTop(image, viewport.Y);
            zoomLabel.Text = (viewport.Zoom * 100).ToString(viewport.Zoom < 0.1 ? "0.0" : "0") + "%";
        }
        private void UpdateDpi()
        {
            PresentationSource presentation = PresentationSource.FromVisual(this);
            if (presentation == null || presentation.CompositionTarget == null) return;
            Matrix transform = presentation.CompositionTarget.TransformToDevice;
            double x = transform.M11, y = transform.M22;
            if (x <= 0 || y <= 0) return;
            if (x != dpiX || y != dpiY)
            {
                dpiX = x; dpiY = y;
                if (bitmap != null) viewport.SetImage(bitmap.PixelWidth / dpiX, bitmap.PixelHeight / dpiY);
                ApplyViewport();
            }
        }
        private IntPtr WindowMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == 0x02E0) Dispatcher.BeginInvoke(new Action(UpdateDpi));
            return IntPtr.Zero;
        }
        private void StopDragging()
        {
            dragging = false;
            if (canvas.IsMouseCaptured) canvas.ReleaseMouseCapture();
            canvas.Cursor = Cursors.Hand;
        }
        private void HandleKey(object sender, KeyEventArgs e)
        {
            Key key = e.Key == Key.System ? e.SystemKey : e.Key;
            bool handled = true;
            switch (key)
            {
                case Key.Escape: Close(); break;
                case Key.Home: case Key.D0: case Key.NumPad0: FitImage(); break;
                case Key.D1: case Key.NumPad1: SetZoom(1); break;
                case Key.Add: case Key.OemPlus: SetZoom(viewport.Zoom * 1.25); break;
                case Key.Subtract: case Key.OemMinus: SetZoom(viewport.Zoom / 1.25); break;
                case Key.Left: PanImage(40, 0); break;
                case Key.Right: PanImage(-40, 0); break;
                case Key.Up: PanImage(0, 40); break;
                case Key.Down: PanImage(0, -40); break;
                default: handled = false; break;
            }
            e.Handled = handled;
        }
        private Button MakeButton(string label, string name, Action action)
        {
            Button button = new Button { Content = label, MinWidth = 64, Height = 32, Margin = new Thickness(0, 0, 6, 0),
                Padding = new Thickness(10, 0, 10, 0), Background = Brush("#353A42"), Foreground = Brush("#F3F4F6"),
                BorderBrush = Brush("#535B67"), BorderThickness = new Thickness(1), Cursor = Cursors.Hand, ToolTip = name };
            // Explicit template retains readable dark text in hover/pressed/focused states.
            FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            border.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            border.SetBinding(Border.BorderBrushProperty, new System.Windows.Data.Binding("BorderBrush") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            border.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));
            FrameworkElementFactory content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            border.AppendChild(content);
            button.Template = new ControlTemplate(typeof(Button)) { VisualTree = border };
            Style style = new Style(typeof(Button));
            Trigger hover = new Trigger { Property = IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#46505D")));
            Trigger pressed = new Trigger { Property = Button.IsPressedProperty, Value = true };
            pressed.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#1F4778")));
            Trigger focused = new Trigger { Property = IsKeyboardFocusedProperty, Value = true };
            focused.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#7CB4FF")));
            style.Triggers.Add(hover); style.Triggers.Add(pressed); style.Triggers.Add(focused);

            // Local background has precedence over triggers: move styling values into setters.
            button.ClearValue(Control.BackgroundProperty); button.ClearValue(Control.BorderBrushProperty);
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brush("#353A42")));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Brush("#535B67")));
            button.Style = style;
            AutomationProperties.SetName(button, name);
            button.Click += delegate { action(); canvas.Focus(); };
            return button;
        }
        private static SolidColorBrush Brush(string hex)
        {
            SolidColorBrush value = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            value.Freeze(); return value;
        }
    }
}




