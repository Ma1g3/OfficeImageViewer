using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace OfficeImageViewer
{
    public static class PreviewPlacement
    {
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct MonitorInfo
        { public int Size; public NativeRect Monitor, Work; public uint Flags; }
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr window, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);

        public static Rect FitBounds(Rect window, Rect work)
        {
            double width = Math.Min(window.Width, work.Width), height = Math.Min(window.Height, work.Height);
            return new Rect(Math.Max(work.Left, Math.Min(window.Left, work.Right - width)),
                Math.Max(work.Top, Math.Min(window.Top, work.Bottom - height)), width, height);
        }

        internal static void KeepVisible(Window window, IntPtr owner)
        {
            IntPtr handle = new WindowInteropHelper(window).Handle;
            MonitorInfo info = new MonitorInfo { Size = Marshal.SizeOf(typeof(MonitorInfo)) };
            NativeRect bounds;
            if (handle == IntPtr.Zero || !GetMonitorInfo(MonitorFromWindow(owner == IntPtr.Zero ? handle : owner, 2), ref info) ||
                !GetWindowRect(handle, out bounds) || info.Work.Right <= info.Work.Left || info.Work.Bottom <= info.Work.Top) return;
            Rect fitted = FitBounds(new Rect(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top),
                new Rect(info.Work.Left, info.Work.Top, info.Work.Right - info.Work.Left, info.Work.Bottom - info.Work.Top));
            PresentationSource source = PresentationSource.FromVisual(window);
            Matrix scale = source == null || source.CompositionTarget == null ? Matrix.Identity : source.CompositionTarget.TransformToDevice;
            window.MinWidth = Math.Min(window.MinWidth, fitted.Width / Math.Max(1, scale.M11));
            window.MinHeight = Math.Min(window.MinHeight, fitted.Height / Math.Max(1, scale.M22));
            SetWindowPos(handle, IntPtr.Zero, (int)fitted.X, (int)fitted.Y, (int)fitted.Width, (int)fitted.Height, 0x0014);
        }
    }
}
