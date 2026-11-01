using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeImageViewer
{
    internal static class NativeOfficeWindow
    {
        private delegate bool EnumerateWindow(IntPtr hwnd, IntPtr data);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumerateWindow callback, IntPtr data);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        public static IntPtr FindPowerPointFrame(string caption)
        {
            uint process = (uint)Process.GetCurrentProcess().Id;
            IntPtr foreground = GetAncestor(GetForegroundWindow(), 3);
            IntPtr exact = IntPtr.Zero, active = IntPtr.Zero, only = IntPtr.Zero;
            int count = 0;
            EnumWindows(delegate(IntPtr hwnd, IntPtr ignored)
            {
                uint ownerProcess;
                GetWindowThreadProcessId(hwnd, out ownerProcess);
                if (ownerProcess != process || !IsWindowVisible(hwnd)) return true;
                StringBuilder kind = new StringBuilder(256);
                GetClassName(hwnd, kind, kind.Capacity);
                if (!kind.ToString().StartsWith("PPTFrameClass", StringComparison.Ordinal)) return true;
                ++count; only = hwnd;
                if (hwnd == foreground) active = hwnd;
                StringBuilder title = new StringBuilder(1024);
                GetWindowText(hwnd, title, title.Capacity);
                if (!String.IsNullOrEmpty(caption) && title.ToString().IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0)
                    exact = hwnd;
                return true;
            }, IntPtr.Zero);
            return active != IntPtr.Zero ? active : exact != IntPtr.Zero ? exact : count == 1 ? only : IntPtr.Zero;
        }
    }
}
