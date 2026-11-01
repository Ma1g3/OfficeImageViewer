using System;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeImageViewer
{
    // The hook belongs only to this add-in's own Office UI thread. Other
    // applications and keyboard messages are neither inspected nor recorded.
    internal sealed class PowerPointClickBridge : IDisposable
    {
        private delegate IntPtr MessageHook(int code, IntPtr remove, IntPtr message);
        [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct Message
        { public IntPtr Hwnd; public uint Id; public IntPtr WParam, LParam; public uint Time; public Point Position; public uint Private; }
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int kind, MessageHook hook, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr remove, IntPtr message);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern uint GetDoubleClickTime();
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int length);
        private readonly ClickSequence clicks = new ClickSequence();
        private readonly Func<long> selection;
        private readonly Func<bool> preview;
        private readonly MessageHook callback;
        private IntPtr hook;
        private long selectedKey;
        public bool Installed { get { return hook != IntPtr.Zero; } }
        public int InstallError { get; private set; }
        public int SurfaceMouseDowns { get; private set; }

        public PowerPointClickBridge(Func<long> getSelection, Func<bool> showPreview)
        {
            selection = getSelection; preview = showPreview; callback = Filter;
            hook = SetWindowsHookEx(3, callback, IntPtr.Zero, GetCurrentThreadId());
            if (hook == IntPtr.Zero) InstallError = Marshal.GetLastWin32Error();
        }
        private static bool IsDocumentSurface(IntPtr hwnd)
        {
            for (int i = 0; hwnd != IntPtr.Zero && i < 12; ++i, hwnd = GetParent(hwnd))
            {
                StringBuilder name = new StringBuilder(256); GetClassName(hwnd, name, name.Capacity);
                string value = name.ToString();
                if (value.Equals("mdiClass", StringComparison.OrdinalIgnoreCase) || value.StartsWith("paneClassDC", StringComparison.Ordinal)) return true;
                if (value.StartsWith("PPTFrameClass", StringComparison.Ordinal)) break;
            }
            return false;
        }
        private IntPtr Filter(int code, IntPtr remove, IntPtr pointer)
        {
            try
            {
                if (code >= 0 && remove == new IntPtr(1))
                {
                    Message msg = (Message)Marshal.PtrToStructure(pointer, typeof(Message));
                    if (msg.Id == 0x0201 || msg.Id == 0x0202 || msg.Id == 0x0203)
                    {
                        if (!IsDocumentSurface(msg.Hwnd)) { clicks.Reset(); selectedKey = 0; }
                        else if (msg.Id == 0x0202) { if (clicks.Pending) selectedKey = selection(); }
                        else
                        {
                            ++SurfaceMouseDowns;
                            bool doubled = msg.Id == 0x0203 || clicks.Register(msg.Hwnd, msg.Position.X, msg.Position.Y, msg.Time,
                                GetDoubleClickTime(), GetSystemMetrics(36), GetSystemMetrics(37));
                            if (!doubled) selectedKey = 0;
                            else
                            {
                                bool same = selectedKey != 0 && selectedKey == selection();
                                clicks.Reset(); selectedKey = 0;
                                if (same && preview())
                                {
                                    msg.Id = 0; msg.WParam = msg.LParam = IntPtr.Zero;
                                    Marshal.StructureToPtr(msg, pointer, false);
                                }
                            }
                        }
                    }
                }
            }
            catch { clicks.Reset(); selectedKey = 0; }
            return CallNextHookEx(hook, code, remove, pointer);
        }
        public void Dispose() { if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; } }
    }
}
