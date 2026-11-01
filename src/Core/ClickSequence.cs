using System;

namespace OfficeImageViewer
{
    public sealed class ClickSequence
    {
        private IntPtr window;
        private int x, y;
        private uint time;
        public bool Pending { get; private set; }
        public void Reset() { Pending = false; window = IntPtr.Zero; }
        public bool Register(IntPtr target, int px, int py, uint tick, uint interval, int width, int height)
        {
            if (Pending && target == window && unchecked(tick - time) <= interval &&
                Math.Abs((long)px - x) <= Math.Max(1, width / 2) && Math.Abs((long)py - y) <= Math.Max(1, height / 2))
            { Reset(); return true; }
            Pending = true; window = target; x = px; y = py; time = tick;
            return false;
        }
    }
}
