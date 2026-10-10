using System;
using System.Windows.Forms;
using RevitBallet.Core.Remote;

namespace RevitBalletRemote
{
    /// <summary>A message-only window that owns one global hotkey and tells when it is pressed. Re-resolves the key when the keyboard layout changes
    /// (the backtick is VK_OEM_3 on a US layout and VK_OEM_8 on a UK one).</summary>
    internal sealed class HotkeyWindow : NativeWindow, IDisposable
    {
        private const int HotkeyId = 0x5242;                    // "RB"
        private HotkeySpec spec;
        private uint registeredVk;
        private IntPtr layout;
        private readonly Timer layoutWatch = new Timer { Interval = 3000 };

        public event Action Pressed;
        public string Error { get; private set; }
        public bool Registered { get { return registeredVk != 0; } }
        public HotkeySpec Spec { get { return spec; } }

        public HotkeyWindow()
        {
            CreateHandle(new CreateParams { Parent = new IntPtr(-3) });      // HWND_MESSAGE
            layoutWatch.Tick += (s, e) => CheckLayout();
            layoutWatch.Start();
        }

        public bool Register(HotkeySpec hotkey)
        {
            spec = hotkey;
            return Reregister();
        }

        private bool Reregister()
        {
            if (registeredVk != 0) { Native.UnregisterHotKey(Handle, HotkeyId); registeredVk = 0; }
            layout = CurrentLayout();
            uint vk = ResolveVk(spec);
            if (vk == 0) { Error = "unknown key " + spec.Key; return false; }
            if (!Native.RegisterHotKey(Handle, HotkeyId, spec.Modifiers | Native.MOD_NOREPEAT, vk))
            {
                Error = "another program already uses " + spec + " (or Windows reserves it)";
                return false;
            }
            registeredVk = vk; Error = null;
            return true;
        }

        private static IntPtr CurrentLayout()
        {
            uint pid; uint thread = Native.GetWindowThreadProcessId(Native.GetForegroundWindow(), out pid);
            return Native.GetKeyboardLayout(thread);
        }

        private void CheckLayout()
        {
            if (spec == null || spec.Key.Length != 1) return;        // named keys do not depend on the layout
            var now = CurrentLayout();
            if (now != layout && now != IntPtr.Zero) Reregister();
        }

        public static uint ResolveVk(HotkeySpec h)
        {
            string k = h.Key;
            if (k.Length == 1)
            {
                short r = Native.VkKeyScan(k[0]);
                if (r != -1) return (uint)(r & 0xFF);
                return k == "`" ? 0xC0u : 0u;
            }
            if (k.Length >= 2 && k[0] == 'F') { int n; if (int.TryParse(k.Substring(1), out n)) return (uint)(0x70 + n - 1); }
            switch (k)
            {
                case "Space": return 0x20; case "Tab": return 0x09; case "Enter": return 0x0D; case "Esc": return 0x1B; case "Insert": return 0x2D; case "Delete": return 0x2E;
                case "Home": return 0x24; case "End": return 0x23; case "Pageup": return 0x21; case "Pagedown": return 0x22;
                case "Left": return 0x25; case "Up": return 0x26; case "Right": return 0x27; case "Down": return 0x28; case "Grave": return 0xC0;
            }
            return 0;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_HOTKEY && m.WParam.ToInt32() == HotkeyId) { var h = Pressed; if (h != null) h(); return; }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            layoutWatch.Stop(); layoutWatch.Dispose();
            if (registeredVk != 0) Native.UnregisterHotKey(Handle, HotkeyId);
            DestroyHandle();
        }
    }
}
