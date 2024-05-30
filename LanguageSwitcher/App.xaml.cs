using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using Application = System.Windows.Application;
using Clipboard = System.Windows.Clipboard;

namespace LanguageSwitcher
{
    public partial class App : Application
    {
        private readonly NotifyIconManager notifyIconManager;

        private const int WM_HOTKEY = 0x0312;
        private const int VK_F4 = 0x73;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern IntPtr CloseClipboard();

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const int INPUT_KEYBOARD = 1;
        private const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const int KEYEVENTF_KEYUP = 0x0002;

        public App()
        {
            notifyIconManager = new NotifyIconManager();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            MainWindow = new MainWindow();
            notifyIconManager.Initialize(OnExitMenuItemClicked);

            IntPtr windowHandle = new WindowInteropHelper(MainWindow).EnsureHandle();
            HwndSource source = HwndSource.FromHwnd(windowHandle);
            source.AddHook(HwndHook);
            RegisterHotKey(windowHandle, 1, 0, VK_F4); // Register F4

            base.OnStartup(e);
        }

        private IntPtr HwndHook(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int id = wParam.ToInt32();
                if (id == 1)
                {
                    OnHotKeyPressed();
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }

        private void OnHotKeyPressed()
        {
            // Simulate changing the selected text to Ukrainian
            SendCtrlC(); // Copy the selected text to clipboard

            // Wait a moment for the clipboard operation to complete
            Thread.Sleep(100);

            // Get the copied text from clipboard
            string copiedText = GetClipboardText();

            // Convert the copied text to Ukrainian
            string convertedText = ConvertToUkrainian(copiedText);

            // Set the clipboard to the converted text
            SetClipboardText(convertedText);

            // Paste the converted text
            SendCtrlV();

            // Force garbage collection
            GC.Collect();
        }

        private static void SendCtrlC()
        {
            INPUT[] inputs = new INPUT[]
            {
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x11, // VK_CONTROL
                            wScan = 0,
                            dwFlags = 0,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x43, // VK_C
                            wScan = 0,
                            dwFlags = 0,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x43, // VK_C
                            wScan = 0,
                            dwFlags = KEYEVENTF_KEYUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x11, // VK_CONTROL
                            wScan = 0,
                            dwFlags = KEYEVENTF_KEYUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                }
            };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        private static void SendCtrlV()
        {
            INPUT[] inputs = new INPUT[]
            {
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x11, // VK_CONTROL
                            wScan = 0,
                            dwFlags = 0,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x56, // VK_V
                            wScan = 0,
                            dwFlags = 0,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x56, // VK_V
                            wScan = 0,
                            dwFlags = KEYEVENTF_KEYUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x11, // VK_CONTROL
                            wScan = 0,
                            dwFlags = KEYEVENTF_KEYUP,
                            time = 0,
                            dwExtraInfo = IntPtr.Zero
                        }
                    }
                }
            };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetOpenClipboardWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        private static Process GetProcessLockingClipboard()
        {
            int processId;
            GetWindowThreadProcessId(GetOpenClipboardWindow(), out processId);

            return Process.GetProcessById(processId);
        }

        private bool IsClipboardBusy()
        {
            return GetProcessLockingClipboard().ProcessName != "Idle";
        }

        private string GetClipboardText()
        {
            string clipboardText = string.Empty;

            int retries = 10;
            while (retries > 0)
            {
                try
                {

                    if (IsClipboardBusy())
                    {
                        CloseClipboard();
                    }
                    clipboardText = Clipboard.GetText();
                    CloseClipboard(); // Ensure clipboard is closed
                    break;
                }
                catch (COMException)
                {
                    retries--;
                    Thread.Sleep(100);
                }
            }
            return clipboardText;
        }

        private void SetClipboardText(string text)
        {
            int retries = 10;
            while (retries > 0)
            {
                try
                {
                    Clipboard.Clear(); // Clear the clipboard before setting text
                    Clipboard.SetText(text);
                    CloseClipboard(); // Ensure clipboard is closed
                    break;
                }
                catch (COMException)
                {
                    retries--;
                    Thread.Sleep(100);
                }
            }
        }

        private string ConvertToUkrainian(string text)
        {
            var keyMapping = new Dictionary<char, char>
            {
                { 'a', 'ф' }, { 'b', 'и' }, { 'c', 'с' }, { 'd', 'в' }, { 'e', 'у' },
                { 'f', 'а' }, { 'g', 'п' }, { 'h', 'р' }, { 'i', 'ш' }, { 'j', 'о' },
                { 'k', 'л' }, { 'l', 'д' }, { 'm', 'ь' }, { 'n', 'т' }, { 'o', 'щ' },
                { 'p', 'з' }, { 'q', 'й' }, { 'r', 'к' }, { 's', 'і' }, { 't', 'е' },
                { 'u', 'г' }, { 'v', 'м' }, { 'w', 'ц' }, { 'x', 'ч' }, { 'y', 'н' },
                { 'z', 'я' }, { ';', 'ж' }, { '\'', 'є' }, { ',', 'б' }, { '.', 'ю' },
                { '[', 'х' }, { ']', 'ї' }, { '-', '-' }, { '`', 'ґ' }, { ' ', ' ' }, // Keep space as is
                { 'A', 'Ф' }, { 'B', 'И' }, { 'C', 'С' }, { 'D', 'В' }, { 'E', 'У' },
                { 'F', 'А' }, { 'G', 'П' }, { 'H', 'Р' }, { 'I', 'Ш' }, { 'J', 'О' },
                { 'K', 'Л' }, { 'L', 'Д' }, { 'M', 'Ь' }, { 'N', 'Т' }, { 'O', 'Щ' },
                { 'P', 'З' }, { 'Q', 'Й' }, { 'R', 'К' }, { 'S', 'І' }, { 'T', 'Е' },
                { 'U', 'Г' }, { 'V', 'М' }, { 'W', 'Ц' }, { 'X', 'Ч' }, { 'Y', 'Н' },
                { 'Z', 'Я' }, { ':', 'Ж' }, { '"', 'Є' }, { '<', 'Б' }, { '>', 'Ю' },
                { '{', 'Х' }, { '}', 'Ї' }, { '_', '_' }, { '~', 'Ґ' },
                { '!', '!' }, { '@', '"' }, { '#', '№' }, { '$', ';' }, { '%', '%' },
                { '^', ':' }, { '&', '?' }, { '*', '*' }, { '(', '(' }, { ')', ')' },
                { '+', '+' }, { '=', '=' }, { '\\', '/' }, { '|', '\\' },
                { '?', ',' }, { '/', '.' }
            };

            var result = new char[text.Length];
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                result[i] = keyMapping.ContainsKey(ch) ? keyMapping[ch] : ch;
            }
            return new string(result);
        }

        private void OnExitMenuItemClicked(object sender, EventArgs e)
        {
            // Handle exit logic
            Application.Current.Shutdown(); // Shutdown the application
        }

        protected override void OnExit(ExitEventArgs e)
        {
            IntPtr windowHandle = new WindowInteropHelper(MainWindow).EnsureHandle();
            UnregisterHotKey(windowHandle, 1);
            notifyIconManager.Dispose();
            base.OnExit(e);
        }
    }
}
