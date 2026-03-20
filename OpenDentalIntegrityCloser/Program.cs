using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Reflection;
using System.IO;

namespace OpenDentalIntegrityCloser
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayAppContext());
        }
    }

    public class TrayAppContext : ApplicationContext
    {
        private readonly NotifyIcon _notifyIcon;
        private readonly Timer _timer;
        private readonly MenuItem _pauseMenuItem;
        private readonly MenuItem _startupMenuItem;

        private bool _paused = false;

        // Adjust these if you want
        private const int CheckIntervalMs = 1500;

        // Main window title must contain this
        private const string OpenDentalTitleMatch = "Open Dental";

        // Popup/dialog title must contain this
        private const string DialogTitleMatch = "Database Integrity";

        // Optional text match inside the dialog. Leave empty to close any "Database Integrity" dialog.
        private const string RequiredDialogTextMatch =
            "Open Dental has detected that this patient's data has been modified";

        public TrayAppContext()
        {
            ContextMenu menu = new ContextMenu();

            _pauseMenuItem = new MenuItem("Pause", OnPauseResumeClicked);
            MenuItem exitMenuItem = new MenuItem("Exit", OnExitClicked);
            _startupMenuItem = new MenuItem("Run at Startup", OnStartupClicked);
            _startupMenuItem.Checked = StartupManager.IsStartupEnabled();

            menu.MenuItems.Add(_pauseMenuItem);
            menu.MenuItems.Add(_startupMenuItem);
            menu.MenuItems.Add("-");
            menu.MenuItems.Add(exitMenuItem);                        

            _notifyIcon = new NotifyIcon
            {
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                ContextMenu = menu,
                Text = "OpenDental Integrity Closer",
                Visible = true
            };

            _timer = new Timer();
            _timer.Interval = CheckIntervalMs;
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void OnStartupClicked(object sender, EventArgs e)
        {
            bool enable = !_startupMenuItem.Checked;

            StartupManager.SetStartup(enable);
            _startupMenuItem.Checked = enable;
        }

        private void OnPauseResumeClicked(object sender, EventArgs e)
        {
            _paused = !_paused;
            _pauseMenuItem.Text = _paused ? "Resume" : "Pause";
            _notifyIcon.Text = _paused
                ? "OpenDental Integrity Closer (Paused)"
                : "OpenDental Integrity Closer";
        }

        private void OnExitClicked(object sender, EventArgs e)
        {
            _timer.Stop();
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            ExitThread();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_paused)
                return;

            try
            {
                if (!AnyTopLevelWindowContains(OpenDentalTitleMatch))
                    return;

                IntPtr dialogHwnd = FindTopLevelWindowContaining(DialogTitleMatch);
                if (dialogHwnd == IntPtr.Zero)
                    return;

                string dialogText = ReadDialogText(dialogHwnd);

                // Bonus: only close if the text matches
                if (!string.IsNullOrWhiteSpace(RequiredDialogTextMatch))
                {
                    if (string.IsNullOrWhiteSpace(dialogText))
                        return;

                    if (dialogText.IndexOf(RequiredDialogTextMatch, StringComparison.OrdinalIgnoreCase) < 0)
                        return;
                }

                CloseDialog(dialogHwnd);
            }
            catch
            {
                // Swallow exceptions so tray app keeps running
            }
        }

        private static bool AnyTopLevelWindowContains(string titlePart)
        {
            bool found = false;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                string title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title) &&
                    title.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = true;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static IntPtr FindTopLevelWindowContaining(string titlePart)
        {
            IntPtr foundHwnd = IntPtr.Zero;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                string title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title) &&
                    title.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundHwnd = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return foundHwnd;
        }

        private static string GetWindowTitle(IntPtr hWnd)
        {
            int len = GetWindowTextLength(hWnd);
            StringBuilder sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static void CloseDialog(IntPtr dialogHwnd)
        {
            // First try to find and click an OK button
            IntPtr okButton = FindChildButton(dialogHwnd, "OK");
            if (okButton != IntPtr.Zero)
            {
                SendMessage(okButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // Try any button if OK not found
            IntPtr anyButton = FindFirstButton(dialogHwnd);
            if (anyButton != IntPtr.Zero)
            {
                SendMessage(anyButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // Fallback
            PostMessage(dialogHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private static IntPtr FindChildButton(IntPtr parentHwnd, string caption)
        {
            IntPtr found = IntPtr.Zero;

            EnumChildWindows(parentHwnd, delegate (IntPtr hWnd, IntPtr lParam)
            {
                string cls = GetClassNameString(hWnd);
                if (!string.Equals(cls, "Button", StringComparison.OrdinalIgnoreCase))
                    return true;

                string text = GetControlText(hWnd);
                if (!string.IsNullOrWhiteSpace(text) &&
                    text.IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static IntPtr FindFirstButton(IntPtr parentHwnd)
        {
            IntPtr found = IntPtr.Zero;

            EnumChildWindows(parentHwnd, delegate (IntPtr hWnd, IntPtr lParam)
            {
                string cls = GetClassNameString(hWnd);
                if (string.Equals(cls, "Button", StringComparison.OrdinalIgnoreCase))
                {
                    found = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static string GetClassNameString(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string GetControlText(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(2048);
            SendMessage(hWnd, WM_GETTEXT, (IntPtr)sb.Capacity, sb);
            string text = sb.ToString();

            if (!string.IsNullOrWhiteSpace(text))
                return text;

            // Fallback
            int len = GetWindowTextLength(hWnd);
            if (len > 0)
            {
                sb.Clear();
                sb.EnsureCapacity(len + 1);
                GetWindowText(hWnd, sb, len + 1);
                text = sb.ToString();
            }

            return text;
        }

        private static string ReadDialogText(IntPtr dialogHwnd)
        {
            StringBuilder allText = new StringBuilder();

            // 1) Try UI Automation first (best for WPF/custom dialogs)
            try
            {
                AutomationElement root = AutomationElement.FromHandle(dialogHwnd);
                if (root != null)
                {
                    Condition cond = new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                    );

                    AutomationElementCollection elements =
                        root.FindAll(TreeScope.Descendants, cond);

                    for (int i = 0; i < elements.Count; i++)
                    {
                        try
                        {
                            string name = elements[i].Current.Name;
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                allText.AppendLine(name.Trim());
                            }
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }

            // 2) Fallback to Win32 child text scraping
            if (allText.Length == 0)
            {
                EnumChildWindows(dialogHwnd, delegate (IntPtr hWnd, IntPtr lParam)
                {
                    string text = GetControlText(hWnd);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        allText.AppendLine(text.Trim());
                    }
                    return true;
                }, IntPtr.Zero);
            }

            // 3) Final fallback to window caption
            if (allText.Length == 0)
            {
                string caption = GetWindowTitle(dialogHwnd);
                if (!string.IsNullOrWhiteSpace(caption))
                    allText.AppendLine(caption);
            }

            return allText.ToString();
        }

        #region Win32

        private const int WM_CLOSE = 0x0010;
        private const int WM_GETTEXT = 0x000D;
        private const int BM_CLICK = 0x00F5;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        #endregion
    }
}