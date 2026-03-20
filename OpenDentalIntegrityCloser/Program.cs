using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenDentalIntegrityCloser
{
    static class Program
    {
        private static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            bool createdNew;
            _mutex = new Mutex(true, "OpenDentalIntegrityCloserMutex", out createdNew);
            if (!createdNew)
            {
                return; // already running
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayAppContext());

            GC.KeepAlive(_mutex);
        }
    }

    public class TrayAppContext : ApplicationContext
    {
        private readonly NotifyIcon _notifyIcon;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly MenuItem _pauseMenuItem;
        private readonly MenuItem _startupMenuItem;

        private bool _paused = false;
        private bool _isHandlingTick = false;

        private const int CheckIntervalMs = 1500;
        private const string OpenDentalProcessName = "OpenDental";
        private const string OpenDentalTitleMatch = "Open Dental";
        private const string DialogTitleMatch = "Database Integrity";
        private const string RequiredDialogTextMatch =
            "Open Dental has detected that this patient's data has been modified";

        public TrayAppContext()
        {
            ContextMenu menu = new ContextMenu();

            _pauseMenuItem = new MenuItem("Pause", OnPauseResumeClicked);
            _startupMenuItem = new MenuItem("Run at Startup", OnStartupClicked);
            _startupMenuItem.Checked = StartupManager.IsStartupEnabled();
            MenuItem exitMenuItem = new MenuItem("Exit", OnExitClicked);

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

            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = CheckIntervalMs;
            _timer.Tick += Timer_Tick;
            _timer.Start();

            Log("Application started.");
        }

        private void OnStartupClicked(object sender, EventArgs e)
        {
            try
            {
                bool enable = !_startupMenuItem.Checked;
                StartupManager.SetStartup(enable);
                _startupMenuItem.Checked = enable;
                Log("Run at Startup set to " + enable + ".");
            }
            catch (Exception ex)
            {
                Log("OnStartupClicked error: " + ex);
            }
        }

        private void OnPauseResumeClicked(object sender, EventArgs e)
        {
            _paused = !_paused;
            _pauseMenuItem.Text = _paused ? "Resume" : "Pause";
            _notifyIcon.Text = _paused
                ? "OpenDental Integrity Closer (Paused)"
                : "OpenDental Integrity Closer";

            Log("Paused set to " + _paused + ".");
        }

        private void OnExitClicked(object sender, EventArgs e)
        {
            try
            {
                _timer.Stop();
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                Log("Application exiting.");
            }
            catch (Exception ex)
            {
                Log("OnExitClicked error: " + ex);
            }

            ExitThread();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_paused)
                return;

            if (_isHandlingTick)
                return;

            _isHandlingTick = true;

            try
            {
                if (!IsAnyOpenDentalMainWindowPresent())
                    return;

                IntPtr dialogHwnd = FindOpenDentalDialogWindow(DialogTitleMatch);
                if (dialogHwnd == IntPtr.Zero)
                    return;

                if (!IsWindow(dialogHwnd) || !IsWindowVisible(dialogHwnd))
                    return;

                if (!IsOpenDentalWindow(dialogHwnd))
                    return;

                string dialogText = ReadDialogText(dialogHwnd);
                if (string.IsNullOrWhiteSpace(dialogText))
                {
                    Log("Dialog found but no text could be read.");
                    return;
                }

                if (!string.IsNullOrWhiteSpace(RequiredDialogTextMatch))
                {
                    if (dialogText.IndexOf(RequiredDialogTextMatch, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        Log("Dialog title matched but text did not match. Text read: " + SanitizeForLog(dialogText));
                        return;
                    }
                }

                if (!IsWindow(dialogHwnd) || !IsWindowVisible(dialogHwnd))
                    return;

                CloseDialog(dialogHwnd);
            }
            catch (Exception ex)
            {
                Log("Timer_Tick error: " + ex);
            }
            finally
            {
                _isHandlingTick = false;
            }
        }

        private static bool IsAnyOpenDentalMainWindowPresent()
        {
            bool found = false;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                if (!IsOpenDentalWindow(hWnd))
                    return true;

                string title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title) &&
                    title.IndexOf(OpenDentalTitleMatch, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = true;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static IntPtr FindOpenDentalDialogWindow(string titlePart)
        {
            IntPtr foundHwnd = IntPtr.Zero;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                if (!IsOpenDentalWindow(hWnd))
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

        private static bool IsOpenDentalWindow(IntPtr hWnd)
        {
            try
            {
                uint pid;
                GetWindowThreadProcessId(hWnd, out pid);
                if (pid == 0)
                    return false;

                Process proc = Process.GetProcessById((int)pid);
                return proc.ProcessName.Equals(OpenDentalProcessName, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string GetWindowTitle(IntPtr hWnd)
        {
            int len = GetWindowTextLength(hWnd);
            if (len <= 0)
                return string.Empty;

            StringBuilder sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static void CloseDialog(IntPtr dialogHwnd)
        {
            if (!IsWindow(dialogHwnd))
                return;

            IntPtr okButton = FindChildButton(dialogHwnd, "OK");
            if (okButton != IntPtr.Zero && IsWindow(okButton))
            {
                Log("Clicking OK button on dialog.");
                SendMessage(okButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            IntPtr anyButton = FindFirstButton(dialogHwnd);
            if (anyButton != IntPtr.Zero && IsWindow(anyButton))
            {
                Log("OK button not found. Clicking first button on dialog.");
                SendMessage(anyButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            Log("No button found. Sending WM_CLOSE to dialog.");
            PostMessage(dialogHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private static IntPtr FindChildButton(IntPtr parentHwnd, string caption)
        {
            IntPtr found = IntPtr.Zero;

            EnumChildWindows(parentHwnd, delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindow(hWnd))
                    return true;

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
                if (!IsWindow(hWnd))
                    return true;

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

            int len = GetWindowTextLength(hWnd);
            if (len > 0)
            {
                sb = new StringBuilder(len + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                text = sb.ToString();
            }

            return text;
        }

        private static string ReadDialogText(IntPtr dialogHwnd)
        {
            if (!IsWindow(dialogHwnd))
                return string.Empty;

            StringBuilder allText = new StringBuilder();

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

                    AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, cond);

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
            catch (Exception ex)
            {
                Log("UI Automation read failed: " + ex.Message);
            }

            if (allText.Length == 0)
            {
                EnumChildWindows(dialogHwnd, delegate (IntPtr hWnd, IntPtr lParam)
                {
                    if (!IsWindow(hWnd))
                        return true;

                    string text = GetControlText(hWnd);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        allText.AppendLine(text.Trim());
                    }
                    return true;
                }, IntPtr.Zero);
            }

            if (allText.Length == 0)
            {
                string caption = GetWindowTitle(dialogHwnd);
                if (!string.IsNullOrWhiteSpace(caption))
                    allText.AppendLine(caption);
            }

            return allText.ToString().Trim();
        }

        private static void Log(string message)
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OpenDentalIntegrityCloser.log");
                File.AppendAllText(path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " - " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        private static string SanitizeForLog(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            return text.Replace("\r", " ").Replace("\n", " ").Trim();
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

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        #endregion
    }
        
}