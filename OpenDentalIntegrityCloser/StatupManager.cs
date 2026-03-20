using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

public static class StartupManager
{
    private const string RUN_KEY = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string APP_NAME = "OpenDentalIntegrityCloser";

    public static void SetStartup(bool enable)
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RUN_KEY, true))
        {
            if (key == null)
                return;

            if (enable)
            {
                string exePath = Assembly.GetExecutingAssembly().Location;

                // Quote path in case of spaces
                key.SetValue(APP_NAME, "\"" + exePath + "\"");
            }
            else
            {
                key.DeleteValue(APP_NAME, false);
            }
        }
    }

    public static bool IsStartupEnabled()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RUN_KEY, false))
        {
            if (key == null)
                return false;

            return key.GetValue(APP_NAME) != null;
        }
    }
}