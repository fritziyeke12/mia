using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

public static class InstallDateResolver
{
    private static Dictionary<string, string> _msiInstallDateCache;

    public static List<InstalledApp> GetInstalledApps()
    {
        var apps = new List<InstalledApp>();

        string[] registryKeys = {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        };

        foreach (var root in new[] { Registry.CurrentUser, Registry.LocalMachine })
        {
            foreach (var keyPath in registryKeys)
            {
                using var key = root.OpenSubKey(keyPath);
                if (key == null) continue;

                foreach (var subkeyName in key.GetSubKeyNames())
                {
                    using var appKey = key.OpenSubKey(subkeyName);
                    if (appKey == null) continue;

                    string name = appKey.GetValue("DisplayName") as string;
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    // Skip hidden/system entries
                    if (appKey.GetValue("SystemComponent") is int sc && sc == 1) continue;
                    if (appKey.GetValue("NoDisplay") is int nd && nd == 1) continue;

                    string version = appKey.GetValue("DisplayVersion") as string;
                    string publisher = appKey.GetValue("Publisher") as string;
                    string installLocation = appKey.GetValue("InstallLocation") as string;
                    string displayIconRaw = appKey.GetValue("DisplayIcon") as string;
                    string displayIcon = displayIconRaw?.Split(',')[0].Trim('"');
                    string installDateRaw = appKey.GetValue("InstallDate") as string;

                    DateTime? registryInstallDate = null;
                    if (!string.IsNullOrWhiteSpace(installDateRaw) &&
                        installDateRaw.Length == 8 &&
                        DateTime.TryParseExact(installDateRaw, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var parsedDate))
                    {
                        registryInstallDate = parsedDate;
                    }

                    apps.Add(new InstalledApp
                    {
                        Name = name,
                        Publisher = publisher,
                        Version = version,
                        InstallLocation = installLocation,
                        DisplayIcon = GetLongPath(displayIcon),
                        RegistryInstallDate = registryInstallDate
                    });
                }
            }
        }

        CacheMsiInstallDates();

        foreach (var app in apps)
        {
            app.ResolvedInstallDate = ResolveInstallDate(app);
        }

        return apps;
    }

    private static DateTime? ResolveInstallDate(InstalledApp app)
    {
        // 1. Registry install date
        if (app.RegistryInstallDate.HasValue)
            return app.RegistryInstallDate;

        // 2. Win32_Product MSI date
        if (_msiInstallDateCache.TryGetValue(app.Name, out var msiDateStr) &&
            DateTime.TryParseExact(msiDateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var msiDate))
        {
            return msiDate;
        }

        // 3. DisplayIcon
        if (!string.IsNullOrWhiteSpace(app.DisplayIcon) && File.Exists(app.DisplayIcon))
        {
            return File.GetCreationTime(app.DisplayIcon);
        }

        // 4. InstallLocation
        if (!string.IsNullOrWhiteSpace(app.InstallLocation) && Directory.Exists(app.InstallLocation))
        {
            return Directory.GetCreationTime(app.InstallLocation);
        }

        return null;
    }

    private static void CacheMsiInstallDates()
    {
        _msiInstallDateCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var searcher = new ManagementObjectSearcher("SELECT Name, InstallDate FROM Win32_Product");
            foreach (ManagementObject obj in searcher.Get())
            {
                string name = obj["Name"]?.ToString();
                string date = obj["InstallDate"]?.ToString();

                if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(date))
                {
                    _msiInstallDateCache[name] = date;
                }
            }
        }
        catch
        {
            // Handle WMI errors silently or log as needed
        }
    }

    // Convert short paths like C:\PROGRA~1 to long form
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern uint GetLongPathName(string shortPath, StringBuilder longPathBuffer, uint bufferLength);

    private static string GetLongPath(string shortPath)
    {
        if (string.IsNullOrWhiteSpace(shortPath)) return null;

        var buffer = new StringBuilder(260);
        uint result = GetLongPathName(shortPath, buffer, (uint)buffer.Capacity);

        return result > 0 ? buffer.ToString() : shortPath;
    }
}
