using System;
using System.IO;

namespace BetterMeV2VSTO
{
    internal static class Logger
    {
        private static readonly object _sync = new object();
        private static string LogFilePath
        {
            get
            {
                try
                {
                    var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BetterMeV2VSTO");
                    Directory.CreateDirectory(dir);
                    return Path.Combine(dir, "add_in_log.txt");
                }
                catch { return Path.GetTempFileName(); }
            }
        }

        public static void Log(string context, Exception ex)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}: {ex}\r\n";
                lock (_sync)
                {
                    File.AppendAllText(LogFilePath, line);
                }
            }
            catch { }
        }

        public static void Log(string context, string info)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}: {info}\r\n";
                lock (_sync)
                {
                    File.AppendAllText(LogFilePath, line);
                }
            }
            catch { }
        }
    }
}
