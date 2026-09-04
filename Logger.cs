using System;
using System.IO;

namespace OutlookJunkRescuer
{
    internal static class Logger
    {
        private static readonly object Sync = new object();

        public static void Write(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookJunkRescuer");

                Directory.CreateDirectory(dir);

                lock (Sync)
                {
                    File.AppendAllText(
                        Path.Combine(dir, "OutlookJunkRescuer.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}" +
                        Environment.NewLine);
                }
            }
            catch
            {
                // Logging is deliberately non-fatal.
            }
        }
    }
}
