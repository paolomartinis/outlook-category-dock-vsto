using System;
using System.IO;

namespace CategoryDockVsto
{
    internal static class Logger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CategoryDockVsto",
            "category-dock-vsto.log");

        public static void Write(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, DateTime.Now.ToString("s") + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        public static void Write(Exception exception)
        {
            Write(exception.GetType().FullName + ": " + exception.Message + Environment.NewLine + exception.StackTrace);
        }
    }
}
