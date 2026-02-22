using System;
using System.Diagnostics;
using System.IO;

namespace IfcViewer
{
    /// <summary>
    /// Minimal session logger.
    /// Writes to Debug output and optionally to a rolling log file under
    /// %ProgramData%\Autodesk\Revit\Addins\RKTools\IfcViewer\Logs\
    /// </summary>
    public static class SessionLogger
    {
        private static readonly string LogFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "Autodesk", "Revit", "Addins", "RKTools", "IfcViewer", "Logs");

        private static readonly string LogFile = Path.Combine(
            LogFolder,
            $"IfcViewer_{DateTime.Now:yyyyMMdd}.log");

        private static bool _initialized;

        private static void EnsureInitialized()
        {
            if (_initialized) return;
            try
            {
                Directory.CreateDirectory(LogFolder);
                _initialized = true;
            }
            catch
            {
                // If we can't create the folder, log to Debug only.
                _initialized = true;
            }
        }

        public static void Info(string message) => Write("INFO ", message);
        public static void Warn(string message) => Write("WARN ", message);
        public static void Error(string message, Exception ex = null)
        {
            Write("ERROR", message);
            if (ex != null) Write("ERROR", ex.ToString());
        }

        private static void Write(string level, string message)
        {
            EnsureInitialized();
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}";
            Debug.WriteLine(line);
            try
            {
                File.AppendAllText(LogFile, line + Environment.NewLine);
            }
            catch
            {
                // Swallow file I/O errors — logs are best-effort.
            }
        }
    }
}
