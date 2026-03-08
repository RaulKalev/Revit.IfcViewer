using System;
using System.Diagnostics;
using System.IO;

namespace IfcCore
{
    /// <summary>
    /// Structured logger for IfcCore operations.
    /// Writes to Debug output and optionally to a rolling log file.
    /// Designed to never throw — all I/O errors are swallowed.
    /// </summary>
    public static class IfcCoreLogger
    {
        private static readonly string LogFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RKTools", "IfcViewer", "Logs");

        private static readonly string LogFile = Path.Combine(
            LogFolder,
            $"IfcCore_{DateTime.Now:yyyyMMdd}.log");

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
                _initialized = true; // Don't retry
            }
        }

        public static void Info(string message) => Write("INFO ", message);
        public static void Warn(string message) => Write("WARN ", message);
        public static void Error(string message, Exception ex = null)
        {
            Write("ERROR", message);
            if (ex != null) Write("ERROR", ex.ToString());
        }

        /// <summary>
        /// Logs a structured summary of a completed load operation.
        /// </summary>
        public static void LogLoadSummary(IfcLoadResult result)
        {
            if (result == null) return;

            var lines = new[]
            {
                "────────────────────────────────────────────────",
                $"IFC Load Summary",
                $"  File:     {result.FilePath}",
                $"  Schema:   {result.SchemaVersion ?? "Unknown"}",
                $"  Provider: {result.StorageProvider ?? "Unknown"}",
                $"  Entities: {result.EntityCount}",
                $"  Products: {result.ProductCount}",
                $"  Duration: {result.Duration.TotalMilliseconds:F0} ms",
                $"  Success:  {result.Success}",
                result.ErrorMessage != null ? $"  Error:    {result.ErrorMessage}" : null,
                result.Warnings.Count > 0 ? $"  Warnings: {result.Warnings.Count}" : null,
                "────────────────────────────────────────────────",
            };

            foreach (var line in lines)
            {
                if (line != null)
                    Write("INFO ", line);
            }
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
                // Swallow — logs are best-effort.
            }
        }
    }
}
