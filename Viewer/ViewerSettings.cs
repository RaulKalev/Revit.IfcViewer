using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Centralised user-configurable viewer settings.
    /// Persisted to JSON in AppData/Roaming.
    /// Owned by <see cref="IfcViewer.UI.IfcViewerWindow"/> and
    /// forwarded to the relevant sub-systems on change.
    /// </summary>
    public sealed class ViewerSettings
    {
        // ── Walk / first-person ───────────────────────────────────────────────
        /// <summary>Base walk speed in metres per second (5 → ~human walking pace).</summary>
        public double WalkSpeed { get; set; } = 5.0;

        /// <summary>Speed multiplier when Shift is held during walk mode.</summary>
        public double SprintMultiplier { get; set; } = 3.0;

        /// <summary>Mouse-look sensitivity in radians per pixel.</summary>
        public double MouseSensitivity { get; set; } = 0.003;

        // ── Zoom (scroll wheel) ───────────────────────────────────────────────
        /// <summary>
        /// Fraction of the current camera-to-target distance moved per scroll notch.
        /// E.g. 0.1 = zoom 10 % of the remaining distance per notch.
        /// Range: 0.01 – 0.50.
        /// </summary>
        public double ZoomStep { get; set; } = 0.10;

        // ── Camera ────────────────────────────────────────────────────────────
        /// <summary>Perspective field of view in degrees.</summary>
        public double FieldOfView { get; set; } = 45.0;

        // ── Window geometry ───────────────────────────────────────────────────
        /// <summary>Saved window position and size. Null = use XAML defaults.</summary>
        public double? WindowLeft   { get; set; }
        public double? WindowTop    { get; set; }
        public double? WindowWidth  { get; set; }
        public double? WindowHeight { get; set; }

        // ── Revit 3D view preference ──────────────────────────────────────────
        /// <summary>
        /// Per-project saved 3D view for Revit geometry export.
        /// Key = Revit document path (PathName), Value = View3D name.
        /// Used when the active Revit view is not a 3D view so the user
        /// does not have to pick a view on every export.
        /// </summary>
        public Dictionary<string, string> SavedRevit3DViews { get; set; }
            = new Dictionary<string, string>();

        // ── Serialization ──────────────────────────────────────────────────────

        /// <summary>Path to the persisted settings JSON file in AppData.</summary>
        private static string SettingsPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk", "Revit", "Addins", "RKTools", "IfcViewer",
                "settings.json");

        /// <summary>Load settings from disk, or return defaults if file doesn't exist.</summary>
        public static ViewerSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    var loaded = JsonConvert.DeserializeObject<ViewerSettings>(json);
                    if (loaded != null) return loaded;
                }
            }
            catch (Exception ex)
            {
                SessionLogger.Warn($"Failed to load settings: {ex.Message}");
            }

            return new ViewerSettings(); // defaults
        }

        /// <summary>Save settings to disk.</summary>
        public void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                string json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(SettingsPath, json);
                SessionLogger.Info("Settings saved.");
            }
            catch (Exception ex)
            {
                SessionLogger.Error($"Failed to save settings: {ex.Message}");
            }
        }
    }
}
