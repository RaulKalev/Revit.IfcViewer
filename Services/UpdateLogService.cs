using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using System.Windows;

namespace ProSchedules.Services
{
    public class UpdateLogService
    {
        private const string AppName = "ProSchedules";
        private const string StateFileName = "state.json";
        private const string RelativeChangelogPath = "ProSchedules/changelog.json"; // Default relative path

        public class Changelog
        {
            [JsonProperty("versions")]
            public List<VersionEntry> Versions { get; set; }
        }

        public class VersionEntry
        {
            [JsonProperty("version")]
            public string Version { get; set; }

            [JsonProperty("released")]
            public string Released { get; set; }

            [JsonProperty("developer")]
            public string Developer { get; set; }

            [JsonProperty("notes")]
            public object NotesRaw { get; set; } // Can be string or List<string>

            [JsonIgnore]
            public List<string> Notes
            {
                get
                {
                    if (NotesRaw is string s) return new List<string> { s };
                    if (NotesRaw is Newtonsoft.Json.Linq.JArray arr) return arr.ToObject<List<string>>();
                    if (NotesRaw is List<string> list) return list;
                    return new List<string>();
                }
            }
        }

        public class LocalState
        {
            public string LastShownVersion { get; set; } = "0.0.0";
        }

        public static void CheckAndShow(Window owner)
        {
            try
            {
                // 1. Find the file
                if (!TryResolveChangelogPath(out string path))
                {
                    // Debug: 
                    // MessageBox.Show("Could not find changelog file.");
                    return; 
                }

                // 2. Parse Remote
                var changelog = ParseChangelog(path);
                if (changelog == null || changelog.Versions == null || changelog.Versions.Count == 0) 
                {
                    // MessageBox.Show($"Found file at {path} but failed to parse or no versions.");
                    return;
                }

                // 3. Load Local State
                var state = LoadState();
                Version local = ParseVersion(state.LastShownVersion);

                // Filter for ALL newer versions
                var newerVersions = changelog.Versions
                    .Where(v => ParseVersion(v.Version) > local)
                    .OrderByDescending(v => ParseVersion(v.Version))
                    .ToList();

                if (newerVersions.Count == 0) return;

                // 4. Show UI with list of versions
                // Use the NEWEST version as the "Current" version for title/state
                var newest = newerVersions.First();
                
                var win = new UI.ChangelogWindow(newerVersions);
                win.Owner = owner;
                win.ShowDialog();

                // 5. Update State to newest
                state.LastShownVersion = newest.Version;
                SaveState(state);
            }
            catch (Exception ex)
            {
                // Verify failure silently or log debug
                System.Diagnostics.Debug.WriteLine($"Update check error: {ex.Message}");
            }
        }

        private static bool TryResolveChangelogPath(out string path)
        {
            path = null;
            // string debugLog = "";

            // 1. Env Variable
            string envPath = Environment.GetEnvironmentVariable("PROSCHEDULES_CHANGELOG_PATH");
            if (!string.IsNullOrEmpty(envPath) && File.Exists(envPath))
            {
                path = envPath;
                return true;
            }

            // 2. Heuristic Search
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            
            // List of potential cloud roots
            var roots = new List<string>
            {
                "EULE Dropbox", // Prioritize EULE since that's what the user has
                "Dropbox",
                "OneDrive", 
                "Google Drive",
                "Documents" 
            };

            foreach (var rootName in roots)
            {
                string rootPath = Path.Combine(userProfile, rootName);
                // debugLog += $"Checking root: {rootPath} (Exists: {Directory.Exists(rootPath)})\n";

                if (Directory.Exists(rootPath))
                {
                    // Check standard relative path
                    string fullPath = Path.Combine(rootPath, RelativeChangelogPath);
                    if (File.Exists(fullPath))
                    {
                        path = fullPath;
                        return true;
                    }

                    // Specific EULE Team path
                    string eulePath = Path.Combine(rootPath, "0_EULE  Team folder (kogu kollektiiv)", "02_EULE REVIT TEMPLATE", "099-scriptid", "Pluginad", "UuendusteInfo", "ProSchedule_CHANGELOG.json");
                    // debugLog += $"  Checking specific EULE path: {eulePath} (Exists: {File.Exists(eulePath)})\n";

                    if (File.Exists(eulePath))
                    {
                        path = eulePath;
                        return true;
                    }
                    
                    // Also check for "MyTeam/Scripts/App_CHANGELOG.json"
                    string customPath = Path.Combine(rootPath, "MyTeam", "Scripts", "ProSchedules_CHANGELOG.json");
                    if (File.Exists(customPath))
                    {
                        path = customPath;
                        return true;
                    }
                }
            }
            
            // If we get here, verify failure
            // MessageBox.Show(debugLog, "Changelog Path Not Found");

            return false;
        }

        private static Changelog ParseChangelog(string path)
        {
            try
            {
                string json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<Changelog>(json);
            }
            catch
            {
                return null;
            }
        }

        private static LocalState LoadState()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string configPath = Path.Combine(appData, AppName, StateFileName);

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    return JsonConvert.DeserializeObject<LocalState>(json) ?? new LocalState();
                }
            }
            catch { }

            return new LocalState();
        }

        private static void SaveState(LocalState state)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string folder = Path.Combine(appData, AppName);
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);

                string configPath = Path.Combine(folder, StateFileName);
                string json = JsonConvert.SerializeObject(state, Formatting.Indented);
                File.WriteAllText(configPath, json);
            }
            catch { }
        }

        private static Version ParseVersion(string v)
        {
            if (string.IsNullOrWhiteSpace(v)) return new Version(0, 0, 0);
            // Handle simple text or "v1.0"
            v = v.Replace("v", "").Trim();
            if (Version.TryParse(v, out var result)) return result;
            return new Version(0, 0, 0);
        }
    }
}
