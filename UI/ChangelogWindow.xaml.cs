using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Text;

namespace ProSchedules.UI
{
    public partial class ChangelogWindow : Window
    {
        public ChangelogWindow(List<Services.UpdateLogService.VersionEntry> versions)
        {
            InitializeComponent();
            
            if (versions == null || versions.Count == 0) return;

            var newest = versions[0];
            VersionText.Text = $"Version {newest.Version}";
            
            // Format notes
            var sb = new StringBuilder();

            for (int i = 0; i < versions.Count; i++)
            {
                var v = versions[i];
                bool isNewest = (i == 0);

                // Separator for older versions (Header)
                // "The older updates should have a ----- version, released ----- as the separating row"
                if (!isNewest)
                {
                    sb.AppendLine($"----- {v.Version}, {v.Released} -----");
                    sb.AppendLine();
                }

                // Notes
                if (v.Notes != null)
                {
                    foreach (var note in v.Notes)
                    {
                        if (string.IsNullOrWhiteSpace(note)) continue;
                        sb.AppendLine(note.Trim());
                        sb.AppendLine();
                    }
                }

                // Footer for all updates
                // "All updates should have ----- released, developer ----- as the last row and an empty row below it"
                sb.AppendLine($"----- {v.Released}, {v.Developer} -----");
                sb.AppendLine();
                sb.AppendLine(); // Empty row below footer
            }
            
            NotesText.Text = sb.ToString().TrimEnd();

            // Allow moving window
            MouseLeftButtonDown += (s, e) => DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
