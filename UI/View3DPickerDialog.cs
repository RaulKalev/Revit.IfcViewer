using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace IfcViewer.UI
{
    /// <summary>
    /// Simple modal dialog shown when the active Revit view is not a 3D view.
    /// Lets the user pick which View3D to use for geometry export; the choice is
    /// saved to <see cref="Viewer.ViewerSettings"/> so subsequent syncs are silent.
    ///
    /// Built entirely in code — no XAML resource dictionary required.
    /// </summary>
    internal sealed class View3DPickerDialog : Window
    {
        /// <summary>
        /// Name of the View3D chosen by the user, or <c>null</c> if the dialog
        /// was cancelled.
        /// </summary>
        public string SelectedViewName { get; private set; }

        private readonly ListBox _list;

        /// <param name="viewNames">Available 3D view names (ordered).</param>
        /// <param name="preselectedName">
        ///   Previously saved view name to pre-select, or <c>null</c> for none.
        /// </param>
        public View3DPickerDialog(IReadOnlyList<string> viewNames,
                                  string preselectedName = null)
        {
            Title                 = "Select 3D View — IFC Viewer";
            Width                 = 380;
            Height                = 420;
            MinWidth              = 280;
            MinHeight             = 260;
            ResizeMode            = ResizeMode.CanResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar         = false;

            // Root grid with three rows: message | list | buttons
            var grid = new Grid { Margin = new Thickness(14) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = grid;

            // ── Message ───────────────────────────────────────────────────────
            var msg = new TextBlock
            {
                Text         = "The active Revit view is not a 3D view.\n" +
                               "Select a 3D view to use for geometry export.\n" +
                               "Your selection will be remembered for this project.",
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 0, 0, 10),
            };
            Grid.SetRow(msg, 0);
            grid.Children.Add(msg);

            // ── View list ─────────────────────────────────────────────────────
            _list = new ListBox { Margin = new Thickness(0, 0, 0, 10) };
            foreach (var name in viewNames)
                _list.Items.Add(name);

            // Pre-select saved preference, or the first item if nothing saved.
            if (preselectedName != null && _list.Items.Contains(preselectedName))
            {
                _list.SelectedItem = preselectedName;
                _list.ScrollIntoView(preselectedName);
            }
            else if (_list.Items.Count > 0)
            {
                _list.SelectedIndex = 0;
            }

            // Double-click confirms immediately.
            _list.MouseDoubleClick += (_, __) =>
            {
                if (_list.SelectedItem != null) Accept();
            };

            Grid.SetRow(_list, 1);
            grid.Children.Add(_list);

            // ── Buttons ───────────────────────────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation         = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };

            var cancelBtn = new Button
            {
                Content  = "Cancel",
                Width    = 75,
                Height   = 26,
                Margin   = new Thickness(0, 0, 8, 0),
                IsCancel = true,   // Esc closes dialog with DialogResult = false
            };
            cancelBtn.Click += (_, __) => { DialogResult = false; };
            btnPanel.Children.Add(cancelBtn);

            var okBtn = new Button
            {
                Content   = "Use Selected",
                Width     = 96,
                Height    = 26,
                IsDefault = true,  // Enter activates this button
            };
            okBtn.Click += (_, __) =>
            {
                if (_list.SelectedItem != null) Accept();
            };
            btnPanel.Children.Add(okBtn);

            Grid.SetRow(btnPanel, 2);
            grid.Children.Add(btnPanel);
        }

        private void Accept()
        {
            SelectedViewName = _list.SelectedItem as string;
            DialogResult     = true;
        }
    }
}
