using IfcViewer.Viewer;
using System;
using System.Windows;
using System.Windows.Input;

namespace IfcViewer.UI
{
    /// <summary>
    /// Floating settings dialog for the IFC Viewer.
    /// Reads/writes <see cref="ViewerSettings"/> in real-time; the
    /// caller supplies a <see cref="Action"/> callback that is invoked
    /// whenever any value changes so the viewer can apply them immediately.
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private readonly ViewerSettings _settings;
        private readonly Action         _onChange;
        private bool _loading;
        private bool _hasChanges;  // track if user has modified anything since load

        /// <summary>
        /// Create the settings window.
        /// </summary>
        /// <param name="settings">The shared settings object to read/write.</param>
        /// <param name="onChange">Called after every user change so the viewer applies new values immediately.</param>
        public SettingsWindow(ViewerSettings settings, Action onChange)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _onChange = onChange;

            // Set _loading BEFORE InitializeComponent so that ValueChanged events
            // fired by the XAML parser (while it sets slider Minimum/Maximum/Value)
            // don't overwrite the saved settings with XAML default values.
            _loading = true;
            InitializeComponent();

            LoadFromSettings(); // sets _loading = true again, writes slider values, then resets to false
        }

        // ── Load current values into controls ─────────────────────────────────

        private void LoadFromSettings()
        {
            _loading = true;

            WalkSpeedSlider.Value  = _settings.WalkSpeed;
            SprintSlider.Value     = _settings.SprintMultiplier;

            // Mouse sensitivity is stored as radians/px (tiny float).
            // Map 0.001–0.010 to slider range 1–10 by multiplying by 1000.
            MouseSensSlider.Value  = Math.Round(_settings.MouseSensitivity * 1000.0);

            // ZoomStep is stored as 0.01–0.50; slider shows percentage 1–50.
            ZoomStepSlider.Value   = Math.Round(_settings.ZoomStep * 100.0);

            FovSlider.Value        = _settings.FieldOfView;

            _loading = false;
            RefreshLabels();
        }

        private void RefreshLabels()
        {
            WalkSpeedLabel.Text  = _settings.WalkSpeed.ToString("F1");
            SprintLabel.Text     = _settings.SprintMultiplier.ToString("F1") + "x";
            MouseSensLabel.Text  = ((int)Math.Round(_settings.MouseSensitivity * 1000)).ToString();
            ZoomStepLabel.Text   = ((int)Math.Round(_settings.ZoomStep * 100)) + "%";
            FovLabel.Text        = ((int)_settings.FieldOfView) + "°";
        }

        // ── Slider change handlers ────────────────────────────────────────────

        private void WalkSpeed_ValueChanged(object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (_loading || _settings == null) return;
            _settings.WalkSpeed = e.NewValue;
            WalkSpeedLabel.Text = e.NewValue.ToString("F1");
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void Sprint_ValueChanged(object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (_loading || _settings == null) return;
            _settings.SprintMultiplier = e.NewValue;
            SprintLabel.Text = e.NewValue.ToString("F1") + "x";
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void MouseSens_ValueChanged(object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (_loading || _settings == null) return;
            _settings.MouseSensitivity = e.NewValue / 1000.0;
            MouseSensLabel.Text = ((int)e.NewValue).ToString();
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void ZoomStep_ValueChanged(object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (_loading || _settings == null) return;
            _settings.ZoomStep = e.NewValue / 100.0;
            ZoomStepLabel.Text = ((int)e.NewValue) + "%";
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void Fov_ValueChanged(object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (_loading || _settings == null) return;
            _settings.FieldOfView = e.NewValue;
            FovLabel.Text = ((int)e.NewValue) + "°";
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void UpdateApplyButtonState()
        {
            if (ApplyButton != null)
                ApplyButton.IsEnabled = _hasChanges;
        }

        // ── Reset ─────────────────────────────────────────────────────────────

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            // Overwrite with fresh defaults
            _settings.WalkSpeed        = 5.0;
            _settings.SprintMultiplier = 3.0;
            _settings.MouseSensitivity = 0.003;
            _settings.ZoomStep         = 0.10;
            _settings.FieldOfView      = 45.0;

            LoadFromSettings();
            _hasChanges = true;
            UpdateApplyButtonState();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            // Apply the current settings to the viewer and save to disk
            _onChange?.Invoke();
            _settings.Save();
            _hasChanges = false;
            UpdateApplyButtonState();
            SessionLogger.Info("Settings applied and saved.");
        }

        // ── Window chrome ─────────────────────────────────────────────────────

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
            => Close();
    }
}
