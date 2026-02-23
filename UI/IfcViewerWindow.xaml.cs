using Autodesk.Revit.UI;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Ifc;
using IfcViewer.Revit;
using IfcViewer.Viewer;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace IfcViewer.UI
{
    /// <summary>
    /// Main modeless viewer window.
    /// NOTE: No Helix types appear in the XAML — Viewport3DX is created
    /// entirely in code after App.OnStartup has registered the assembly resolver.
    /// This prevents the XAML parser from triggering a Helix assembly load
    /// before the resolver hook is in place.
    /// </summary>
    public partial class IfcViewerWindow : Window
    {
        // ── P/Invoke for window resizing ──────────────────────────────────────
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTLEFT           = 10;
        private const int HTRIGHT          = 11;
        private const int HTBOTTOM         = 15;
        private const int HTBOTTOMLEFT     = 16;
        private const int HTBOTTOMRIGHT    = 17;


        // ── State ─────────────────────────────────────────────────────────────
        private readonly UIApplication _uiApp;
        private ViewerHost   _viewerHost;
        private Viewport3DX  _viewport;
        private GroupModel3D _sceneRoot;
        private bool _isDarkTheme = true;

        // Loaded IFC models — bound to ModelListBox
        private readonly ObservableCollection<IfcModel> _loadedModels
            = new ObservableCollection<IfcModel>();

        // Stage 3: Revit sync glue
        private SyncRevitEvent _syncRevitEvent;
        private RevitModel     _revitModel;

        // Stage 4a: first-person controller + section plane + settings
        private FirstPersonController _fpController;
        private SectionPlaneManager   _sectionMgr;
        private ViewerSettings        _settings;

        // ── Constructor ───────────────────────────────────────────────────────
        public IfcViewerWindow(UIApplication uiApp)
        {
            _uiApp = uiApp;
            InitializeComponent();
            Loaded += OnWindowLoaded;
        }

        // ── Loaded: create Viewport3DX in code AFTER resolver is registered ──
        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Create the ViewerHost (EffectsManager + Camera)
                _viewerHost = new ViewerHost();

                // 2. Create Viewport3DX entirely in code — no XAML reference
                _viewport = new Viewport3DX
                {
                    EffectsManager           = _viewerHost.EffectsManager,
                    Camera                   = _viewerHost.Camera,
                    ShowCoordinateSystem     = false,
                    ShowFrameRate            = true,
                    EnableSSAO               = false,
                    // MSAA off — pure technical viewer, FXAA is sufficient and costs ~0ms
                    MSAA                     = MSAALevel.Disable,
                    Background               = new System.Windows.Media.SolidColorBrush(
                                                  System.Windows.Media.Color.FromRgb(26, 26, 26)),
                    // FXAA Medium replaces MSAA — single post-process pass, near-free
                    FXAALevel                = FXAALevel.Medium,
                    // Camera controller needs focus to receive mouse/keyboard input
                    Focusable                = true,
                    IsTabStop                = true,
                    // Zoom to cursor
                    ZoomAroundMouseDownPoint = true,
                    CameraMode               = CameraMode.Inspect,
                    // Disable default bindings so we can install our own below
                    UseDefaultGestures       = false,
                    // Shadows off — pure technical viewer
                    IsShadowMappingEnabled   = false,
                };

                // 3. Scene root
                _sceneRoot = new GroupModel3D();
                _viewport.Items.Add(_sceneRoot);

                // 4. Insert viewport as first child of ViewportContainer (behind the status bar)
                // Give keyboard focus on any click so CameraController receives events.
                _viewport.MouseDown += (s, ev) => _viewport.Focus();
                ViewportContainer.Children.Insert(0, _viewport);

                // 5a. Wire custom mouse bindings once the template is applied.
                //     UseDefaultGestures=false clears Helix's built-ins; we re-add only
                //     what we want. The bindings go on the viewport itself — the
                //     CameraController child handles the matching RoutedCommands.
                //     Convention: right-click drag = rotate, middle-click drag = pan,
                //                 scroll wheel = zoom (Helix handles scroll internally).
                _viewport.IsPanEnabled    = true;
                _viewport.IsZoomEnabled   = true;

                _viewport.Loaded += (s, ev) =>
                {
                    _viewport.InputBindings.Clear();

                    // Right-click drag → Rotate
                    _viewport.InputBindings.Add(new MouseBinding(
                        ViewportCommands.Rotate,
                        new MouseGesture(MouseAction.RightClick)));

                    // Middle-click drag → Pan
                    _viewport.InputBindings.Add(new MouseBinding(
                        ViewportCommands.Pan,
                        new MouseGesture(MouseAction.MiddleClick)));

                    // Note: Viewport3DX always uses DX11ImageSourceRenderHost (WPF D3DImage
                    // path) regardless of AllowsTransparency. This architecture has no
                    // GPU-level VSync — tearing is a known limitation of HelixToolkit.Wpf.SharpDX
                    // inside a WPF window. No further throttle attempts are made here.
                };

                // 5. Build test scene
                _viewerHost.BuildTestScene(_sceneRoot);

                // 6. Bind the model list
                ModelListBox.ItemsSource = _loadedModels;

                // 7. Create the Revit sync ExternalEvent (must be on UI thread)
                _syncRevitEvent = new SyncRevitEvent(Dispatcher);

                // 8. First-person controller
                //    keyTarget   = viewport (keyboard events)
                //    mouseTarget = this Window (PreviewMouse tunnel fires before Helix)
                _fpController = new FirstPersonController(_viewerHost.Camera, _viewport, this);

                // Block WASD / arrow keys from reaching Helix's CameraController when
                // walk mode is NOT active. Without this guard, any mouse click that gives
                // the viewport focus lets Helix handle those keys internally, causing the
                // camera to move even in normal orbit mode.
                this.PreviewKeyDown += (s, ev) =>
                {
                    if (_fpController != null && !_fpController.IsActive && IsMovementKey(ev.Key))
                        ev.Handled = true;
                };

                // 9. Section plane manager + attach its visual quad to the scene root
                _sectionMgr = new SectionPlaneManager();
                _sectionMgr.AttachVisual(_sceneRoot);

                // 10. Load settings from disk (or defaults) and apply them
                _settings = ViewerSettings.Load();
                ApplySettings();

                // 11. Intercept scroll wheel at Window level to implement instant, inertia-free
                //     zoom. Mark e.Handled=true so Helix's smooth animated zoom never fires.
                this.PreviewMouseWheel += OnPreviewMouseWheel;

                UpdateStatus($"GPU Viewport Active");
                SessionLogger.Info("Viewport3DX created in code-behind — test scene rendered.");
            }
            catch (Exception ex)
            {
                SessionLogger.Error("Failed to initialize Helix viewport.", ex);
                UpdateStatus($"Viewport error: {ex.Message}");
            }
        }

        // ── Close: dispose DirectX resources ─────────────────────────────────
        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                this.PreviewMouseWheel -= OnPreviewMouseWheel;
                _fpController?.Dispose();
                _sectionMgr?.DetachVisual();
                _syncRevitEvent?.Dispose();
                _sceneRoot?.Children.Clear();
                _viewport?.Items.Clear();
                _viewerHost?.Dispose();
            }
            catch (Exception ex)
            {
                SessionLogger.Error("Error during viewport disposal.", ex);
            }
        }

        // ── Title-bar ─────────────────────────────────────────────────────────
        private void TitleBar_Loaded(object sender, RoutedEventArgs e) { }
        private void Window_MouseMove(object sender, MouseEventArgs e) { }
        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e) { }

        // ── Theme toggle ──────────────────────────────────────────────────────
        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = !_isDarkTheme;
            SessionLogger.Info($"Theme toggled → {(_isDarkTheme ? "Dark" : "Light")}");
        }

        // ── Toolbar ───────────────────────────────────────────────────────────
        private async void AddIfc_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title            = "Open IFC file(s)",
                Filter           = "IFC Files (*.ifc)|*.ifc|All Files (*.*)|*.*",
                Multiselect      = true,
                CheckFileExists  = true,
            };
            if (dlg.ShowDialog(this) != true) return;

            foreach (string path in dlg.FileNames)
            {
                // Guard: don't load the same file twice
                bool alreadyLoaded = false;
                foreach (var m in _loadedModels)
                    if (string.Equals(m.FilePath, path, StringComparison.OrdinalIgnoreCase))
                    { alreadyLoaded = true; break; }
                if (alreadyLoaded) continue;

                UpdateStatus("Loading: " + System.IO.Path.GetFileName(path) + " …");
                AddIfcButton.IsEnabled    = false;
                RemoveIfcButton.IsEnabled = false;

                try
                {
                    IfcModel ifcModel = await IfcLoader.LoadAsync(path, Dispatcher,
                        onProgress: msg => UpdateStatus(msg));

                    // Attach to scene on UI thread
                    _viewerHost.IfcRoot.Children.Add(ifcModel.SceneGroup);
                    _loadedModels.Add(ifcModel);
                    ModelListBox.SelectedItem = ifcModel;

                    // Register all cross-section meshes with the section plane manager
                    _sectionMgr?.RegisterGroup(ifcModel.SceneGroup);

                    // Update section slider range to cover the full scene
                    UpdateSectionBounds();

                    // Fit camera to the loaded geometry
                    _viewerHost.FitView(ifcModel.Bounds);

                    UpdateStatus(ifcModel.DisplayName + "  |  " +
                                 ifcModel.MeshCount + " elements  |  " +
                                 ifcModel.TriangleCount + " triangles");
                    SessionLogger.Info("Loaded: " + ifcModel.DisplayName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this,
                        "Failed to load IFC file:\n\n" + ex.Message,
                        "IFC Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    SessionLogger.Error("IFC load failed: " + path, ex);
                    UpdateStatus("Load failed — see log.");
                }
                finally
                {
                    AddIfcButton.IsEnabled    = true;
                    RemoveIfcButton.IsEnabled = true;
                }
            }
        }

        private void RemoveIfc_Click(object sender, RoutedEventArgs e)
        {
            if (!(ModelListBox.SelectedItem is IfcModel selected)) return;

            // Remove from scene
            _sectionMgr?.UnregisterGroup(selected.SceneGroup);
            _viewerHost.IfcRoot.Children.Remove(selected.SceneGroup);
            _loadedModels.Remove(selected);

            UpdateStatus(_loadedModels.Count == 0
                ? "GPU Viewport Active"
                : _loadedModels.Count + " model(s) loaded");
            SessionLogger.Info("Removed: " + selected.DisplayName);
        }

        private void SyncRevit_Click(object sender, RoutedEventArgs e)
        {
            if (_syncRevitEvent == null) return;

            // Disable the button while export runs
            SyncRevitButton.IsEnabled = false;
            UpdateStatus("Exporting Revit geometry…");

            _syncRevitEvent.Request(
                onComplete: model =>
                {
                    // Remove any previously exported Revit geometry
                    if (_revitModel != null)
                    {
                        _sectionMgr?.UnregisterGroup(_revitModel.SceneGroup);
                        _viewerHost.RevitRoot.Children.Remove(_revitModel.SceneGroup);
                    }

                    _revitModel = model;
                    _viewerHost.RevitRoot.Children.Add(model.SceneGroup);
                    _sectionMgr?.RegisterGroup(model.SceneGroup);
                    UpdateSectionBounds();

                    // If no IFC is loaded, fit camera to Revit geometry
                    if (_loadedModels.Count == 0)
                        _viewerHost.FitView(model.Bounds);

                    SyncRevitButton.IsEnabled = true;
                    UpdateStatus($"Revit: {model.DisplayName}  |  {model.MeshCount} materials  |  {model.TriangleCount} triangles");
                    SessionLogger.Info($"Revit sync complete: {model.TriangleCount} triangles");
                },
                onError: ex =>
                {
                    SyncRevitButton.IsEnabled = true;
                    UpdateStatus("Revit export failed — see log.");
                    MessageBox.Show(this,
                        "Failed to export Revit geometry:\n\n" + ex.Message,
                        "Revit Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    SessionLogger.Error("Revit export failed.", ex);
                });
        }

        private void ResetCamera_Click(object sender, RoutedEventArgs e)
            => _viewerHost?.ResetCamera();

        private void Wireframe_Checked(object sender, RoutedEventArgs e)
            => ApplyWireframe(true);

        private void Wireframe_Unchecked(object sender, RoutedEventArgs e)
            => ApplyWireframe(false);

        private void ApplyWireframe(bool on)
        {
            if (_sceneRoot != null) ApplyWireframeToGroup(_sceneRoot, on);
            SessionLogger.Info($"Wireframe {(on ? "on" : "off")}.");
        }

        private static void ApplyWireframeToGroup(GroupModel3D group, bool on)
        {
            foreach (var child in group.Children)
            {
                if (child is MeshGeometryModel3D mesh)
                    mesh.RenderWireframe = on;
                else if (child is GroupModel3D sub)
                    ApplyWireframeToGroup(sub, on);
            }
        }

        // ── Opacity sliders ───────────────────────────────────────────────────
        private void IfcOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_viewerHost == null) return;
            SetGroupOpacity(_viewerHost.IfcRoot, e.NewValue);
        }

        private void RevitOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_viewerHost == null) return;
            SetGroupOpacity(_viewerHost.RevitRoot, e.NewValue);
        }

        private static void SetGroupOpacity(GroupModel3D group, double opacity)
        {
            foreach (var child in group.Children)
            {
                if (child is MeshGeometryModel3D mesh && mesh.Material is PhongMaterial mat)
                {
                    var c = mat.DiffuseColor;
                    mat.DiffuseColor = new SharpDX.Color4(c.Red, c.Green, c.Blue, (float)opacity);
                }
                else if (child is GroupModel3D sub)
                    SetGroupOpacity(sub, opacity);
            }
        }

        // ── Status / helpers ──────────────────────────────────────────────────
        public void UpdateStatus(string text)
        {
            if (StatusBar != null)
                StatusBar.Text = text;
        }

        public ViewerHost ViewerHost => _viewerHost;
        public GroupModel3D SceneRoot => _sceneRoot;

        private int CountTriangles()
            => _sceneRoot != null ? CountTrianglesInGroup(_sceneRoot) : 0;

        private static int CountTrianglesInGroup(GroupModel3D group)
        {
            int total = 0;
            foreach (var child in group.Children)
            {
                if (child is MeshGeometryModel3D mesh && mesh.Geometry?.Indices != null)
                    total += mesh.Geometry.Indices.Count / 3;
                else if (child is GroupModel3D sub)
                    total += CountTrianglesInGroup(sub);
            }
            return total;
        }

        // ── Walk mode ─────────────────────────────────────────────────────────

        /// <summary>Keys that Helix's CameraController uses for movement — blocked in orbit mode.</summary>
        private static bool IsMovementKey(Key k)
            => k == Key.W || k == Key.A || k == Key.S || k == Key.D
            || k == Key.Up || k == Key.Down || k == Key.Left || k == Key.Right
            || k == Key.Q || k == Key.E || k == Key.PageUp || k == Key.PageDown;

        private void WalkMode_Checked(object sender, RoutedEventArgs e)
        {
            if (_fpController == null || _viewport == null) return;

            // Reduce near plane so geometry doesn't clip when camera is close/inside surfaces
            _viewerHost.SetWalkMode(true);

            // Switch Helix to WalkAround and disable its rotation so it won't fight us.
            // Pan is also disabled — our PreviewMouse handler owns right-drag fully.
            _viewport.CameraMode        = CameraMode.WalkAround;
            _viewport.IsRotationEnabled = false;
            _viewport.IsPanEnabled      = false;

            _fpController.Activate();
            UpdateStatus("Walk mode  |  WASD / arrows = move  |  Right-drag = look  |  Q/E = up/down  |  Shift = sprint");
            SessionLogger.Info("Walk mode activated.");
        }

        private void WalkMode_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_fpController == null || _viewport == null) return;

            _fpController.Deactivate();

            // Restore orbit mode
            _viewerHost.SetWalkMode(false);
            _viewport.CameraMode        = CameraMode.Inspect;
            _viewport.IsRotationEnabled = true;
            _viewport.IsPanEnabled      = true;

            UpdateStatus("Orbit mode restored.");
            SessionLogger.Info("Walk mode deactivated.");
        }

        // ── Section plane ─────────────────────────────────────────────────────
        private void SectionPlane_Checked(object sender, RoutedEventArgs e)
        {
            if (_sectionMgr == null) return;
            _sectionMgr.Enabled = true;
            SessionLogger.Info("Section plane enabled.");
        }

        private void SectionPlane_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_sectionMgr == null) return;
            _sectionMgr.Enabled = false;
            SessionLogger.Info("Section plane disabled.");
        }

        private void SectionAxis_Changed(object sender, RoutedEventArgs e)
        {
            if (_sectionMgr == null) return;
            if (SectionAxisX?.IsChecked == true)      _sectionMgr.Axis = SectionAxis.X;
            else if (SectionAxisY?.IsChecked == true) _sectionMgr.Axis = SectionAxis.Y;
            else                                      _sectionMgr.Axis = SectionAxis.Z;

            // Re-centre slider range on the new axis
            UpdateSectionBounds();
        }

        private void SectionOffset_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_sectionMgr == null) return;
            _sectionMgr.Offset = (float)e.NewValue;
            if (SectionOffsetLabel != null)
                SectionOffsetLabel.Text = e.NewValue.ToString("F2");
        }

        /// <summary>
        /// Recalculate the section slider min/max from the combined scene bounding box
        /// so the slider always covers the full extent of loaded geometry.
        /// </summary>
        private void UpdateSectionBounds()
        {
            if (_sectionMgr == null || SectionOffsetSlider == null) return;

            // Collect all known bounds
            float min = float.MaxValue, max = float.MinValue;

            foreach (var m in _loadedModels)
            {
                var b = m.Bounds;
                if (b.Maximum == b.Minimum) continue;
                UpdateMinMax(b, ref min, ref max);
            }
            if (_revitModel != null)
            {
                var b = _revitModel.Bounds;
                if (b.Maximum != b.Minimum)
                    UpdateMinMax(b, ref min, ref max);
            }

            if (min > max) { min = -50; max = 50; }

            // Add 10% padding
            float pad = Math.Max((max - min) * 0.1f, 1f);
            min -= pad; max += pad;

            _sectionMgr.MinBound = min;
            _sectionMgr.MaxBound = max;

            SectionOffsetSlider.Minimum = min;
            SectionOffsetSlider.Maximum = max;
        }

        private static void UpdateMinMax(SharpDX.BoundingBox b, ref float min, ref float max)
        {
            // Choose the relevant component based on current axis selection would be ideal,
            // but using the full extents keeps this simple and always correct.
            min = Math.Min(min, Math.Min(b.Minimum.X, Math.Min(b.Minimum.Y, b.Minimum.Z)));
            max = Math.Max(max, Math.Max(b.Maximum.X, Math.Max(b.Maximum.Y, b.Maximum.Z)));
        }

        // ── Settings ──────────────────────────────────────────────────────────

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            var win = new SettingsWindow(_settings, ApplySettings)
            {
                Owner = this
            };
            win.Show();
        }

        /// <summary>
        /// Push all values from <see cref="_settings"/> into the live sub-systems.
        /// Called once on startup and whenever the settings window reports a change.
        /// </summary>
        private void ApplySettings()
        {
            // Walk controller
            if (_fpController != null)
            {
                _fpController.WalkSpeed        = _settings.WalkSpeed;
                _fpController.SprintMultiplier = _settings.SprintMultiplier;
                _fpController.MouseSensitivity = _settings.MouseSensitivity;
            }

            // Camera FOV
            if (_viewerHost?.Camera != null)
                _viewerHost.Camera.FieldOfView = _settings.FieldOfView;
        }

        // ── Instant scroll-wheel zoom (no inertia) ────────────────────────────

        /// <summary>
        /// Handle scroll-wheel zoom ourselves at the Window level so we can apply
        /// an instant, fixed-step zoom and block Helix's smooth (inertia) zoom
        /// before it ever fires.
        /// </summary>
        private void OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_viewport == null || _viewerHost?.Camera == null) return;
            if (_fpController != null && _fpController.IsActive) return; // walk mode handles its own speed

            var cam = _viewerHost.Camera;

            // Camera-to-look-target distance
            var lookDir = cam.LookDirection;
            double dist = lookDir.Length; // LookDirection magnitude = distance to target

            if (dist < 0.001) dist = 0.001;

            // Fraction to move per notch (delta comes in multiples of 120 for a standard wheel)
            double fraction = _settings.ZoomStep * (e.Delta / 120.0);
            double move     = dist * fraction;

            // Normalise look direction and move position along it
            lookDir.Normalize();
            cam.Position = new System.Windows.Media.Media3D.Point3D(
                cam.Position.X + lookDir.X * move,
                cam.Position.Y + lookDir.Y * move,
                cam.Position.Z + lookDir.Z * move);

            // Block Helix's own zoom handler completely
            e.Handled = true;
        }

        // ── Resize edges ──────────────────────────────────────────────────────
        private void LeftEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => ResizeWindow(HTLEFT);
        private void RightEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => ResizeWindow(HTRIGHT);
        private void BottomEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => ResizeWindow(HTBOTTOM);
        private void BottomLeftCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => ResizeWindow(HTBOTTOMLEFT);
        private void BottomRightCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            => ResizeWindow(HTBOTTOMRIGHT);

        private void ResizeWindow(int direction)
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            ReleaseCapture();
            SendMessage(helper.Handle, WM_NCLBUTTONDOWN, (IntPtr)direction, IntPtr.Zero);
        }
    }
}
