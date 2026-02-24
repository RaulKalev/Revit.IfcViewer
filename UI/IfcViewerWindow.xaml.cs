using Autodesk.Revit.UI;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Ifc;
using IfcViewer.Revit;
using IfcViewer.Viewer;
using Microsoft.Win32;
using SharpDX;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

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
        private GroupModel3D _wireframeRoot;   // hard-edge line overlay
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

        // Stage 4: IFC file watchers + reload banner
        private readonly Dictionary<string, IfcFileWatcher> _fileWatchers
            = new Dictionary<string, IfcFileWatcher>(StringComparer.OrdinalIgnoreCase);
        private IfcModel _pendingReload;

        // Stage 5: Element selection + properties panel
        // Flat maps of all loaded meshes → their extracted element info.
        // IFC map maintained in sync with _loadedModels; Revit map rebuilt on each sync.
        private readonly Dictionary<MeshGeometryModel3D, IfcElementInfo>   _ifcElementMap
            = new Dictionary<MeshGeometryModel3D, IfcElementInfo>();
        private readonly Dictionary<MeshGeometryModel3D, RevitElementInfo> _revitElementMap
            = new Dictionary<MeshGeometryModel3D, RevitElementInfo>();
        private MeshGeometryModel3D  _selectedMesh;
        private Color4               _selectedOriginalEmissive;
        private System.Windows.Point _selectionMouseDown;

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
                    // SSAO — ambient occlusion darkens corners and contact zones, making
                    // object shapes much easier to read in dense architectural models.
                    // Sampling radius is in world units (metres); default is sub-metre
                    // which is invisible at building scale — 1.5 m covers wall/floor corners.
                    EnableSSAO               = true,
                    SSAOSamplingRadius       = 1.5,
                    SSAOIntensity            = 1.5,
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

                // Hook left-click for element selection using PREVIEW (tunnel) events.
                // Bubble events (MouseLeftButtonDown/Up) may be consumed by Helix's
                // CameraController child before they reach our handlers; Preview events
                // fire top-down, reaching the viewport before any child sees them.
                _viewport.PreviewMouseLeftButtonDown += Viewport_MouseLeftButtonDown;
                _viewport.PreviewMouseLeftButtonUp   += Viewport_MouseLeftButtonUp;

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

                // 5b. Wireframe overlay group — lives alongside IfcRoot/RevitRoot.
                //     Populated on demand by RebuildWireframe(); never holds mesh nodes.
                _wireframeRoot = new GroupModel3D();
                _sceneRoot.Children.Add(_wireframeRoot);

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
                RestoreWindowGeometry(_settings);

                // Wire up 3D view picker callbacks now that _settings is loaded.
                // These are called from the Revit API thread via Dispatcher.Invoke
                // when no active 3D view is present and a picker dialog is needed.
                _syncRevitEvent.PickView3DCallback = (viewNames, preselected) =>
                {
                    string result = null;
                    Dispatcher.Invoke((Action)(() =>
                    {
                        var dlg = new View3DPickerDialog(viewNames, preselected) { Owner = this };
                        if (dlg.ShowDialog() == true)
                            result = dlg.SelectedViewName;
                    }));
                    return result;
                };
                _syncRevitEvent.SaveViewCallback = (docPath, viewName) =>
                {
                    _settings.SavedRevit3DViews[docPath] = viewName;
                    _settings.Save();
                };
                _syncRevitEvent.GetSavedViewCallback = docPath =>
                {
                    _settings.SavedRevit3DViews.TryGetValue(docPath, out string name);
                    return name;
                };

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
                // Persist window geometry so it reopens in the same position/size.
                if (_settings != null)
                {
                    _settings.WindowLeft   = Left;
                    _settings.WindowTop    = Top;
                    _settings.WindowWidth  = Width;
                    _settings.WindowHeight = Height;
                    _settings.Save();
                }

                // Stop auto-sync and dispose file watchers before tearing down the scene.
                _syncRevitEvent?.StopAutoSync();
                foreach (var w in _fileWatchers.Values) w.Dispose();
                _fileWatchers.Clear();

                // Clear selection state so no stale material references remain.
                _selectedMesh = null;
                _ifcElementMap.Clear();
                _revitElementMap.Clear();

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

                    // Populate the element selection map with this model's per-element meshes
                    foreach (var kvp in ifcModel.ElementMap)
                        _ifcElementMap[kvp.Key] = kvp.Value;

                    // Register all cross-section meshes with the section plane manager
                    _sectionMgr?.RegisterGroup(ifcModel.SceneGroup);

                    // Rebuild wireframe if the overlay is already enabled by the user
                    if (WireframeToggle?.IsChecked == true)
                        RebuildWireframe();

                    // Update section slider range to cover the full scene
                    UpdateSectionBounds();

                    // Fit camera only on the very first geometry load so the user's
                    // current camera position is preserved when adding additional models.
                    if (_loadedModels.Count == 1 && _revitModel == null)
                        _viewerHost.FitView(ifcModel.Bounds);

                    // Watch the file for external changes (e.g. re-export from Revit)
                    var watcher = new IfcFileWatcher(path,
                        () => Dispatcher.BeginInvoke((Action)(() => ShowReloadBanner(path))));
                    _fileWatchers[path] = watcher;

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

            // Dispose file watcher for this path
            if (_fileWatchers.TryGetValue(selected.FilePath, out var watcher))
            {
                watcher.Dispose();
                _fileWatchers.Remove(selected.FilePath);
            }

            // Clear pending reload if it was for this model
            if (_pendingReload != null &&
                string.Equals(_pendingReload.FilePath, selected.FilePath, StringComparison.OrdinalIgnoreCase))
            {
                _pendingReload = null;
                ReloadBanner.Visibility = Visibility.Collapsed;
            }

            // Clear selection if the selected element belongs to the model being removed
            if (_selectedMesh != null && selected.ElementMap.ContainsKey(_selectedMesh))
                ClearSelection();

            // Remove this model's meshes from the flat element map
            foreach (var mesh in selected.ElementMap.Keys)
                _ifcElementMap.Remove(mesh);

            // Remove from scene
            _sectionMgr?.UnregisterGroup(selected.SceneGroup);
            _viewerHost.IfcRoot.Children.Remove(selected.SceneGroup);
            _loadedModels.Remove(selected);

            // Rebuild wireframe without the removed model's edges
            if (WireframeToggle?.IsChecked == true) RebuildWireframe();

            UpdateStatus(_loadedModels.Count == 0
                ? "GPU Viewport Active"
                : _loadedModels.Count + " model(s) loaded");
            SessionLogger.Info("Removed: " + selected.DisplayName);
        }

        private void SyncRevit_Click(object sender, RoutedEventArgs e)
        {
            if (_syncRevitEvent == null) return;

            SyncRevitButton.IsEnabled = false;
            UpdateStatus("Exporting Revit geometry…");

            _syncRevitEvent.Request(
                onComplete: model =>
                {
                    ApplyRevitUpdate(model, fitCamera: _revitModel == null && _loadedModels.Count == 0);
                    SyncRevitButton.IsEnabled = true;
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

        // ── Auto-sync ─────────────────────────────────────────────────────────

        private void AutoSync_Checked(object sender, RoutedEventArgs e)
        {
            if (_syncRevitEvent == null) return;

            SyncRevitButton.IsEnabled = false;
            UpdateStatus("Auto-sync: initial export…");

            _syncRevitEvent.StartAutoSync(
                onUpdate: model => ApplyRevitUpdate(model,
                    fitCamera: _revitModel == null && _loadedModels.Count == 0),
                onError: ex =>
                {
                    if (AutoSyncToggle != null) AutoSyncToggle.IsChecked = false;
                    SyncRevitButton.IsEnabled = true;
                    UpdateStatus("Auto-sync error — see log.");
                    MessageBox.Show(this,
                        "Auto-sync failed:\n\n" + ex.Message,
                        "Auto-sync Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    SessionLogger.Error("Auto-sync failed.", ex);
                });
        }

        private void AutoSync_Unchecked(object sender, RoutedEventArgs e)
        {
            _syncRevitEvent?.StopAutoSync();
            SyncRevitButton.IsEnabled = true;
            UpdateStatus("Auto-sync stopped.");
            SessionLogger.Info("Auto-sync deactivated.");
        }

        /// <summary>
        /// Applies a new or incrementally-updated <see cref="RevitModel"/> to the scene.
        /// Handles full replacement (new SceneGroup) and in-place incremental patch
        /// (same SceneGroup object mutated by PatchScene).
        /// </summary>
        private void ApplyRevitUpdate(RevitModel model, bool fitCamera = false)
        {
            if (_revitModel != null)
            {
                // Unregister all currently tracked Revit meshes.
                _sectionMgr?.UnregisterGroup(_revitModel.SceneGroup);

                // Remove stale SectionPlaneManager entries for meshes that were removed
                // in-place by PatchScene (they are no longer in SceneGroup.Children).
                _sectionMgr?.PruneDetachedEntries();

                if (!ReferenceEquals(_revitModel.SceneGroup, model.SceneGroup))
                {
                    // Full replacement — swap the SceneGroup in RevitRoot.
                    _viewerHost.RevitRoot.Children.Remove(_revitModel.SceneGroup);
                    _viewerHost.RevitRoot.Children.Add(model.SceneGroup);
                }
                // else: incremental — same SceneGroup mutated in-place by PatchScene; already live.
            }
            else
            {
                // First sync — add to scene.
                _viewerHost.RevitRoot.Children.Add(model.SceneGroup);
            }

            _revitModel = model;
            _sectionMgr?.RegisterGroup(_revitModel.SceneGroup);
            UpdateSectionBounds();

            // Rebuild the mesh → info reverse-map for hit-testing.
            // If the currently selected mesh was a Revit element that no longer
            // exists after an incremental patch, clear the selection.
            _revitElementMap.Clear();
            foreach (var kv in model.ElementMeshes)
            {
                if (model.ElementInfos.TryGetValue(kv.Key, out var info))
                    _revitElementMap[kv.Value] = info;
            }
            if (_selectedMesh != null && !_ifcElementMap.ContainsKey(_selectedMesh)
                                      && !_revitElementMap.ContainsKey(_selectedMesh))
                ClearSelection();

            if (fitCamera) _viewerHost.FitView(model.Bounds);

            // Rebuild wireframe if the overlay is already enabled by the user
            if (WireframeToggle?.IsChecked == true)
                RebuildWireframe();

            UpdateStatus($"Revit: {model.DisplayName}  |  {model.MeshCount} elements  |  {model.TriangleCount} triangles");
            SessionLogger.Info($"Revit update applied: {model.TriangleCount} triangles");
        }

        // ── IFC reload banner ─────────────────────────────────────────────────

        private void ShowReloadBanner(string path)
        {
            // Find the model currently loaded from this path
            IfcModel target = null;
            foreach (var m in _loadedModels)
                if (string.Equals(m.FilePath, path, StringComparison.OrdinalIgnoreCase))
                { target = m; break; }
            if (target == null) return;

            _pendingReload          = target;
            ReloadBannerText.Text   = System.IO.Path.GetFileName(path) + " changed on disk";
            ReloadBanner.Visibility = Visibility.Visible;
        }

        private async void ReloadBannerButton_Click(object sender, RoutedEventArgs e)
        {
            ReloadBanner.Visibility = Visibility.Collapsed;
            var model = _pendingReload;
            _pendingReload = null;
            if (model == null) return;

            string path = model.FilePath;

            // Clear selection if the selected element belongs to the model being reloaded
            if (_selectedMesh != null && model.ElementMap.ContainsKey(_selectedMesh))
                ClearSelection();

            // Remove old model's meshes from the flat element map
            foreach (var mesh in model.ElementMap.Keys)
                _ifcElementMap.Remove(mesh);

            // Remove the old model from the scene
            _sectionMgr?.UnregisterGroup(model.SceneGroup);
            _viewerHost.IfcRoot.Children.Remove(model.SceneGroup);
            _loadedModels.Remove(model);

            // Dispose and recreate the watcher (file may have moved/been recreated)
            if (_fileWatchers.TryGetValue(path, out var oldWatcher))
            {
                oldWatcher.Dispose();
                _fileWatchers.Remove(path);
            }

            AddIfcButton.IsEnabled    = false;
            RemoveIfcButton.IsEnabled = false;
            UpdateStatus("Reloading: " + System.IO.Path.GetFileName(path) + " …");

            try
            {
                IfcModel newModel = await IfcLoader.LoadAsync(path, Dispatcher,
                    onProgress: msg => UpdateStatus(msg));

                _viewerHost.IfcRoot.Children.Add(newModel.SceneGroup);
                _loadedModels.Add(newModel);
                ModelListBox.SelectedItem = newModel;

                // Populate element map with the freshly loaded model's meshes
                foreach (var kvp in newModel.ElementMap)
                    _ifcElementMap[kvp.Key] = kvp.Value;

                _sectionMgr?.RegisterGroup(newModel.SceneGroup);
                UpdateSectionBounds();
                // Don't move the camera on reload — user keeps their current view position.

                // Rebuild wireframe for the freshly reloaded geometry
                if (WireframeToggle?.IsChecked == true)
                    RebuildWireframe();

                // Restart the watcher for the reloaded path
                var newWatcher = new IfcFileWatcher(path,
                    () => Dispatcher.BeginInvoke((Action)(() => ShowReloadBanner(path))));
                _fileWatchers[path] = newWatcher;

                UpdateStatus(newModel.DisplayName + "  |  " +
                             newModel.MeshCount + " elements  |  " +
                             newModel.TriangleCount + " triangles");
                SessionLogger.Info("Reloaded: " + newModel.DisplayName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Failed to reload IFC file:\n\n" + ex.Message,
                    "Reload Error", MessageBoxButton.OK, MessageBoxImage.Error);
                SessionLogger.Error("IFC reload failed: " + path, ex);
                UpdateStatus("Reload failed — see log.");
            }
            finally
            {
                AddIfcButton.IsEnabled    = true;
                RemoveIfcButton.IsEnabled = true;
            }
        }

        private void ReloadBannerDismiss_Click(object sender, RoutedEventArgs e)
        {
            ReloadBanner.Visibility = Visibility.Collapsed;
            _pendingReload = null;
        }

        // ── Element selection ─────────────────────────────────────────────────

        private void Viewport_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Record the down position so MouseLeftButtonUp can distinguish a click
            // (small movement) from a camera-orbit drag (large movement).
            _selectionMouseDown = e.GetPosition(_viewport);
        }

        private void Viewport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // Selection is allowed in both orbit mode and walk mode.
            // In walk mode right-drag drives mouse-look (right button only), so
            // left-click is free to hit-test and select elements.

            // Reject drags: only treat as a click if the mouse moved ≤5 px.
            var pos = e.GetPosition(_viewport);
            double dx = pos.X - _selectionMouseDown.X;
            double dy = pos.Y - _selectionMouseDown.Y;
            if (dx * dx + dy * dy > 25) return;

            // Hit-test the 3D scene at the click position.
            var hits = _viewport?.FindHits(pos);

            if (hits == null || hits.Count == 0) { ClearSelection(); return; }

            foreach (var hit in hits)
            {
                var mesh = hit.ModelHit as MeshGeometryModel3D;
                if (mesh == null) continue;

                if (_ifcElementMap.TryGetValue(mesh, out IfcElementInfo ifcInfo))
                {
                    SelectElement(mesh, ifcInfo);
                    return;
                }
                if (_revitElementMap.TryGetValue(mesh, out RevitElementInfo revitInfo))
                {
                    SelectElement(mesh, revitInfo);
                    return;
                }
            }

            ClearSelection();
        }

        /// <summary>Shared highlight logic — teal emissive glow on <paramref name="mesh"/>.</summary>
        private void HighlightMesh(MeshGeometryModel3D mesh)
        {
            if (_selectedMesh != null && _selectedMesh.Material is PhongMaterial prev)
                prev.EmissiveColor = _selectedOriginalEmissive;

            _selectedMesh = mesh;

            if (mesh.Material is PhongMaterial mat)
            {
                _selectedOriginalEmissive = mat.EmissiveColor;
                mat.EmissiveColor = new Color4(0.08f, 0.42f, 0.42f, 1f);
            }
        }

        private void SelectElement(MeshGeometryModel3D mesh, IfcElementInfo info)
        {
            HighlightMesh(mesh);
            ShowElementProperties(info);
            SessionLogger.Info($"Selected IFC: {info.Type} \"{info.Name}\"");
        }

        private void SelectElement(MeshGeometryModel3D mesh, RevitElementInfo info)
        {
            HighlightMesh(mesh);
            ShowElementProperties(info);
            SessionLogger.Info($"Selected Revit: {info.Category} \"{info.Name}\"");
        }

        /// <summary>Removes the current highlight and clears the properties panel.</summary>
        private void ClearSelection()
        {
            if (_selectedMesh != null && _selectedMesh.Material is PhongMaterial mat)
                mat.EmissiveColor = _selectedOriginalEmissive;

            _selectedMesh = null;
            HideProperties();
        }

        /// <summary>
        /// Populates the right-hand properties panel for an IFC element.
        /// Pass <c>null</c> to return to the "click an element" empty state.
        /// </summary>
        private void ShowElementProperties(IfcElementInfo info)
        {
            if (info == null) { HideProperties(); return; }

            ShowProperties(
                name:       string.IsNullOrWhiteSpace(info.Name) ? "(unnamed)" : info.Name,
                typeLine:   info.Type,
                idLine:     info.GlobalId,
                propertySets: info.PropertySets);
        }

        /// <summary>Populates the right-hand properties panel for a Revit element.</summary>
        private void ShowElementProperties(RevitElementInfo info)
        {
            if (info == null) { HideProperties(); return; }

            string typeLine = string.IsNullOrWhiteSpace(info.FamilyName)
                ? info.Category
                : $"{info.Category}  •  {info.FamilyName} : {info.TypeName}";

            ShowProperties(
                name:         string.IsNullOrWhiteSpace(info.Name) ? "(unnamed)" : info.Name,
                typeLine:     typeLine,
                idLine:       $"ElementId {info.ElementId}",
                propertySets: info.PropertySets);
        }

        private void HideProperties()
        {
            PropEmptyHint.Visibility  = Visibility.Visible;
            PropContent.Visibility    = Visibility.Collapsed;
            PropTabContent.Visibility = Visibility.Collapsed;
        }

        private void ShowProperties(string name, string typeLine, string idLine,
                                    Dictionary<string, Dictionary<string, string>> propertySets)
        {
            PropEmptyHint.Visibility  = Visibility.Collapsed;
            PropContent.Visibility    = Visibility.Visible;
            PropTabContent.Visibility = Visibility.Visible;

            PropName.Text     = name;
            PropType.Text     = typeLine;
            PropGlobalId.Text = idLine;

            PropTabBar.Children.Clear();
            PropTabOverflowPanel.Children.Clear();
            PropPropsPanel.Children.Clear();
            PropTabOverflowPopup.IsOpen = false;

            if (propertySets.Count == 0)
            {
                PropPropsPanel.Children.Add(new TextBlock
                {
                    Text       = "No property sets found.",
                    FontSize   = 10,
                    Foreground = (Brush)FindResource("IconBrush"),
                    Margin     = new Thickness(0, 4, 0, 0),
                });
                return;
            }

            // Build one tab + one overflow item per property set.
            var psets = propertySets.ToList();
            Border firstTab = null;
            foreach (var pset in psets)
            {
                var capturedProps = pset.Value;

                // ── Tab in the tab bar ────────────────────────────────────────
                var tabLabel = new TextBlock
                {
                    Text              = pset.Key,
                    FontSize          = 10,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                var tab = new Border
                {
                    Padding         = new Thickness(10, 7, 10, 5),
                    BorderThickness = new Thickness(0, 0, 0, 2),
                    BorderBrush     = Brushes.Transparent,
                    Background      = Brushes.Transparent,
                    Cursor          = Cursors.Hand,
                    Child           = tabLabel,
                };
                tab.MouseLeftButtonDown += (_, __) => SelectPropTab(tab, capturedProps);
                PropTabBar.Children.Add(tab);
                if (firstTab == null) firstTab = tab;

                // ── Matching row in the overflow dropdown ─────────────────────
                var capturedTab = tab;
                var itemLabel = new TextBlock
                {
                    Text              = pset.Key,
                    FontSize          = 10,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground        = (Brush)FindResource("ForegroundBrush"),
                };
                var item = new Border
                {
                    Padding    = new Thickness(14, 7, 14, 7),
                    Background = Brushes.Transparent,
                    Cursor     = Cursors.Hand,
                    Child      = itemLabel,
                };
                item.MouseEnter += (_, __) =>
                    item.Background = new SolidColorBrush(
                        System.Windows.Media.Color.FromArgb(30, 255, 255, 255));
                item.MouseLeave += (_, __) =>
                    item.Background = Brushes.Transparent;
                item.MouseLeftButtonDown += (_, __) =>
                {
                    SelectPropTab(capturedTab, capturedProps);
                    capturedTab.BringIntoView();
                    PropTabOverflowPopup.IsOpen = false;
                };
                PropTabOverflowPanel.Children.Add(item);
            }

            // Activate the first tab by default.
            if (firstTab != null)
                SelectPropTab(firstTab, psets[0].Value);
        }

        private void PropTabOverflow_Click(object sender, MouseButtonEventArgs e)
        {
            PropTabOverflowPopup.IsOpen = !PropTabOverflowPopup.IsOpen;
        }

        /// <summary>
        /// Highlights the active tab with a teal underline and rebuilds
        /// <see cref="PropPropsPanel"/> with that tab's properties.
        /// </summary>
        private void SelectPropTab(Border activeTab, Dictionary<string, string> props)
        {
            var teal = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x70, 0xBA, 0xBC));
            var dim  = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xAA, 0xAA, 0xAA));

            foreach (Border tab in PropTabBar.Children)
            {
                bool active = tab == activeTab;
                tab.BorderBrush = active ? teal : Brushes.Transparent;
                if (tab.Child is TextBlock tb)
                    tb.Foreground = active ? teal : dim;
            }

            // Scroll the tab bar so the active tab is visible.
            activeTab.BringIntoView();

            PropPropsPanel.Children.Clear();

            foreach (var prop in props)
            {
                var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition
                    { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition
                    { Width = new GridLength(1.2, GridUnitType.Star) });

                var nameBlock = new TextBlock
                {
                    Text         = prop.Key,
                    FontSize     = 10,
                    Foreground   = (Brush)FindResource("IconBrush"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip      = prop.Key,
                };
                var valBlock = new TextBlock
                {
                    Text         = prop.Value,
                    FontSize     = 10,
                    Foreground   = (Brush)FindResource("ForegroundBrush"),
                    TextWrapping = TextWrapping.Wrap,
                };

                Grid.SetColumn(nameBlock, 0);
                Grid.SetColumn(valBlock,  1);
                row.Children.Add(nameBlock);
                row.Children.Add(valBlock);
                PropPropsPanel.Children.Add(row);
            }
        }

        private void ResetCamera_Click(object sender, RoutedEventArgs e)
            => _viewerHost?.ResetCamera();

        private void Wireframe_Checked(object sender, RoutedEventArgs e)
            => RebuildWireframe();

        private void Wireframe_Unchecked(object sender, RoutedEventArgs e)
        {
            _wireframeRoot?.Children.Clear();
            SessionLogger.Info("Wireframe off.");
        }

        /// <summary>
        /// Rebuilds the hard-edge wireframe overlay for all geometry currently in the
        /// scene. Edge extraction runs on a thread-pool thread to avoid freezing the UI
        /// for large models; the overlay is populated back on the dispatcher thread.
        /// </summary>
        private void RebuildWireframe()
        {
            if (_wireframeRoot == null) return;
            _wireframeRoot.Children.Clear();
            if (WireframeToggle?.IsChecked != true) return;

            // Snapshot geometry references on the UI thread before going async —
            // Vector3Collection / IntCollection are WPF observable collections that
            // must not be enumerated concurrently with mutations.
            var meshList = new List<MeshGeometry3D>();
            CollectMeshGeometries(_sceneRoot, meshList);

            if (meshList.Count == 0) return;
            SessionLogger.Info($"Wireframe on — extracting hard edges from {meshList.Count} mesh(es).");

            Task.Run(() =>
            {
                var lineGeoms = new List<LineGeometry3D>(meshList.Count);
                foreach (var mg in meshList)
                {
                    var lg = WireframeHelper.ExtractHardEdges(mg);
                    if (lg != null) lineGeoms.Add(lg);
                }
                return lineGeoms;
            }).ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    SessionLogger.Error("Wireframe edge extraction failed.", t.Exception?.GetBaseException());
                    return;
                }
                // Guard: user may have toggled off while we were computing
                if (WireframeToggle?.IsChecked != true) return;

                _wireframeRoot.Children.Clear();
                foreach (var lg in t.Result)
                {
                    _wireframeRoot.Children.Add(new LineGeometryModel3D
                    {
                        Geometry  = lg,
                        // Near-black (#222222) contrasts against any mesh surface colour.
                        Color     = System.Windows.Media.Color.FromRgb(0x22, 0x22, 0x22),
                        Thickness = 1.5,
                    });
                }
                SessionLogger.Info($"Wireframe overlay: {t.Result.Count} line object(s).");
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        /// <summary>
        /// Recursively collects all <see cref="MeshGeometry3D"/> geometry objects from
        /// the scene hierarchy, skipping <see cref="_wireframeRoot"/> (line nodes only).
        /// </summary>
        private void CollectMeshGeometries(GroupModel3D group, List<MeshGeometry3D> results)
        {
            foreach (var child in group.Children)
            {
                if (ReferenceEquals(child, _wireframeRoot)) continue;
                if (child is MeshGeometryModel3D m && m.Geometry is MeshGeometry3D mg)
                    results.Add(mg);
                else if (child is GroupModel3D sub)
                    CollectMeshGeometries(sub, results);
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
        /// Restore the window to its last saved position and size.
        /// Clamps to the virtual screen so the window is never off-screen.
        /// </summary>
        private void RestoreWindowGeometry(ViewerSettings s)
        {
            if (s.WindowWidth == null || s.WindowHeight == null) return;

            double screenW = SystemParameters.VirtualScreenWidth;
            double screenH = SystemParameters.VirtualScreenHeight;
            double screenX = SystemParameters.VirtualScreenLeft;
            double screenY = SystemParameters.VirtualScreenTop;

            double w = Math.Max(MinWidth,  Math.Min(s.WindowWidth.Value,  screenW));
            double h = Math.Max(MinHeight, Math.Min(s.WindowHeight.Value, screenH));
            double l = s.WindowLeft ?? ((screenW - w) / 2 + screenX);
            double t = s.WindowTop  ?? ((screenH - h) / 2 + screenY);

            // Clamp so at least the title bar (top 40px) stays on screen.
            l = Math.Max(screenX, Math.Min(l, screenX + screenW - w));
            t = Math.Max(screenY, Math.Min(t, screenY + screenH - 40));

            Left   = l;
            Top    = t;
            Width  = w;
            Height = h;
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
