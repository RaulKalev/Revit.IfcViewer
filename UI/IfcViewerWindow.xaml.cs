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
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Media3D = System.Windows.Media.Media3D;
using WpfColor  = System.Windows.Media.Color;
using WpfPoint  = System.Windows.Point;
using WpfShapes = System.Windows.Shapes;

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
        private readonly ObservableCollection<IfcModelListItem> _ifcCatalogItems
            = new ObservableCollection<IfcModelListItem>();
        private readonly Dictionary<string, CachedIfcModel> _ifcModelCache
            = new Dictionary<string, CachedIfcModel>(StringComparer.OrdinalIgnoreCase);
        private string _activeIfcFolder;

        // CancellationToken support — cancel previous load when starting a new one
        private CancellationTokenSource _loadCts;

        // Stage 3: Revit sync glue
        private SyncRevitEvent _syncRevitEvent;
        private RevitModel     _revitModel;

        // Stage 4a: first-person controller + section plane + settings
        private FirstPersonController _fpController;
        private SectionPlaneManager   _sectionMgr;
        private ViewerSettings        _settings;
        private ViewerFocusService    _viewerFocusService;
        private FollowSelectionService _followSelectionService;
        private bool _applyingFollowSelectionState;
        private readonly Dictionary<string, MeshGeometryModel3D> _ifcGuidMeshMap
            = new Dictionary<string, MeshGeometryModel3D>(StringComparer.OrdinalIgnoreCase);
        private static readonly string[] IfcGuidParameterNames = new[]
        {
            "IfcGUID",
            "IFC GUID",
            "IFC_GUID",
            "IFCGUID",
            "GlobalId",
            "Global ID",
            "IfcGlobalId",
            "IFC GlobalId",
        };

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
        private readonly Stack<MeshGeometryModel3D> _hiddenMeshes = new Stack<MeshGeometryModel3D>();



        // ── ViewCube compass ring ──────────────────────────────────────────────
        // WPF 2D overlay wrapped around the Helix built-in ViewCube.
        // Cube and compass are projected from the same 3D basis so they rotate
        // together with the full camera orientation (yaw + pitch + roll).
        private WpfShapes.Path _compassRing;
        private readonly Dictionary<string, TextBlock> _compassLabels
            = new Dictionary<string, TextBlock>(StringComparer.Ordinal);
        private enum OrientationCubeFace { Front, Back, Left, Right, Top, Bottom }
        private enum OrientationArrowDirection { Left, Right, Up, Down }
        private struct OrientationCubeEdgeSpec
        {
            public readonly int V0;
            public readonly int V1;
            public readonly OrientationCubeFace FaceA;
            public readonly OrientationCubeFace FaceB;

            public OrientationCubeEdgeSpec(int v0, int v1, OrientationCubeFace faceA, OrientationCubeFace faceB)
            {
                V0 = v0;
                V1 = v1;
                FaceA = faceA;
                FaceB = faceB;
            }
        }
        private struct OrientationCubeCornerSpec
        {
            public readonly int Vertex;
            public readonly OrientationCubeFace FaceA;
            public readonly OrientationCubeFace FaceB;
            public readonly OrientationCubeFace FaceC;

            public OrientationCubeCornerSpec(
                int vertex,
                OrientationCubeFace faceA,
                OrientationCubeFace faceB,
                OrientationCubeFace faceC)
            {
                Vertex = vertex;
                FaceA = faceA;
                FaceB = faceB;
                FaceC = faceC;
            }
        }
        private readonly Dictionary<OrientationCubeFace, WpfShapes.Polygon> _orientationCubeFaces
            = new Dictionary<OrientationCubeFace, WpfShapes.Polygon>();
        private readonly Dictionary<OrientationCubeFace, TextBlock> _orientationCubeFaceLabels
            = new Dictionary<OrientationCubeFace, TextBlock>();
        private readonly Dictionary<OrientationCubeFace, SolidColorBrush> _orientationCubeFaceBrushes
            = new Dictionary<OrientationCubeFace, SolidColorBrush>();
        private readonly Dictionary<OrientationCubeFace, SolidColorBrush> _orientationCubeFaceHoverBrushes
            = new Dictionary<OrientationCubeFace, SolidColorBrush>();
        private readonly Dictionary<int, WpfShapes.Polygon> _orientationCubeEdgeHotspots
            = new Dictionary<int, WpfShapes.Polygon>();
        private readonly Dictionary<int, WpfShapes.Ellipse> _orientationCubeCornerHotspots
            = new Dictionary<int, WpfShapes.Ellipse>();
        private readonly Dictionary<OrientationArrowDirection, WpfShapes.Polygon> _orientationCubeArrowButtons
            = new Dictionary<OrientationArrowDirection, WpfShapes.Polygon>();
        private readonly Dictionary<OrientationArrowDirection, OrientationCubeFace> _orientationCubeArrowTargets
            = new Dictionary<OrientationArrowDirection, OrientationCubeFace>();
        private readonly SolidColorBrush _orientationArrowFillBrush
            = new SolidColorBrush(WpfColor.FromArgb(230, 205, 205, 205));
        private readonly SolidColorBrush _orientationArrowHoverFillBrush
            = new SolidColorBrush(WpfColor.FromArgb(245, 170, 170, 170));
        private readonly SolidColorBrush _orientationEdgeHoverFillBrush
            = new SolidColorBrush(WpfColor.FromArgb(110, 142, 196, 255));
        private readonly SolidColorBrush _orientationEdgeHoverStrokeBrush
            = new SolidColorBrush(WpfColor.FromArgb(235, 90, 153, 232));
        private readonly SolidColorBrush _orientationCornerHoverFillBrush
            = new SolidColorBrush(WpfColor.FromArgb(145, 142, 196, 255));
        private readonly SolidColorBrush _orientationCornerHoverStrokeBrush
            = new SolidColorBrush(WpfColor.FromArgb(245, 90, 153, 232));
        private OrientationCubeFace? _hoveredOrientationCubeFace;
        private OrientationArrowDirection? _hoveredOrientationArrow;
        private int? _hoveredOrientationEdgeIndex;
        private int? _hoveredOrientationCornerIndex;
        private static readonly OrientationCubeFace[] OrientationCubeFaceOrder = new[]
        {
            OrientationCubeFace.Front,
            OrientationCubeFace.Back,
            OrientationCubeFace.Left,
            OrientationCubeFace.Right,
            OrientationCubeFace.Top,
            OrientationCubeFace.Bottom,
        };
        private static readonly OrientationArrowDirection[] OrientationArrowOrder = new[]
        {
            OrientationArrowDirection.Left,
            OrientationArrowDirection.Right,
            OrientationArrowDirection.Up,
            OrientationArrowDirection.Down,
        };
        private static readonly Media3D.Point3D[] OrientationCubeVertices = new[]
        {
            new Media3D.Point3D(-1, -1, -1), // 0
            new Media3D.Point3D( 1, -1, -1), // 1
            new Media3D.Point3D( 1,  1, -1), // 2
            new Media3D.Point3D(-1,  1, -1), // 3
            new Media3D.Point3D(-1, -1,  1), // 4
            new Media3D.Point3D( 1, -1,  1), // 5
            new Media3D.Point3D( 1,  1,  1), // 6
            new Media3D.Point3D(-1,  1,  1), // 7
        };
        private static readonly int[] OrientationFaceFront  = { 0, 1, 2, 3 };
        private static readonly int[] OrientationFaceBack   = { 4, 5, 6, 7 };
        private static readonly int[] OrientationFaceLeft   = { 0, 4, 7, 3 };
        private static readonly int[] OrientationFaceRight  = { 1, 5, 6, 2 };
        private static readonly int[] OrientationFaceTop    = { 3, 2, 6, 7 };
        private static readonly int[] OrientationFaceBottom = { 0, 1, 5, 4 };
        private static readonly OrientationCubeEdgeSpec[] OrientationCubeEdges = new[]
        {
            new OrientationCubeEdgeSpec(0, 1, OrientationCubeFace.Front, OrientationCubeFace.Bottom),
            new OrientationCubeEdgeSpec(1, 2, OrientationCubeFace.Front, OrientationCubeFace.Right),
            new OrientationCubeEdgeSpec(2, 3, OrientationCubeFace.Front, OrientationCubeFace.Top),
            new OrientationCubeEdgeSpec(3, 0, OrientationCubeFace.Front, OrientationCubeFace.Left),
            new OrientationCubeEdgeSpec(4, 5, OrientationCubeFace.Back,  OrientationCubeFace.Bottom),
            new OrientationCubeEdgeSpec(5, 6, OrientationCubeFace.Back,  OrientationCubeFace.Right),
            new OrientationCubeEdgeSpec(6, 7, OrientationCubeFace.Back,  OrientationCubeFace.Top),
            new OrientationCubeEdgeSpec(7, 4, OrientationCubeFace.Back,  OrientationCubeFace.Left),
            new OrientationCubeEdgeSpec(0, 4, OrientationCubeFace.Left,  OrientationCubeFace.Bottom),
            new OrientationCubeEdgeSpec(1, 5, OrientationCubeFace.Right, OrientationCubeFace.Bottom),
            new OrientationCubeEdgeSpec(2, 6, OrientationCubeFace.Right, OrientationCubeFace.Top),
            new OrientationCubeEdgeSpec(3, 7, OrientationCubeFace.Left,  OrientationCubeFace.Top),
        };
        private static readonly OrientationCubeCornerSpec[] OrientationCubeCorners = new[]
        {
            new OrientationCubeCornerSpec(0, OrientationCubeFace.Left,  OrientationCubeFace.Bottom, OrientationCubeFace.Front),
            new OrientationCubeCornerSpec(1, OrientationCubeFace.Right, OrientationCubeFace.Bottom, OrientationCubeFace.Front),
            new OrientationCubeCornerSpec(2, OrientationCubeFace.Right, OrientationCubeFace.Top,    OrientationCubeFace.Front),
            new OrientationCubeCornerSpec(3, OrientationCubeFace.Left,  OrientationCubeFace.Top,    OrientationCubeFace.Front),
            new OrientationCubeCornerSpec(4, OrientationCubeFace.Left,  OrientationCubeFace.Bottom, OrientationCubeFace.Back),
            new OrientationCubeCornerSpec(5, OrientationCubeFace.Right, OrientationCubeFace.Bottom, OrientationCubeFace.Back),
            new OrientationCubeCornerSpec(6, OrientationCubeFace.Right, OrientationCubeFace.Top,    OrientationCubeFace.Back),
            new OrientationCubeCornerSpec(7, OrientationCubeFace.Left,  OrientationCubeFace.Top,    OrientationCubeFace.Back),
        };
        private const double OrientationCubeScale = 14.0;
        private double _orientationCubeCenterX;
        private double _orientationCubeCenterY;

        private DependencyPropertyDescriptor _cameraLookDirectionDescriptor;
        private DependencyPropertyDescriptor _cameraUpDirectionDescriptor;
        private DependencyPropertyDescriptor _cameraPositionDescriptor;
        private EventHandler _cameraOrientationChangedHandler;
        private bool _cameraSyncAttached;

        private Media3D.Point3D  _lastCameraPosition;
        private Media3D.Vector3D _lastCameraLookDirection;
        private Media3D.Vector3D _lastCameraUpDirection;
        private bool _hasCameraSnapshot;

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
                    InfoBackground           = Brushes.Transparent,
                    TitleBackground          = Brushes.Transparent,
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
                    // ── Built-in ViewCube ──────────────────────────────────────
                    // Custom Revit-style texture applied below after construction.
                    // Position: top-right (HorizontalPosition + VerticalPosition are
                    // normalised device coords: +x = right, +y = top).
                    ShowViewCube               = true,
                    ViewCubeSize               = 80,
                    ViewCubeHorizontalPosition = 0.75,
                    ViewCubeVerticalPosition   = 0.90,
                    IsViewCubeEdgeClicksEnabled = true,
                    IsViewCubeMoverEnabled      = false,
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

                // 5b. Wireframe overlay group — user-toggled hard-edge lines.
                _wireframeRoot = new GroupModel3D();
                _sceneRoot.Children.Add(_wireframeRoot);

                // 5c. Always-on outline removed — scene renders clean shaded meshes.

                // 5d. ViewCube: apply Revit-style face texture and add compass ring.
                _viewport.ViewCubeTexture = CreateRevitViewCubeTexture();
                ViewportContainer.Children.Add(BuildCompassOverlay());

                // Keep orientation overlays synced with camera motion (including
                // animated transitions) and force redraw so the ViewCube updates.
                AttachCameraOrientationSync();
                UpdateCameraOrientationWidgets(force: true);

                // 6. Bind the IFC catalog list
                ModelListBox.ItemsSource = _ifcCatalogItems;

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
                    if (ev.Key == Key.Space && TryHandleSpacebar())
                    {
                        ev.Handled = true;
                        return;
                    }

                    if (_fpController != null && !_fpController.IsActive && IsMovementKey(ev.Key))
                        ev.Handled = true;
                };

                // 9. Section plane manager + attach its visual quad to the scene root
                _sectionMgr = new SectionPlaneManager();
                _sectionMgr.AttachVisual(_sceneRoot);

                // 10. Follow-selection services
                _viewerFocusService = new ViewerFocusService(_viewerHost.Camera);
                _followSelectionService = new FollowSelectionService(
                    _uiApp, OnRevitPrimarySelectionChanged);

                // 11. Load settings from disk (or defaults) and apply them
                _settings = ViewerSettings.Load();
                ApplySettings();
                RestoreWindowGeometry(_settings);
                LoadSavedIfcFolderForCurrentProject();

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

                // Cancel any in-flight IFC load
                _loadCts?.Cancel();
                _loadCts?.Dispose();
                _loadCts = null;

                // Stop auto-sync and dispose file watchers before tearing down the scene.
                _syncRevitEvent?.StopAutoSync();
                foreach (var w in _fileWatchers.Values) w.Dispose();
                _fileWatchers.Clear();

                // Clear selection state so no stale material references remain.
                _selectedMesh = null;
                _ifcElementMap.Clear();
                _hiddenMeshes.Clear();
                _revitElementMap.Clear();
                _ifcGuidMeshMap.Clear();
                _ifcCatalogItems.Clear();
                _ifcModelCache.Clear();

                this.PreviewMouseWheel -= OnPreviewMouseWheel;
                DetachCameraOrientationSync();
                _followSelectionService?.Dispose();
                _fpController?.Dispose();
                _sectionMgr?.DetachVisual();
                _syncRevitEvent?.Dispose();
                _sceneRoot?.Children.Clear();
                _viewport?.Items.Clear();
                _viewerHost?.Dispose();
                _viewerFocusService = null;
                _followSelectionService = null;
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
        private void AddIfc_Click(object sender, RoutedEventArgs e)
        {
            string folderPath;
            if (!TrySelectIfcFolder(out folderPath)) return;
            SetIfcFolder(folderPath, persistForProject: true);
        }

        private async void ModelListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = ModelListBox.SelectedItem as IfcModelListItem;
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath)) return;
            if (item.IsLoaded) return;

            await LoadIfcModelByPathAsync(item.FilePath);
        }

        private void RemoveIfc_Click(object sender, RoutedEventArgs e)
        {
            var item = ModelListBox.SelectedItem as IfcModelListItem;
            if (item == null || !item.IsLoaded) return;

            IfcModel selected = FindLoadedModel(item.FilePath);
            if (selected == null)
            {
                item.IsLoaded = false;
                return;
            }

            UnloadIfcModel(selected, removeFromCache: false);
            SessionLogger.Info("Removed: " + selected.DisplayName);
        }

        // ── Model list context menu ───────────────────────────────────────────

        /// <summary>
        /// Selects the right-clicked list item so SelectedItem is correct
        /// when the context menu Click handlers fire.
        /// </summary>
        private void ModelListItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var listBoxItem = sender as ListBoxItem;
            if (listBoxItem != null)
                listBoxItem.IsSelected = true;
        }

        private async void ModelContextMenu_Load_Click(object sender, RoutedEventArgs e)
        {
            var item = ModelListBox.SelectedItem as IfcModelListItem;
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath)) return;
            if (item.IsLoaded) return;

            await LoadIfcModelByPathAsync(item.FilePath);
        }

        private async void ModelContextMenu_Reload_Click(object sender, RoutedEventArgs e)
        {
            var item = ModelListBox.SelectedItem as IfcModelListItem;
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath)) return;

            string path = item.FilePath;

            // Unload from scene if currently loaded
            IfcModel loaded = FindLoadedModel(path);
            if (loaded != null)
                UnloadIfcModel(loaded, removeFromCache: true, updateStatus: false);

            // Purge in-memory cache entry and disk cache so next load rebuilds fully
            _ifcModelCache.Remove(path);
            IfcLoader.InvalidateCache(path);

            SessionLogger.Info("Reloading (cache cleared): " + Path.GetFileNameWithoutExtension(path));
            await LoadIfcModelByPathAsync(path);
        }

        private void ModelContextMenu_Remove_Click(object sender, RoutedEventArgs e)
        {
            var item = ModelListBox.SelectedItem as IfcModelListItem;
            if (item == null || !item.IsLoaded) return;

            IfcModel selected = FindLoadedModel(item.FilePath);
            if (selected == null)
            {
                item.IsLoaded = false;
                return;
            }

            UnloadIfcModel(selected, removeFromCache: false);
            SessionLogger.Info("Removed: " + selected.DisplayName);
        }

        private bool TrySelectIfcFolder(out string folderPath)
        {
            folderPath = null;

            var dlg = new OpenFileDialog
            {
                Title = "Select IFC folder",
                Filter = "Folder selection|*.folder",
                FileName = "Select this folder",
                CheckFileExists = false,
                CheckPathExists = true,
                ValidateNames = false,
                Multiselect = false,
            };

            if (dlg.ShowDialog(this) != true) return false;

            string candidate = null;
            try { candidate = Path.GetDirectoryName(dlg.FileName); }
            catch { }

            if (string.IsNullOrWhiteSpace(candidate))
                candidate = dlg.FileName;

            if (string.IsNullOrWhiteSpace(candidate) || !Directory.Exists(candidate))
            {
                MessageBox.Show(this,
                    "Invalid folder selection.",
                    "Folder Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            folderPath = candidate;
            return true;
        }

        private void SetIfcFolder(string folderPath, bool persistForProject)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return;

            _activeIfcFolder = folderPath;

            if (persistForProject && _settings != null)
            {
                string projectKey = GetCurrentProjectKey();
                if (!string.IsNullOrWhiteSpace(projectKey))
                {
                    _settings.IfcFoldersByProject[projectKey] = folderPath;
                    _settings.Save();
                }
            }

            RefreshIfcCatalog(folderPath);

            if (_ifcCatalogItems.Count == 0)
                UpdateStatus("No IFC files found in folder.");
            else
                UpdateStatus("Folder loaded — double-click a model to load it.");

            SessionLogger.Info("IFC folder selected: " + folderPath);
        }

        private void LoadSavedIfcFolderForCurrentProject()
        {
            if (_settings == null || _settings.IfcFoldersByProject == null) return;

            string projectKey = GetCurrentProjectKey();
            if (string.IsNullOrWhiteSpace(projectKey)) return;

            string folderPath;
            if (!_settings.IfcFoldersByProject.TryGetValue(projectKey, out folderPath))
                return;

            if (!Directory.Exists(folderPath))
            {
                SessionLogger.Warn("Saved IFC folder not found: " + folderPath);
                return;
            }

            SetIfcFolder(folderPath, persistForProject: false);
        }

        private string GetCurrentProjectKey()
        {
            var doc = _uiApp?.ActiveUIDocument?.Document;
            if (doc == null) return "NO_ACTIVE_PROJECT";

            string pathName = null;
            try { pathName = doc.PathName; }
            catch { }

            if (!string.IsNullOrWhiteSpace(pathName))
                return pathName;

            string title = null;
            try { title = doc.Title; }
            catch { }

            if (string.IsNullOrWhiteSpace(title))
                title = "UNTITLED";

            return "UNSAVED::" + title;
        }

        private void RefreshIfcCatalog(string folderPath)
        {
            string[] filePaths;
            try
            {
                filePaths = Directory.GetFiles(folderPath, "*.ifc", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Failed to read IFC folder:\n\n" + ex.Message,
                    "Folder Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
                SessionLogger.Error("Failed to enumerate IFC folder: " + folderPath, ex);
                return;
            }

            Array.Sort(filePaths, StringComparer.OrdinalIgnoreCase);
            var inFolder = new HashSet<string>(filePaths, StringComparer.OrdinalIgnoreCase);

            // Keep scene consistent with the active catalog: unload models outside this folder.
            foreach (IfcModel model in _loadedModels.ToList())
            {
                if (!inFolder.Contains(model.FilePath))
                    UnloadIfcModel(model, removeFromCache: false, updateStatus: false);
            }

            _ifcCatalogItems.Clear();
            foreach (string path in filePaths)
            {
                var item = new IfcModelListItem(path)
                {
                    IsLoaded = FindLoadedModel(path) != null
                };
                _ifcCatalogItems.Add(item);
            }
        }

        private async Task LoadIfcModelByPathAsync(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            if (!File.Exists(path))
            {
                MessageBox.Show(this,
                    "IFC file not found:\n\n" + path,
                    "IFC Load Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IfcModel alreadyLoaded = FindLoadedModel(path);
            if (alreadyLoaded != null)
            {
                SetCatalogLoadedState(path, true);
                SelectCatalogItem(path);
                return;
            }

            UpdateStatus("Loading: " + Path.GetFileName(path) + " …");
            AddIfcButton.IsEnabled = false;
            RemoveIfcButton.IsEnabled = false;

            bool loadedFromCache = false;
            try
            {
                // Cancel any previous in-flight load
                _loadCts?.Cancel();
                _loadCts = new CancellationTokenSource();

                DateTime writeUtc = File.GetLastWriteTimeUtc(path);
                CachedIfcModel cached;
                IfcModel ifcModel = null;

                if (_ifcModelCache.TryGetValue(path, out cached)
                    && cached != null
                    && cached.Model != null
                    && cached.LastWriteUtc == writeUtc)
                {
                    ifcModel = cached.Model;
                    loadedFromCache = true;
                }
                else
                {
                    ifcModel = await IfcLoader.LoadAsync(path, Dispatcher, onProgress: UpdateStatus,
                                                      cancellationToken: _loadCts.Token);
                    _ifcModelCache[path] = new CachedIfcModel(ifcModel, writeUtc);
                }

                AttachIfcModel(ifcModel);
                SetCatalogLoadedState(path, true);
                SelectCatalogItem(path);

                UpdateStatus(ifcModel.DisplayName + "  |  " +
                             ifcModel.MeshCount + " elements  |  " +
                             ifcModel.TriangleCount + " triangles" +
                             (loadedFromCache ? "  |  cache" : ""));
                SessionLogger.Info("Loaded: " + ifcModel.DisplayName +
                                   (loadedFromCache ? " (cache)" : ""));
            }
            catch (OperationCanceledException)
            {
                // Load was cancelled (e.g. user started a new load or closed the window)
                SessionLogger.Info("IFC load cancelled: " + path);
                UpdateStatus("Load cancelled.");
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
                AddIfcButton.IsEnabled = true;
                RemoveIfcButton.IsEnabled = true;
            }
        }

        private void AttachIfcModel(IfcModel ifcModel)
        {
            if (ifcModel == null) return;

            _viewerHost.IfcRoot.Children.Add(ifcModel.SceneGroup);
            _loadedModels.Add(ifcModel);

            foreach (var kvp in ifcModel.ElementMap)
                _ifcElementMap[kvp.Key] = kvp.Value;
            RebuildIfcGuidMap();

            _sectionMgr?.RegisterGroup(ifcModel.SceneGroup);
            RebuildOutline();
            if (WireframeToggle?.IsChecked == true)
                RebuildWireframe();

            UpdateSectionBounds();

            if (_loadedModels.Count == 1 && _revitModel == null)
                _viewerHost.FitView(ifcModel.Bounds);

            RegisterIfcWatcher(ifcModel.FilePath);
        }

        private void RegisterIfcWatcher(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;

            IfcFileWatcher existing;
            if (_fileWatchers.TryGetValue(path, out existing))
            {
                existing.Dispose();
                _fileWatchers.Remove(path);
            }

            var watcher = new IfcFileWatcher(path,
                () => Dispatcher.BeginInvoke((Action)(() => ShowReloadBanner(path))));
            _fileWatchers[path] = watcher;
        }

        private IfcModel FindLoadedModel(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            foreach (IfcModel model in _loadedModels)
            {
                if (string.Equals(model.FilePath, path, StringComparison.OrdinalIgnoreCase))
                    return model;
            }
            return null;
        }

        private void SetCatalogLoadedState(string path, bool isLoaded)
        {
            foreach (IfcModelListItem item in _ifcCatalogItems)
            {
                if (!string.Equals(item.FilePath, path, StringComparison.OrdinalIgnoreCase))
                    continue;

                item.IsLoaded = isLoaded;
                return;
            }
        }

        private void SelectCatalogItem(string path)
        {
            foreach (IfcModelListItem item in _ifcCatalogItems)
            {
                if (!string.Equals(item.FilePath, path, StringComparison.OrdinalIgnoreCase))
                    continue;

                ModelListBox.SelectedItem = item;
                return;
            }
        }

        private void UnloadIfcModel(IfcModel selected, bool removeFromCache, bool updateStatus = true)
        {
            if (selected == null) return;

            IfcFileWatcher watcher;
            if (_fileWatchers.TryGetValue(selected.FilePath, out watcher))
            {
                watcher.Dispose();
                _fileWatchers.Remove(selected.FilePath);
            }

            if (_pendingReload != null &&
                string.Equals(_pendingReload.FilePath, selected.FilePath, StringComparison.OrdinalIgnoreCase))
            {
                _pendingReload = null;
                ReloadBanner.Visibility = Visibility.Collapsed;
            }

            if (_selectedMesh != null && selected.ElementMap.ContainsKey(_selectedMesh))
                ClearSelection();

            foreach (var mesh in selected.ElementMap.Keys)
                _ifcElementMap.Remove(mesh);

            if (_hiddenMeshes.Count > 0)
            {
                var kept = _hiddenMeshes.Where(m => !selected.ElementMap.ContainsKey(m)).ToList();
                _hiddenMeshes.Clear();
                for (int i = kept.Count - 1; i >= 0; i--)
                    _hiddenMeshes.Push(kept[i]);
            }
            RebuildIfcGuidMap();

            _sectionMgr?.UnregisterGroup(selected.SceneGroup);
            _viewerHost.IfcRoot.Children.Remove(selected.SceneGroup);
            _loadedModels.Remove(selected);
            SetCatalogLoadedState(selected.FilePath, false);

            if (removeFromCache)
                _ifcModelCache.Remove(selected.FilePath);

            RebuildOutline();
            if (WireframeToggle?.IsChecked == true)
                RebuildWireframe();
            UpdateSectionBounds();

            if (updateStatus)
            {
                if (_loadedModels.Count == 0)
                    UpdateStatus(_ifcCatalogItems.Count == 0
                        ? "GPU Viewport Active"
                        : "No IFC loaded — double-click a model name.");
                else
                    UpdateStatus(_loadedModels.Count + " model(s) loaded");
            }
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

            // Rebuild element outlines (always-on) and optional wireframe overlay
            RebuildOutline();
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

            // Force full rebuild after an external file change.
            _ifcModelCache.Remove(path);
            UnloadIfcModel(model, removeFromCache: false, updateStatus: false);
            await LoadIfcModelByPathAsync(path);
            SessionLogger.Info("Reloaded: " + Path.GetFileNameWithoutExtension(path));
        }

        private void ReloadBannerDismiss_Click(object sender, RoutedEventArgs e)
        {
            ReloadBanner.Visibility = Visibility.Collapsed;
            _pendingReload = null;
        }

        private bool TryHandleSpacebar()
        {
            if (_selectedMesh != null)
            {
                _hiddenMeshes.Push(_selectedMesh);
                _selectedMesh.Visibility = Visibility.Collapsed;
                ClearSelection();

                // Rebuild wireframe/outline from visible meshes only — the hidden mesh's
                // lines must not linger after the mesh itself disappears.
                RebuildOutline();
                if (WireframeToggle?.IsChecked == true)
                    RebuildWireframe();

                SessionLogger.Info("Hidden selected element.");
                return true;
            }
            else if (_hiddenMeshes.Count > 0)
            {
                var mesh = _hiddenMeshes.Pop();
                mesh.Visibility = Visibility.Visible;
                
                if (_ifcElementMap.TryGetValue(mesh, out IfcElementInfo ifcInfo))
                    SelectElement(mesh, ifcInfo);
                else if (_revitElementMap.TryGetValue(mesh, out RevitElementInfo revitInfo))
                    SelectElement(mesh, revitInfo);

                RebuildOutline();
                if (WireframeToggle?.IsChecked == true)
                    RebuildWireframe();

                SessionLogger.Info("Unhidden last element.");
                return true;
            }

            return false;
        }

        // ── Element selection ─────────────────────────────────────────────────

        private void UnhideAll_Click(object sender, RoutedEventArgs e)
        {
            if (_hiddenMeshes.Count == 0) return;

            while (_hiddenMeshes.Count > 0)
            {
                var mesh = _hiddenMeshes.Pop();
                mesh.Visibility = Visibility.Visible;
            }

            RebuildOutline();
            if (WireframeToggle?.IsChecked == true)
                RebuildWireframe();

            SessionLogger.Info("Unhid all hidden elements.");
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

            // Orbit around the selected element's centre when right-dragging.
            if (_viewport != null && mesh.Geometry != null)
            {
                var bb = mesh.Geometry.Bound;
                var c  = (bb.Minimum + bb.Maximum) * 0.5f;
                _viewport.FixedRotationPoint        = new Media3D.Point3D(c.X, c.Y, c.Z);
                _viewport.FixedRotationPointEnabled = true;
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

            // Restore default orbit behaviour (scene-centre / mouse-down pivot).
            if (_viewport != null)
                _viewport.FixedRotationPointEnabled = false;
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

        private void FollowSelection_Checked(object sender, RoutedEventArgs e)
            => SetFollowSelectionEnabled(true);

        private void FollowSelection_Unchecked(object sender, RoutedEventArgs e)
            => SetFollowSelectionEnabled(false);

        private void SetFollowSelectionEnabled(bool enabled)
        {
            if (_applyingFollowSelectionState) return;

            if (_settings != null)
            {
                _settings.FollowSelectionEnabled = enabled;
                _settings.Save();
            }

            if (_followSelectionService != null)
                _followSelectionService.IsEnabled = enabled;

            SessionLogger.Info(enabled
                ? "Follow selection enabled."
                : "Follow selection disabled.");
        }

        private void OnRevitPrimarySelectionChanged(
            UIDocument uiDoc,
            Autodesk.Revit.DB.ElementId elementId)
        {
            if (uiDoc == null
                || elementId == null
                || elementId == Autodesk.Revit.DB.ElementId.InvalidElementId)
                return;

            MeshGeometryModel3D targetMesh;
            string resolutionMode;
            if (!TryResolveViewerMesh(uiDoc, elementId, out targetMesh, out resolutionMode))
                return;

            Action focusAction = () =>
            {
                if (_viewerFocusService?.FocusByMesh(targetMesh) == true)
                {
                    SessionLogger.Info(
                        $"Follow selection: focused ElementId {elementId.Value} via {resolutionMode}.");
                }
            };

            if (Dispatcher.CheckAccess()) focusAction();
            else Dispatcher.BeginInvoke(focusAction);
        }

        private bool TryResolveViewerMesh(
            UIDocument uiDoc,
            Autodesk.Revit.DB.ElementId elementId,
            out MeshGeometryModel3D mesh,
            out string resolutionMode)
        {
            mesh = null;
            resolutionMode = null;

            // A) Direct Revit ElementId → rendered mesh mapping.
            if (_revitModel?.ElementMeshes != null
                && _revitModel.ElementMeshes.TryGetValue(elementId, out mesh))
            {
                resolutionMode = "ElementId";
                return true;
            }

            var doc = uiDoc.Document;
            var element = doc?.GetElement(elementId);
            if (element == null) return false;

            // B) IFC GUID mapping (when IFC model is loaded and GUIDs are available).
            string ifcGuid;
            if (TryGetIfcGuid(element, out ifcGuid)
                && _ifcGuidMeshMap.TryGetValue(ifcGuid, out mesh))
            {
                resolutionMode = "IFC GUID";
                return true;
            }

            // C) Fallback spatial mapping using element bounding-box center.
            Vector3 center;
            if (TryGetElementCenter(uiDoc, element, out center))
            {
                float toleranceM = 0.2f;
                if (_settings != null)
                {
                    toleranceM = (float)Math.Max(10.0,
                        _settings.FollowSelectionSpatialToleranceMm) / 1000f;
                }

                if (TryFindNearestRenderedMesh(center, toleranceM, out mesh))
                {
                    resolutionMode = "spatial";
                    return true;
                }
            }

            return false;
        }

        private bool TryGetIfcGuid(Autodesk.Revit.DB.Element element, out string ifcGuid)
        {
            ifcGuid = null;
            if (element == null) return false;

            if (TryGetIfcGuidFromElement(element, out ifcGuid))
                return true;

            try
            {
                Autodesk.Revit.DB.ElementId typeId = element.GetTypeId();
                if (typeId != null && typeId != Autodesk.Revit.DB.ElementId.InvalidElementId)
                {
                    Autodesk.Revit.DB.Element typeElement = element.Document.GetElement(typeId);
                    if (TryGetIfcGuidFromElement(typeElement, out ifcGuid))
                        return true;
                }
            }
            catch
            {
                // Type lookup is best-effort only.
            }

            return false;
        }

        private static bool TryGetIfcGuidFromElement(
            Autodesk.Revit.DB.Element element,
            out string ifcGuid)
        {
            ifcGuid = null;
            if (element == null) return false;

            foreach (string parameterName in IfcGuidParameterNames)
            {
                var p = element.LookupParameter(parameterName);
                if (TryReadGuidParameter(p, out ifcGuid))
                    return true;
            }

            foreach (Autodesk.Revit.DB.Parameter parameter in element.Parameters)
            {
                string parameterName = parameter?.Definition?.Name;
                if (string.IsNullOrWhiteSpace(parameterName)) continue;

                foreach (string candidate in IfcGuidParameterNames)
                {
                    if (!string.Equals(parameterName, candidate, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (TryReadGuidParameter(parameter, out ifcGuid))
                        return true;
                }
            }

            return false;
        }

        private static bool TryReadGuidParameter(
            Autodesk.Revit.DB.Parameter parameter,
            out string value)
        {
            value = null;
            if (parameter == null) return false;

            try { value = parameter.AsString(); }
            catch { }

            if (string.IsNullOrWhiteSpace(value))
            {
                try { value = parameter.AsValueString(); }
                catch { }
            }

            value = NormalizeIfcGuid(value);
            return !string.IsNullOrEmpty(value);
        }

        private static bool TryGetElementCenter(
            UIDocument uiDoc,
            Autodesk.Revit.DB.Element element,
            out Vector3 center)
        {
            center = new Vector3();
            if (element == null) return false;

            Autodesk.Revit.DB.BoundingBoxXYZ bb = null;
            try { bb = element.get_BoundingBox(uiDoc?.ActiveView); }
            catch { }

            if (bb == null)
            {
                try { bb = element.get_BoundingBox(null); }
                catch { }
            }

            if (bb == null || bb.Min == null || bb.Max == null) return false;

            Autodesk.Revit.DB.XYZ xyzCenter = (bb.Min + bb.Max) * 0.5;
            center = ToViewerPoint(xyzCenter);
            return true;
        }

        private bool TryFindNearestRenderedMesh(
            Vector3 point,
            float toleranceMeters,
            out MeshGeometryModel3D nearestMesh)
        {
            nearestMesh = null;

            float bestDistSq = toleranceMeters * toleranceMeters;
            Vector3 center;

            foreach (MeshGeometryModel3D mesh in _ifcElementMap.Keys)
            {
                if (!TryGetMeshCenter(mesh, out center)) continue;
                float distSq = (center - point).LengthSquared();
                if (distSq > bestDistSq) continue;

                bestDistSq = distSq;
                nearestMesh = mesh;
            }

            if (_revitModel?.ElementMeshes != null)
            {
                foreach (MeshGeometryModel3D mesh in _revitModel.ElementMeshes.Values)
                {
                    if (!TryGetMeshCenter(mesh, out center)) continue;
                    float distSq = (center - point).LengthSquared();
                    if (distSq > bestDistSq) continue;

                    bestDistSq = distSq;
                    nearestMesh = mesh;
                }
            }

            return nearestMesh != null;
        }

        private static bool TryGetMeshCenter(MeshGeometryModel3D mesh, out Vector3 center)
        {
            center = new Vector3();
            if (mesh?.Geometry == null) return false;

            BoundingBox bb = mesh.Geometry.Bound;
            center = (bb.Minimum + bb.Maximum) * 0.5f;
            return true;
        }

        private static Vector3 ToViewerPoint(Autodesk.Revit.DB.XYZ point)
        {
            const float feetToMeters = 0.3048f;
            return new Vector3(
                (float)point.X * feetToMeters,
                (float)point.Z * feetToMeters,
               -(float)point.Y * feetToMeters);
        }

        private void RebuildIfcGuidMap()
        {
            _ifcGuidMeshMap.Clear();

            foreach (var kv in _ifcElementMap)
            {
                string key = NormalizeIfcGuid(kv.Value?.GlobalId);
                if (string.IsNullOrEmpty(key)) continue;
                if (_ifcGuidMeshMap.ContainsKey(key)) continue;

                _ifcGuidMeshMap[key] = kv.Key;
            }
        }

        private static string NormalizeIfcGuid(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            return raw.Trim().Trim('{', '}', '(', ')').ToUpperInvariant();
        }

        private void Wireframe_Checked(object sender, RoutedEventArgs e)
            => RebuildWireframe();

        private void Wireframe_Unchecked(object sender, RoutedEventArgs e)
        {
            _wireframeRoot?.Children.Clear();
            SessionLogger.Info("Wireframe off.");
        }

        // ── ViewCube helpers ──────────────────────────────────────────────────

        /// <summary>
        /// Creates a TextureModel (6 x 128 px strip, Revit-style) with six face labels
        /// in Helix's expected order: Front | Back | Left | Right | Top | Bottom.
        /// The bitmap is PNG-encoded into a MemoryStream so TextureModel can consume it.
        /// </summary>
        private static TextureModel CreateRevitViewCubeTexture()
        {
            const int S = 128; // pixels per face
            var labels = new[] { "FRONT", "BACK", "LEFT", "RIGHT", "TOP", "BOTTOM" };

            // Stone-gray faces with subtle per-face brightness shifts (Revit-like).
            var faceColors = new[]
            {
                WpfColor.FromRgb(0xC2, 0xC7, 0xCB), // Front  — neutral
                WpfColor.FromRgb(0xA7, 0xAC, 0xB0), // Back   — darker
                WpfColor.FromRgb(0xB7, 0xBC, 0xC1), // Left   — mid-dark
                WpfColor.FromRgb(0xD0, 0xD4, 0xD8), // Right  — lighter
                WpfColor.FromRgb(0xDB, 0xDE, 0xE2), // Top    — lightest
                WpfColor.FromRgb(0x9F, 0xA4, 0xA8), // Bottom — darkest
            };

            var outerBorderPen = new Pen(new SolidColorBrush(WpfColor.FromRgb(0x7A, 0x7E, 0x82)), 1.0);
            var innerBorderPen = new Pen(new SolidColorBrush(WpfColor.FromArgb(155, 255, 255, 255)), 1.0);
            var textBrush = new SolidColorBrush(WpfColor.FromRgb(0x2A, 0x2A, 0x2A));
            var typeface  = new Typeface(
                new FontFamily("Segoe UI"),
                FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                for (int i = 0; i < labels.Length; i++)
                {
                    var rect = new Rect(i * S, 0, S, S);
                    var baseColor = faceColors[i];
                    var faceBrush = new LinearGradientBrush(
                        ShiftColor(baseColor, +14),
                        ShiftColor(baseColor, -16),
                        new WpfPoint(0, 0),
                        new WpfPoint(1, 1));

                    dc.DrawRectangle(faceBrush, outerBorderPen, rect);
                    dc.DrawRectangle(null, innerBorderPen, new Rect(rect.X + 1.5, rect.Y + 1.5, S - 3, S - 3));

                    double fontSize = labels[i].Length > 4 ? 16.0 : 20.0;
                    var ft = new FormattedText(
                        labels[i], CultureInfo.InvariantCulture,
                        FlowDirection.LeftToRight, typeface, fontSize, textBrush, 1.0);

                    dc.DrawText(ft, new WpfPoint(
                        rect.X + (S - ft.Width)  / 2,
                        rect.Y + (S - ft.Height) / 2));
                }
            }

            var rtb = new RenderTargetBitmap(S * labels.Length, S, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);

            // TextureModel accepts a Stream; encode the bitmap as PNG in memory.
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            var ms = new System.IO.MemoryStream();
            encoder.Save(ms);
            ms.Seek(0, System.IO.SeekOrigin.Begin);
            return new TextureModel(ms);
        }

        /// <summary>
        /// Builds the WPF compass ring overlay that wraps the ViewCube.
        /// Compass ring and centre cube are both projected from full camera
        /// orientation, so the ring stays attached under the cube in 3D.
        /// </summary>
        private FrameworkElement BuildCompassOverlay()
        {
            const double D  = 96;           // overall canvas size (px)
            const double cx = D / 2, cy = D / 2;
            var rootCanvas = new Canvas { Width = D, Height = D, Opacity = 0.95 };

            InitializeCompass3D(rootCanvas);
            InitializeOrientationCube(rootCanvas, cx, cy);

            // ── Container: anchors top-right and wraps around the ViewCube ───
            return new Border
            {
                Width               = D,
                Height              = D,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Top,
                Margin              = new Thickness(0, 2, 24, 0),
                Child               = rootCanvas,
            };
        }

        private void InitializeCompass3D(Canvas canvas)
        {
            _compassRing = new WpfShapes.Path
            {
                Fill            = new SolidColorBrush(WpfColor.FromArgb(255, 236, 236, 236)),
                Stroke          = new SolidColorBrush(WpfColor.FromArgb(255, 95, 95, 95)),
                StrokeThickness = 1.0,
                StrokeLineJoin  = PenLineJoin.Round,
                IsHitTestVisible = false,
            };
            canvas.Children.Add(_compassRing);

            _compassLabels.Clear();
            var northLabel = new TextBlock
            {
                Text             = "N",
                FontSize         = 31.0,
                FontWeight       = FontWeights.ExtraBold,
                Foreground       = new SolidColorBrush(WpfColor.FromArgb(240, 60, 60, 60)),
                IsHitTestVisible = false,
            };
            var southLabel = new TextBlock
            {
                Text             = "S",
                FontSize         = 31.0,
                FontWeight       = FontWeights.ExtraBold,
                Foreground       = new SolidColorBrush(WpfColor.FromArgb(220, 80, 80, 80)),
                IsHitTestVisible = false,
            };
            var eastLabel = new TextBlock
            {
                Text             = "E",
                FontSize         = 31.0,
                FontWeight       = FontWeights.ExtraBold,
                Foreground       = new SolidColorBrush(WpfColor.FromArgb(220, 80, 80, 80)),
                IsHitTestVisible = false,
            };
            var westLabel = new TextBlock
            {
                Text             = "W",
                FontSize         = 31.0,
                FontWeight       = FontWeights.ExtraBold,
                Foreground       = new SolidColorBrush(WpfColor.FromArgb(220, 80, 80, 80)),
                IsHitTestVisible = false,
            };

            _compassLabels["N"] = northLabel;
            _compassLabels["S"] = southLabel;
            _compassLabels["E"] = eastLabel;
            _compassLabels["W"] = westLabel;

            canvas.Children.Add(northLabel);
            canvas.Children.Add(southLabel);
            canvas.Children.Add(eastLabel);
            canvas.Children.Add(westLabel);
        }

        private void InitializeOrientationCube(Canvas canvas, double cx, double cy)
        {
            _orientationCubeCenterX = cx;
            _orientationCubeCenterY = cy;
            _orientationCubeFaces.Clear();
            _orientationCubeFaceLabels.Clear();
            _orientationCubeFaceBrushes.Clear();
            _orientationCubeFaceHoverBrushes.Clear();
            _orientationCubeEdgeHotspots.Clear();
            _orientationCubeCornerHotspots.Clear();
            _orientationCubeArrowButtons.Clear();
            _orientationCubeArrowTargets.Clear();
            _hoveredOrientationCubeFace = null;
            _hoveredOrientationArrow = null;
            _hoveredOrientationEdgeIndex = null;
            _hoveredOrientationCornerIndex = null;

            foreach (var face in OrientationCubeFaceOrder)
            {
                var baseColor = GetOrientationCubeFaceColor(face);
                var baseBrush = new SolidColorBrush(baseColor);
                var hoverBrush = new SolidColorBrush(ShiftColor(baseColor, +28));
                _orientationCubeFaceBrushes[face] = baseBrush;
                _orientationCubeFaceHoverBrushes[face] = hoverBrush;

                var poly = new WpfShapes.Polygon
                {
                    Stroke          = new SolidColorBrush(WpfColor.FromArgb(255, 120, 120, 120)),
                    StrokeThickness = 0.8,
                    Fill            = baseBrush,
                    Visibility      = System.Windows.Visibility.Collapsed,
                    IsHitTestVisible = true,
                    Cursor          = Cursors.Hand,
                    Tag             = face,
                };
                poly.MouseEnter += OrientationCubeFace_MouseEnter;
                poly.MouseLeave += OrientationCubeFace_MouseLeave;
                poly.MouseLeftButtonDown += OrientationCubeFace_MouseLeftButtonDown;

                _orientationCubeFaces.Add(face, poly);
                canvas.Children.Add(poly);

                var label = new TextBlock
                {
                    Text             = GetOrientationCubeFaceText(face),
                    FontSize         = 8.0,
                    FontWeight       = FontWeights.SemiBold,
                    FontFamily       = new FontFamily("Segoe UI"),
                    Foreground       = new SolidColorBrush(WpfColor.FromArgb(230, 65, 65, 65)),
                    Visibility       = System.Windows.Visibility.Collapsed,
                    IsHitTestVisible = false,
                };
                _orientationCubeFaceLabels.Add(face, label);
                canvas.Children.Add(label);
            }

            for (int i = 0; i < OrientationCubeEdges.Length; i++)
            {
                var edgeHotspot = new WpfShapes.Polygon
                {
                    Fill             = Brushes.Transparent,
                    Stroke           = Brushes.Transparent,
                    StrokeThickness  = 1.2,
                    Visibility       = System.Windows.Visibility.Collapsed,
                    IsHitTestVisible = true,
                    Cursor           = Cursors.Hand,
                    Tag              = i,
                };
                edgeHotspot.MouseEnter += OrientationCubeEdge_MouseEnter;
                edgeHotspot.MouseLeave += OrientationCubeEdge_MouseLeave;
                edgeHotspot.MouseLeftButtonDown += OrientationCubeEdge_MouseLeftButtonDown;
                _orientationCubeEdgeHotspots[i] = edgeHotspot;
                canvas.Children.Add(edgeHotspot);
            }

            for (int i = 0; i < OrientationCubeCorners.Length; i++)
            {
                var cornerHotspot = new WpfShapes.Ellipse
                {
                    Width            = 8.0,
                    Height           = 8.0,
                    Fill             = Brushes.Transparent,
                    Stroke           = Brushes.Transparent,
                    StrokeThickness  = 1.1,
                    Visibility       = System.Windows.Visibility.Collapsed,
                    IsHitTestVisible = true,
                    Cursor           = Cursors.Hand,
                    Tag              = i,
                };
                cornerHotspot.MouseEnter += OrientationCubeCorner_MouseEnter;
                cornerHotspot.MouseLeave += OrientationCubeCorner_MouseLeave;
                cornerHotspot.MouseLeftButtonDown += OrientationCubeCorner_MouseLeftButtonDown;
                _orientationCubeCornerHotspots[i] = cornerHotspot;
                canvas.Children.Add(cornerHotspot);
            }

            foreach (var direction in OrientationArrowOrder)
            {
                var arrow = new WpfShapes.Polygon
                {
                    Fill             = _orientationArrowFillBrush,
                    Stroke           = new SolidColorBrush(WpfColor.FromArgb(255, 95, 95, 95)),
                    StrokeThickness  = 1.0,
                    Visibility       = System.Windows.Visibility.Collapsed,
                    IsHitTestVisible = true,
                    Cursor           = Cursors.Hand,
                    Tag              = direction,
                };
                arrow.MouseEnter += OrientationArrow_MouseEnter;
                arrow.MouseLeave += OrientationArrow_MouseLeave;
                arrow.MouseLeftButtonDown += OrientationArrow_MouseLeftButtonDown;
                _orientationCubeArrowButtons[direction] = arrow;
                canvas.Children.Add(arrow);
            }
        }

        private static WpfColor GetOrientationCubeFaceColor(OrientationCubeFace face)
        {
            switch (face)
            {
                case OrientationCubeFace.Top:    return WpfColor.FromArgb(255, 0xE7, 0xE9, 0xEC);
                case OrientationCubeFace.Bottom: return WpfColor.FromArgb(255, 0xA2, 0xA6, 0xAA);
                case OrientationCubeFace.Right:  return WpfColor.FromArgb(255, 0xC9, 0xCD, 0xD1);
                case OrientationCubeFace.Left:   return WpfColor.FromArgb(255, 0xB8, 0xBC, 0xC0);
                case OrientationCubeFace.Back:   return WpfColor.FromArgb(255, 0xAC, 0xB0, 0xB4);
                default:                         return WpfColor.FromArgb(255, 0xCD, 0xD1, 0xD5); // Front
            }
        }

        private static string GetOrientationCubeFaceText(OrientationCubeFace face)
        {
            switch (face)
            {
                case OrientationCubeFace.Front:  return "FRONT";
                case OrientationCubeFace.Back:   return "BACK";
                case OrientationCubeFace.Left:   return "LEFT";
                case OrientationCubeFace.Right:  return "RIGHT";
                case OrientationCubeFace.Top:    return "TOP";
                case OrientationCubeFace.Bottom: return "BOTTOM";
                default:                         return "FRONT";
            }
        }

        private static int[] GetOrientationCubeFaceIndices(OrientationCubeFace face)
        {
            switch (face)
            {
                case OrientationCubeFace.Front:  return OrientationFaceFront;
                case OrientationCubeFace.Back:   return OrientationFaceBack;
                case OrientationCubeFace.Left:   return OrientationFaceLeft;
                case OrientationCubeFace.Right:  return OrientationFaceRight;
                case OrientationCubeFace.Top:    return OrientationFaceTop;
                case OrientationCubeFace.Bottom: return OrientationFaceBottom;
                default:                         return OrientationFaceFront;
            }
        }

        private void UpdateOrientationCubeVisual()
        {
            if (_viewerHost?.Camera == null || _orientationCubeFaces.Count == 0) return;

            var forward = _viewerHost.Camera.LookDirection;
            if (forward.LengthSquared < 1e-9) return;
            forward.Normalize();

            var up = _viewerHost.Camera.UpDirection;
            if (up.LengthSquared < 1e-9)
                up = new Media3D.Vector3D(0, 1, 0);
            else
                up.Normalize();

            var right = Media3D.Vector3D.CrossProduct(forward, up);
            if (right.LengthSquared < 1e-9)
            {
                right = Media3D.Vector3D.CrossProduct(forward, new Media3D.Vector3D(0, 1, 0));
                if (right.LengthSquared < 1e-9)
                    right = Media3D.Vector3D.CrossProduct(forward, new Media3D.Vector3D(1, 0, 0));
            }
            right.Normalize();
            up = Media3D.Vector3D.CrossProduct(right, forward);
            up.Normalize();

            UpdateCompass3DVisual(right, up, forward);

            var projectedVertices = new WpfPoint[OrientationCubeVertices.Length];
            var projectedDepths = new double[OrientationCubeVertices.Length];
            for (int i = 0; i < OrientationCubeVertices.Length; i++)
                projectedVertices[i] = ProjectWidgetPoint(OrientationCubeVertices[i], right, up, forward, out projectedDepths[i]);

            var visibleFaces = new List<KeyValuePair<OrientationCubeFace, double>>();

            foreach (var face in OrientationCubeFaceOrder)
            {
                var poly = _orientationCubeFaces[face];
                var label = _orientationCubeFaceLabels[face];
                var faceNormal = GetOrientationCubeFaceNormal(face);
                if (Media3D.Vector3D.DotProduct(faceNormal, forward) >= -0.001)
                {
                    poly.Visibility = System.Windows.Visibility.Collapsed;
                    label.Visibility = System.Windows.Visibility.Collapsed;
                    continue;
                }

                var indices = GetOrientationCubeFaceIndices(face);
                var points = new PointCollection(indices.Length);
                double depth = 0.0;

                for (int i = 0; i < indices.Length; i++)
                {
                    points.Add(projectedVertices[indices[i]]);
                    depth += projectedDepths[indices[i]];
                }

                poly.Points = points;
                poly.Visibility = System.Windows.Visibility.Visible;
                visibleFaces.Add(new KeyValuePair<OrientationCubeFace, double>(face, depth / indices.Length));
            }

            // Painter order: far faces first.
            visibleFaces.Sort((a, b) => a.Value.CompareTo(b.Value));
            for (int i = 0; i < visibleFaces.Count; i++)
            {
                var face = visibleFaces[i].Key;
                var poly = _orientationCubeFaces[face];
                var label = _orientationCubeFaceLabels[face];
                ApplyOrientationCubeFaceFill(face, poly);
                Panel.SetZIndex(poly, 100 + i);
                PositionOrientationCubeFaceLabel(face, poly.Points, label, 220 + i);
            }

            foreach (var face in OrientationCubeFaceOrder)
            {
                bool isVisible = false;
                for (int i = 0; i < visibleFaces.Count; i++)
                    if (visibleFaces[i].Key == face) { isVisible = true; break; }

                if (!isVisible)
                    _orientationCubeFaceLabels[face].Visibility = System.Windows.Visibility.Collapsed;
            }

            UpdateOrientationCubeEdgeAndCornerHotspots(visibleFaces, projectedVertices, projectedDepths);
            UpdateOrientationCubeArrows(visibleFaces);

            if (_hoveredOrientationCubeFace.HasValue)
            {
                bool hoveredStillVisible = false;
                for (int i = 0; i < visibleFaces.Count; i++)
                {
                    if (visibleFaces[i].Key == _hoveredOrientationCubeFace.Value)
                    {
                        hoveredStillVisible = true;
                        break;
                    }
                }
                if (!hoveredStillVisible)
                    _hoveredOrientationCubeFace = null;
            }
        }

        private void OrientationCubeFace_MouseEnter(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationCubeFace)) return;

            var face = (OrientationCubeFace)poly.Tag;
            _hoveredOrientationCubeFace = face;
            ApplyOrientationCubeFaceFill(face, poly);
        }

        private void OrientationCubeFace_MouseLeave(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationCubeFace)) return;

            var face = (OrientationCubeFace)poly.Tag;
            if (_hoveredOrientationCubeFace.HasValue && _hoveredOrientationCubeFace.Value == face)
                _hoveredOrientationCubeFace = null;
            ApplyOrientationCubeFaceFill(face, poly);
        }

        private void OrientationCubeFace_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationCubeFace)) return;

            var face = (OrientationCubeFace)poly.Tag;
            SnapCameraToOrientationFace(face);
            e.Handled = true;
        }

        private void ApplyOrientationCubeFaceFill(OrientationCubeFace face, WpfShapes.Polygon poly)
        {
            if (poly == null) return;

            bool isHovered = _hoveredOrientationCubeFace.HasValue && _hoveredOrientationCubeFace.Value == face;
            SolidColorBrush brush;
            if (isHovered && _orientationCubeFaceHoverBrushes.TryGetValue(face, out brush))
            {
                poly.Fill = brush;
                return;
            }

            if (_orientationCubeFaceBrushes.TryGetValue(face, out brush))
                poly.Fill = brush;
        }

        private void OrientationCubeEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is int)) return;

            int edgeIndex = (int)poly.Tag;
            if (edgeIndex < 0 || edgeIndex >= OrientationCubeEdges.Length) return;

            var edge = OrientationCubeEdges[edgeIndex];
            SnapCameraToOrientationNormals(
                GetOrientationCubeFaceNormal(edge.FaceA),
                GetOrientationCubeFaceNormal(edge.FaceB));
            e.Handled = true;
        }

        private void OrientationCubeEdge_MouseEnter(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is int)) return;

            int edgeIndex = (int)poly.Tag;
            _hoveredOrientationEdgeIndex = edgeIndex;
            ApplyOrientationCubeEdgeHighlight(edgeIndex, poly);
        }

        private void OrientationCubeEdge_MouseLeave(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is int)) return;

            int edgeIndex = (int)poly.Tag;
            if (_hoveredOrientationEdgeIndex.HasValue && _hoveredOrientationEdgeIndex.Value == edgeIndex)
                _hoveredOrientationEdgeIndex = null;
            ApplyOrientationCubeEdgeHighlight(edgeIndex, poly);
        }

        private void OrientationCubeCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var ellipse = sender as WpfShapes.Ellipse;
            if (ellipse == null || !(ellipse.Tag is int)) return;

            int cornerIndex = (int)ellipse.Tag;
            if (cornerIndex < 0 || cornerIndex >= OrientationCubeCorners.Length) return;

            var corner = OrientationCubeCorners[cornerIndex];
            SnapCameraToOrientationNormals(
                GetOrientationCubeFaceNormal(corner.FaceA),
                GetOrientationCubeFaceNormal(corner.FaceB),
                GetOrientationCubeFaceNormal(corner.FaceC));
            e.Handled = true;
        }

        private void OrientationCubeCorner_MouseEnter(object sender, MouseEventArgs e)
        {
            var ellipse = sender as WpfShapes.Ellipse;
            if (ellipse == null || !(ellipse.Tag is int)) return;

            int cornerIndex = (int)ellipse.Tag;
            _hoveredOrientationCornerIndex = cornerIndex;
            ApplyOrientationCubeCornerHighlight(cornerIndex, ellipse);
        }

        private void OrientationCubeCorner_MouseLeave(object sender, MouseEventArgs e)
        {
            var ellipse = sender as WpfShapes.Ellipse;
            if (ellipse == null || !(ellipse.Tag is int)) return;

            int cornerIndex = (int)ellipse.Tag;
            if (_hoveredOrientationCornerIndex.HasValue && _hoveredOrientationCornerIndex.Value == cornerIndex)
                _hoveredOrientationCornerIndex = null;
            ApplyOrientationCubeCornerHighlight(cornerIndex, ellipse);
        }

        private void ApplyOrientationCubeEdgeHighlight(int edgeIndex, WpfShapes.Polygon hotspot)
        {
            if (hotspot == null) return;
            bool isHovered = _hoveredOrientationEdgeIndex.HasValue && _hoveredOrientationEdgeIndex.Value == edgeIndex;
            hotspot.Fill = isHovered ? _orientationEdgeHoverFillBrush : Brushes.Transparent;
            hotspot.Stroke = isHovered ? _orientationEdgeHoverStrokeBrush : Brushes.Transparent;
        }

        private void ApplyOrientationCubeCornerHighlight(int cornerIndex, WpfShapes.Ellipse hotspot)
        {
            if (hotspot == null) return;
            bool isHovered = _hoveredOrientationCornerIndex.HasValue && _hoveredOrientationCornerIndex.Value == cornerIndex;
            hotspot.Fill = isHovered ? _orientationCornerHoverFillBrush : Brushes.Transparent;
            hotspot.Stroke = isHovered ? _orientationCornerHoverStrokeBrush : Brushes.Transparent;
        }

        private void OrientationArrow_MouseEnter(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationArrowDirection)) return;

            var direction = (OrientationArrowDirection)poly.Tag;
            _hoveredOrientationArrow = direction;
            ApplyOrientationArrowFill(direction, poly);
        }

        private void OrientationArrow_MouseLeave(object sender, MouseEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationArrowDirection)) return;

            var direction = (OrientationArrowDirection)poly.Tag;
            if (_hoveredOrientationArrow.HasValue && _hoveredOrientationArrow.Value == direction)
                _hoveredOrientationArrow = null;
            ApplyOrientationArrowFill(direction, poly);
        }

        private void OrientationArrow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var poly = sender as WpfShapes.Polygon;
            if (poly == null || !(poly.Tag is OrientationArrowDirection)) return;

            var direction = (OrientationArrowDirection)poly.Tag;
            OrientationCubeFace targetFace;
            if (_orientationCubeArrowTargets.TryGetValue(direction, out targetFace))
                SnapCameraToOrientationFace(targetFace);

            e.Handled = true;
        }

        private void UpdateOrientationCubeEdgeAndCornerHotspots(
            List<KeyValuePair<OrientationCubeFace, double>> visibleFaces,
            WpfPoint[] projectedVertices,
            double[] projectedDepths)
        {
            var visibleSet = new HashSet<OrientationCubeFace>();
            for (int i = 0; i < visibleFaces.Count; i++)
                visibleSet.Add(visibleFaces[i].Key);

            const double edgeThickness = 6.0;
            for (int i = 0; i < OrientationCubeEdges.Length; i++)
            {
                WpfShapes.Polygon hotspot;
                if (!_orientationCubeEdgeHotspots.TryGetValue(i, out hotspot)) continue;

                var edge = OrientationCubeEdges[i];
                bool show = visibleSet.Contains(edge.FaceA) || visibleSet.Contains(edge.FaceB);
                if (!show)
                {
                    if (_hoveredOrientationEdgeIndex.HasValue && _hoveredOrientationEdgeIndex.Value == i)
                        _hoveredOrientationEdgeIndex = null;
                    hotspot.Visibility = System.Windows.Visibility.Collapsed;
                    continue;
                }

                var a = projectedVertices[edge.V0];
                var b = projectedVertices[edge.V1];
                double dx = b.X - a.X;
                double dy = b.Y - a.Y;
                double len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 0.001)
                {
                    hotspot.Visibility = System.Windows.Visibility.Collapsed;
                    continue;
                }

                double nx = -dy / len * edgeThickness * 0.5;
                double ny =  dx / len * edgeThickness * 0.5;

                hotspot.Points = new PointCollection
                {
                    new WpfPoint(a.X + nx, a.Y + ny),
                    new WpfPoint(b.X + nx, b.Y + ny),
                    new WpfPoint(b.X - nx, b.Y - ny),
                    new WpfPoint(a.X - nx, a.Y - ny),
                };
                hotspot.Visibility = System.Windows.Visibility.Visible;
                Panel.SetZIndex(hotspot, 300);
                ApplyOrientationCubeEdgeHighlight(i, hotspot);
            }

            const double cornerSize = 8.0;
            for (int i = 0; i < OrientationCubeCorners.Length; i++)
            {
                WpfShapes.Ellipse hotspot;
                if (!_orientationCubeCornerHotspots.TryGetValue(i, out hotspot)) continue;

                var corner = OrientationCubeCorners[i];
                bool show = projectedDepths[corner.Vertex] > 0.0
                    && (visibleSet.Contains(corner.FaceA)
                        || visibleSet.Contains(corner.FaceB)
                        || visibleSet.Contains(corner.FaceC));
                if (!show)
                {
                    if (_hoveredOrientationCornerIndex.HasValue && _hoveredOrientationCornerIndex.Value == i)
                        _hoveredOrientationCornerIndex = null;
                    hotspot.Visibility = System.Windows.Visibility.Collapsed;
                    continue;
                }

                var p = projectedVertices[corner.Vertex];
                Canvas.SetLeft(hotspot, p.X - cornerSize * 0.5);
                Canvas.SetTop(hotspot,  p.Y - cornerSize * 0.5);
                hotspot.Visibility = System.Windows.Visibility.Visible;
                Panel.SetZIndex(hotspot, 310);
                ApplyOrientationCubeCornerHighlight(i, hotspot);
            }
        }

        private void UpdateOrientationCubeArrows(List<KeyValuePair<OrientationCubeFace, double>> visibleFaces)
        {
            if (visibleFaces.Count != 1)
            {
                HideOrientationCubeArrows();
                return;
            }

            var face = visibleFaces[0].Key;
            WpfShapes.Polygon facePoly;
            if (!_orientationCubeFaces.TryGetValue(face, out facePoly)
                || facePoly.Points == null
                || facePoly.Points.Count < 4)
            {
                HideOrientationCubeArrows();
                return;
            }

            var points = facePoly.Points;
            var p0 = points[0];
            var p1 = points[1];
            var p3 = points[3];

            var center = PolygonCentroid(points);
            var xAxis = new Media3D.Vector3D(p1.X - p0.X, p1.Y - p0.Y, 0.0);
            var yAxis = new Media3D.Vector3D(p3.X - p0.X, p3.Y - p0.Y, 0.0);
            double xLen = Math.Sqrt(xAxis.X * xAxis.X + xAxis.Y * xAxis.Y);
            double yLen = Math.Sqrt(yAxis.X * yAxis.X + yAxis.Y * yAxis.Y);
            if (xLen < 0.001 || yLen < 0.001)
            {
                HideOrientationCubeArrows();
                return;
            }

            xAxis = new Media3D.Vector3D(xAxis.X / xLen, xAxis.Y / xLen, 0.0);
            yAxis = new Media3D.Vector3D(yAxis.X / yLen, yAxis.Y / yLen, 0.0);

            var indices = GetOrientationCubeFaceIndices(face);
            var v0 = OrientationCubeVertices[indices[0]];
            var v1 = OrientationCubeVertices[indices[1]];
            var v3 = OrientationCubeVertices[indices[3]];

            var localX = new Media3D.Vector3D(v1.X - v0.X, v1.Y - v0.Y, v1.Z - v0.Z);
            var localY = new Media3D.Vector3D(v3.X - v0.X, v3.Y - v0.Y, v3.Z - v0.Z);
            if (localX.LengthSquared < 1e-9 || localY.LengthSquared < 1e-9)
            {
                HideOrientationCubeArrows();
                return;
            }
            localX.Normalize();
            localY.Normalize();

            _orientationCubeArrowTargets[OrientationArrowDirection.Left] = GetOrientationFaceFromNormal(
                new Media3D.Vector3D(-localX.X, -localX.Y, -localX.Z));
            _orientationCubeArrowTargets[OrientationArrowDirection.Right] = GetOrientationFaceFromNormal(localX);
            _orientationCubeArrowTargets[OrientationArrowDirection.Up] = GetOrientationFaceFromNormal(localY);
            _orientationCubeArrowTargets[OrientationArrowDirection.Down] = GetOrientationFaceFromNormal(
                new Media3D.Vector3D(-localY.X, -localY.Y, -localY.Z));

            PositionOrientationArrow(
                OrientationArrowDirection.Left,
                center, xAxis, yAxis, xLen * 0.5, yLen * 0.5);
            PositionOrientationArrow(
                OrientationArrowDirection.Right,
                center, xAxis, yAxis, xLen * 0.5, yLen * 0.5);
            PositionOrientationArrow(
                OrientationArrowDirection.Up,
                center, xAxis, yAxis, xLen * 0.5, yLen * 0.5);
            PositionOrientationArrow(
                OrientationArrowDirection.Down,
                center, xAxis, yAxis, xLen * 0.5, yLen * 0.5);
        }

        private void PositionOrientationArrow(
            OrientationArrowDirection direction,
            WpfPoint center,
            Media3D.Vector3D xAxis,
            Media3D.Vector3D yAxis,
            double halfWidth,
            double halfHeight)
        {
            WpfShapes.Polygon arrow;
            if (!_orientationCubeArrowButtons.TryGetValue(direction, out arrow)) return;

            // Arrow is drawn as a balanced triangle inside an invisible square.
            const double arrowBoxSize = 10.0;
            const double arrowGapFromCube = 5.5;
            double arrowBoxHalf = arrowBoxSize * 0.5;
            double triHeight = arrowBoxSize * 0.68;
            double triHalfBase = triHeight / Math.Sqrt(3.0);
            // Tip must face toward the cube.
            double tipOffset = -triHeight * 0.5;
            double baseOffset = triHeight * 0.5;

            Media3D.Vector3D outward;
            Media3D.Vector3D side;
            double edgeDistance;
            switch (direction)
            {
                case OrientationArrowDirection.Left:
                    outward = new Media3D.Vector3D(-xAxis.X, -xAxis.Y, 0.0);
                    side = yAxis;
                    edgeDistance = halfWidth;
                    break;
                case OrientationArrowDirection.Right:
                    outward = xAxis;
                    side = yAxis;
                    edgeDistance = halfWidth;
                    break;
                case OrientationArrowDirection.Up:
                    outward = yAxis;
                    side = xAxis;
                    edgeDistance = halfHeight;
                    break;
                default:
                    outward = new Media3D.Vector3D(-yAxis.X, -yAxis.Y, 0.0);
                    side = xAxis;
                    edgeDistance = halfHeight;
                    break;
            }

            var arrowCenter = new WpfPoint(
                center.X + outward.X * (edgeDistance + arrowGapFromCube + arrowBoxHalf),
                center.Y + outward.Y * (edgeDistance + arrowGapFromCube + arrowBoxHalf));

            var tip = new WpfPoint(
                arrowCenter.X + outward.X * tipOffset,
                arrowCenter.Y + outward.Y * tipOffset);
            var baseCenter = new WpfPoint(
                arrowCenter.X + outward.X * baseOffset,
                arrowCenter.Y + outward.Y * baseOffset);

            arrow.Points = new PointCollection
            {
                tip,
                new WpfPoint(baseCenter.X + side.X * triHalfBase, baseCenter.Y + side.Y * triHalfBase),
                new WpfPoint(baseCenter.X - side.X * triHalfBase, baseCenter.Y - side.Y * triHalfBase),
            };
            arrow.Visibility = System.Windows.Visibility.Visible;
            Panel.SetZIndex(arrow, 360);
            ApplyOrientationArrowFill(direction, arrow);
        }

        private void HideOrientationCubeArrows()
        {
            _orientationCubeArrowTargets.Clear();
            _hoveredOrientationArrow = null;
            foreach (var direction in OrientationArrowOrder)
            {
                WpfShapes.Polygon arrow;
                if (_orientationCubeArrowButtons.TryGetValue(direction, out arrow))
                    arrow.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void ApplyOrientationArrowFill(OrientationArrowDirection direction, WpfShapes.Polygon arrow)
        {
            if (arrow == null) return;

            bool isHovered = _hoveredOrientationArrow.HasValue && _hoveredOrientationArrow.Value == direction;
            arrow.Fill = isHovered ? _orientationArrowHoverFillBrush : _orientationArrowFillBrush;
        }

        private void SnapCameraToOrientationFace(OrientationCubeFace face)
        {
            SnapCameraToOrientationNormals(GetOrientationCubeFaceNormal(face));
        }

        private void SnapCameraToOrientationNormals(params Media3D.Vector3D[] outwardNormals)
        {
            if (_viewerHost?.Camera == null || outwardNormals == null || outwardNormals.Length == 0) return;

            var outward = new Media3D.Vector3D(0, 0, 0);
            for (int i = 0; i < outwardNormals.Length; i++)
            {
                outward.X += outwardNormals[i].X;
                outward.Y += outwardNormals[i].Y;
                outward.Z += outwardNormals[i].Z;
            }

            if (outward.LengthSquared < 1e-9) return;
            outward.Normalize();

            var viewDirection = new Media3D.Vector3D(-outward.X, -outward.Y, -outward.Z);
            var upDirection = ComputeCameraUpDirection(viewDirection);
            SnapCameraToOrientation(viewDirection, upDirection);
        }

        private void SnapCameraToOrientation(Media3D.Vector3D viewDirection, Media3D.Vector3D upDirection)
        {
            if (_viewerHost?.Camera == null) return;
            if (viewDirection.LengthSquared < 1e-9) return;
            if (upDirection.LengthSquared < 1e-9) upDirection = new Media3D.Vector3D(0, 1, 0);

            var camera = _viewerHost.Camera;
            double distance = camera.LookDirection.Length;
            if (distance < 0.001) distance = 10.0;

            viewDirection.Normalize();
            upDirection.Normalize();

            var target = camera.Position + camera.LookDirection;
            var newLookDirection = new Media3D.Vector3D(
                viewDirection.X * distance,
                viewDirection.Y * distance,
                viewDirection.Z * distance);

            camera.LookDirection = newLookDirection;
            camera.UpDirection = upDirection;
            camera.Position = target - newLookDirection;

            UpdateCameraOrientationWidgets(force: true);
            _viewport.InvalidateRender();
        }

        private static Media3D.Vector3D ComputeCameraUpDirection(Media3D.Vector3D viewDirection)
        {
            var upHint = new Media3D.Vector3D(0, 1, 0);
            var view = viewDirection;
            if (view.LengthSquared < 1e-9) return upHint;
            view.Normalize();

            if (Math.Abs(Media3D.Vector3D.DotProduct(view, upHint)) > 0.97)
                upHint = new Media3D.Vector3D(0, 0, -1);

            var right = Media3D.Vector3D.CrossProduct(view, upHint);
            if (right.LengthSquared < 1e-9)
                right = Media3D.Vector3D.CrossProduct(view, new Media3D.Vector3D(1, 0, 0));
            if (right.LengthSquared < 1e-9)
                return new Media3D.Vector3D(0, 1, 0);

            right.Normalize();
            var up = Media3D.Vector3D.CrossProduct(right, view);
            if (up.LengthSquared < 1e-9)
                return new Media3D.Vector3D(0, 1, 0);
            up.Normalize();
            return up;
        }

        private static OrientationCubeFace GetOrientationFaceFromNormal(Media3D.Vector3D normalDirection)
        {
            if (normalDirection.LengthSquared < 1e-9)
                return OrientationCubeFace.Front;

            var dir = normalDirection;
            dir.Normalize();

            var bestFace = OrientationCubeFace.Front;
            double bestDot = double.NegativeInfinity;
            foreach (var face in OrientationCubeFaceOrder)
            {
                var normal = GetOrientationCubeFaceNormal(face);
                double dot = Media3D.Vector3D.DotProduct(dir, normal);
                if (dot > bestDot)
                {
                    bestDot = dot;
                    bestFace = face;
                }
            }

            return bestFace;
        }

        private static Media3D.Vector3D GetOrientationCubeFaceNormal(OrientationCubeFace face)
        {
            switch (face)
            {
                case OrientationCubeFace.Front:  return new Media3D.Vector3D(0, 0, -1);
                case OrientationCubeFace.Back:   return new Media3D.Vector3D(0, 0,  1);
                case OrientationCubeFace.Left:   return new Media3D.Vector3D(-1, 0, 0);
                case OrientationCubeFace.Right:  return new Media3D.Vector3D( 1, 0, 0);
                case OrientationCubeFace.Top:    return new Media3D.Vector3D(0,  1, 0);
                case OrientationCubeFace.Bottom: return new Media3D.Vector3D(0, -1, 0);
                default:                         return new Media3D.Vector3D(0, 0, -1);
            }
        }

        private void PositionOrientationCubeFaceLabel(
            OrientationCubeFace face,
            PointCollection points,
            TextBlock label,
            int zIndex)
        {
            double area = PolygonArea(points);
            if (area < 85.0)
            {
                label.Visibility = System.Windows.Visibility.Collapsed;
                return;
            }

            label.Text = GetOrientationCubeFaceText(face);
            label.FontSize = face == OrientationCubeFace.Top ? 8.0 : 7.6;
            label.Visibility = System.Windows.Visibility.Visible;
            label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            var size = label.DesiredSize;
            if (points.Count < 4 || size.Width < 0.1 || size.Height < 0.1)
            {
                var fallbackCenter = PolygonCentroid(points);
                Canvas.SetLeft(label, fallbackCenter.X - size.Width * 0.5);
                Canvas.SetTop(label,  fallbackCenter.Y - size.Height * 0.52);
                label.RenderTransform = Transform.Identity;
                Panel.SetZIndex(label, zIndex);
                return;
            }

            var p0 = points[0];
            var p1 = points[1];
            var p3 = points[3];
            var center = PolygonCentroid(points);

            var xBasis = new Media3D.Vector3D(p1.X - p0.X, p1.Y - p0.Y, 0.0);
            var yBasis = new Media3D.Vector3D(p3.X - p0.X, p3.Y - p0.Y, 0.0);
            double xLen = Math.Sqrt(xBasis.X * xBasis.X + xBasis.Y * xBasis.Y);
            double yLen = Math.Sqrt(yBasis.X * yBasis.X + yBasis.Y * yBasis.Y);
            if (xLen < 0.001 || yLen < 0.001)
            {
                Canvas.SetLeft(label, center.X - size.Width * 0.5);
                Canvas.SetTop(label,  center.Y - size.Height * 0.52);
                label.RenderTransform = Transform.Identity;
                Panel.SetZIndex(label, zIndex);
                return;
            }

            // Keep cube face text screen-size stable instead of scaling by face foreshortening.
            double targetHeightPx = face == OrientationCubeFace.Top
                ? 10.0
                : 9.2;
            double scalePerPixel = targetHeightPx / size.Height;

            var xVec = new Media3D.Vector3D(
                xBasis.X / xLen * scalePerPixel,
                xBasis.Y / xLen * scalePerPixel,
                0.0);
            var yVec = new Media3D.Vector3D(
                yBasis.X / yLen * scalePerPixel,
                yBasis.Y / yLen * scalePerPixel,
                0.0);

            // Keep text from being mirrored when the projected face basis flips.
            double det = xVec.X * yVec.Y - xVec.Y * yVec.X;
            if (det < 0.0)
                xVec = new Media3D.Vector3D(-xVec.X, -xVec.Y, 0.0);

            // Side labels (FRONT/BACK/LEFT/RIGHT) are rotated 180 deg vs desired.
            // Flip both basis vectors to keep them upright on the face.
            if (face == OrientationCubeFace.Front
                || face == OrientationCubeFace.Back
                || face == OrientationCubeFace.Left
                || face == OrientationCubeFace.Right)
            {
                xVec = new Media3D.Vector3D(-xVec.X, -xVec.Y, 0.0);
                yVec = new Media3D.Vector3D(-yVec.X, -yVec.Y, 0.0);
            }

            var offset = new WpfPoint(
                center.X - (xVec.X * size.Width * 0.5 + yVec.X * size.Height * 0.52),
                center.Y - (xVec.Y * size.Width * 0.5 + yVec.Y * size.Height * 0.52));

            label.RenderTransform = new MatrixTransform(new System.Windows.Media.Matrix(
                xVec.X, xVec.Y,
                yVec.X, yVec.Y,
                offset.X, offset.Y));
            Canvas.SetLeft(label, 0.0);
            Canvas.SetTop(label,  0.0);
            Panel.SetZIndex(label, zIndex);
        }

        private static double PolygonArea(PointCollection points)
        {
            if (points == null || points.Count < 3) return 0.0;
            double a = 0.0;
            for (int i = 0; i < points.Count; i++)
            {
                var p1 = points[i];
                var p2 = points[(i + 1) % points.Count];
                a += p1.X * p2.Y - p2.X * p1.Y;
            }
            return Math.Abs(a) * 0.5;
        }

        private static WpfPoint PolygonCentroid(PointCollection points)
        {
            if (points == null || points.Count == 0) return new WpfPoint(0, 0);
            double x = 0.0, y = 0.0;
            for (int i = 0; i < points.Count; i++)
            {
                x += points[i].X;
                y += points[i].Y;
            }
            return new WpfPoint(x / points.Count, y / points.Count);
        }

        private static Geometry BuildCompassRingGeometry(PointCollection outerPoints, PointCollection innerPoints)
        {
            if (outerPoints == null || innerPoints == null || outerPoints.Count < 3 || innerPoints.Count < 3)
                return Geometry.Empty;

            var outerFigure = new PathFigure
            {
                StartPoint = outerPoints[0],
                IsClosed = true,
                IsFilled = true,
            };
            for (int i = 1; i < outerPoints.Count; i++)
                outerFigure.Segments.Add(new LineSegment(outerPoints[i], true));

            var innerFigure = new PathFigure
            {
                StartPoint = innerPoints[0],
                IsClosed = true,
                IsFilled = true,
            };
            for (int i = 1; i < innerPoints.Count; i++)
                innerFigure.Segments.Add(new LineSegment(innerPoints[i], true));

            return new PathGeometry(
                new[] { outerFigure, innerFigure },
                FillRule.EvenOdd,
                null);
        }

        private void UpdateCompass3DVisual(
            Media3D.Vector3D right,
            Media3D.Vector3D up,
            Media3D.Vector3D forward)
        {
            if (_compassRing == null) return;

            const int segments = 56;
            const double ringY = -1.0;
            const double ringOuterRadius = 2.95;
            const double ringInnerRadius = 2.15;
            const double labelRadius = (ringOuterRadius + ringInnerRadius) * 0.5;

            var outerPoints = new PointCollection(segments + 1);
            var innerPoints = new PointCollection(segments + 1);
            for (int i = 0; i <= segments; i++)
            {
                double a = (Math.PI * 2.0 * i) / segments;
                double d;
                outerPoints.Add(ProjectWidgetPoint(
                    new Media3D.Point3D(Math.Cos(a) * ringOuterRadius, ringY, Math.Sin(a) * ringOuterRadius),
                    right, up, forward, out d));
                innerPoints.Add(ProjectWidgetPoint(
                    new Media3D.Point3D(Math.Cos(a) * ringInnerRadius, ringY, Math.Sin(a) * ringInnerRadius),
                    right, up, forward, out d));
            }
            _compassRing.Data = BuildCompassRingGeometry(outerPoints, innerPoints);

            // Cardinal mapping with model north = -Z.
            var cardinalAngles = new[] { -Math.PI * 0.5, Math.PI * 0.5, 0.0, Math.PI };
            var cardinalNames = new[] { "N", "S", "E", "W" };

            for (int i = 0; i < 4; i++)
            {
                double a = cardinalAngles[i];
                double depth;
                TextBlock label;
                if (_compassLabels.TryGetValue(cardinalNames[i], out label))
                {
                    var tangentLocal = new Media3D.Vector3D(-Math.Sin(a), 0.0, Math.Cos(a));
                    var inwardLocal  = new Media3D.Vector3D(-Math.Cos(a), 0.0, -Math.Sin(a));
                    const double basisStep = 0.35;

                    var labelPoint = ProjectWidgetPoint(
                        new Media3D.Point3D(Math.Cos(a) * labelRadius, ringY, Math.Sin(a) * labelRadius),
                        right, up, forward, out depth);

                    label.Visibility = System.Windows.Visibility.Visible;

                    var tangentPoint = ProjectWidgetPoint(
                        new Media3D.Point3D(
                            Math.Cos(a) * labelRadius + tangentLocal.X * basisStep,
                            ringY,
                            Math.Sin(a) * labelRadius + tangentLocal.Z * basisStep),
                        right, up, forward, out depth);

                    var inwardPoint = ProjectWidgetPoint(
                        new Media3D.Point3D(
                            Math.Cos(a) * labelRadius + inwardLocal.X * basisStep,
                            ringY,
                            Math.Sin(a) * labelRadius + inwardLocal.Z * basisStep),
                        right, up, forward, out depth);

                    var tangent2D = new Media3D.Vector3D(
                        tangentPoint.X - labelPoint.X,
                        tangentPoint.Y - labelPoint.Y,
                        0.0);
                    var inward2D = new Media3D.Vector3D(
                        inwardPoint.X - labelPoint.X,
                        inwardPoint.Y - labelPoint.Y,
                        0.0);

                    bool rotateQuarterTurn = cardinalNames[i] == "E" || cardinalNames[i] == "W";
                    double localRotationDegrees = rotateQuarterTurn ? 90.0 : 0.0;
                    if (cardinalNames[i] == "E")
                        localRotationDegrees += 180.0;
                    PositionCompassLabelOnPlane(
                        label,
                        labelPoint,
                        tangent2D,
                        inward2D,
                        zIndex: 80,
                        localRotationDegrees: localRotationDegrees);
                }
            }
        }

        private static void PositionCompassLabelOnPlane(
            TextBlock label,
            WpfPoint anchor,
            Media3D.Vector3D tangent2D,
            Media3D.Vector3D inward2D,
            int zIndex,
            double localRotationDegrees = 0.0)
        {
            label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            var size = label.DesiredSize;
            if (size.Width < 0.1 || size.Height < 0.1) return;

            const double targetHeightWorld = 2.00;
            const double basisStep = 0.35; // must match caller
            double scalePerPixel = targetHeightWorld / size.Height;

            var xVec = new Media3D.Vector3D(
                tangent2D.X / basisStep * scalePerPixel,
                tangent2D.Y / basisStep * scalePerPixel,
                0.0);
            var yVec = new Media3D.Vector3D(
                inward2D.X / basisStep * scalePerPixel,
                inward2D.Y / basisStep * scalePerPixel,
                0.0);

            if (Math.Abs(localRotationDegrees) > 0.0001)
            {
                double radians = localRotationDegrees * (Math.PI / 180.0);
                double cos = Math.Cos(radians);
                double sin = Math.Sin(radians);

                var rotatedX = new Media3D.Vector3D(
                    xVec.X * cos + yVec.X * sin,
                    xVec.Y * cos + yVec.Y * sin,
                    0.0);
                var rotatedY = new Media3D.Vector3D(
                    -xVec.X * sin + yVec.X * cos,
                    -xVec.Y * sin + yVec.Y * cos,
                    0.0);

                xVec = rotatedX;
                yVec = rotatedY;
            }

            var offset = new WpfPoint(
                anchor.X - (xVec.X * size.Width * 0.5 + yVec.X * size.Height * 0.55),
                anchor.Y - (xVec.Y * size.Width * 0.5 + yVec.Y * size.Height * 0.55));

            label.RenderTransform = new MatrixTransform(new System.Windows.Media.Matrix(
                xVec.X, xVec.Y,
                yVec.X, yVec.Y,
                offset.X, offset.Y));
            Canvas.SetLeft(label, 0.0);
            Canvas.SetTop(label,  0.0);
            Panel.SetZIndex(label, zIndex);
        }

        private WpfPoint ProjectWidgetPoint(
            Media3D.Point3D local,
            Media3D.Vector3D right,
            Media3D.Vector3D up,
            Media3D.Vector3D forward,
            out double depth)
        {
            double camX = local.X * right.X + local.Y * right.Y + local.Z * right.Z;
            double camY = local.X * up.X    + local.Y * up.Y    + local.Z * up.Z;

            // Orthographic projection for Revit-like orientation widget behavior.
            depth = -(local.X * forward.X + local.Y * forward.Y + local.Z * forward.Z);

            return new WpfPoint(
                _orientationCubeCenterX + camX * OrientationCubeScale,
                _orientationCubeCenterY - camY * OrientationCubeScale);
        }

        private static WpfColor ShiftColor(WpfColor color, int delta)
        {
            int r = Math.Max(0, Math.Min(255, color.R + delta));
            int g = Math.Max(0, Math.Min(255, color.G + delta));
            int b = Math.Max(0, Math.Min(255, color.B + delta));
            return WpfColor.FromArgb(color.A, (byte)r, (byte)g, (byte)b);
        }

        private void AttachCameraOrientationSync()
        {
            if (_cameraSyncAttached || _viewerHost?.Camera == null) return;

            _cameraOrientationChangedHandler = (s, e) => UpdateCameraOrientationWidgets(force: false);
            _cameraLookDirectionDescriptor = DependencyPropertyDescriptor
                .FromProperty(Media3D.PerspectiveCamera.LookDirectionProperty,
                              typeof(Media3D.PerspectiveCamera));
            _cameraUpDirectionDescriptor = DependencyPropertyDescriptor
                .FromProperty(Media3D.PerspectiveCamera.UpDirectionProperty,
                              typeof(Media3D.PerspectiveCamera));
            _cameraPositionDescriptor = DependencyPropertyDescriptor
                .FromProperty(Media3D.PerspectiveCamera.PositionProperty,
                              typeof(Media3D.PerspectiveCamera));

            if (_cameraLookDirectionDescriptor != null)
                _cameraLookDirectionDescriptor.AddValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);
            if (_cameraUpDirectionDescriptor != null)
                _cameraUpDirectionDescriptor.AddValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);
            if (_cameraPositionDescriptor != null)
                _cameraPositionDescriptor.AddValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);

            CompositionTarget.Rendering += OnCameraRenderTick;
            _cameraSyncAttached = true;
        }

        private void DetachCameraOrientationSync()
        {
            if (!_cameraSyncAttached || _viewerHost?.Camera == null) return;

            if (_cameraLookDirectionDescriptor != null && _cameraOrientationChangedHandler != null)
                _cameraLookDirectionDescriptor.RemoveValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);
            if (_cameraUpDirectionDescriptor != null && _cameraOrientationChangedHandler != null)
                _cameraUpDirectionDescriptor.RemoveValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);
            if (_cameraPositionDescriptor != null && _cameraOrientationChangedHandler != null)
                _cameraPositionDescriptor.RemoveValueChanged(_viewerHost.Camera, _cameraOrientationChangedHandler);

            CompositionTarget.Rendering -= OnCameraRenderTick;

            _cameraLookDirectionDescriptor = null;
            _cameraUpDirectionDescriptor = null;
            _cameraPositionDescriptor = null;
            _cameraOrientationChangedHandler = null;
            _cameraSyncAttached = false;
            _hasCameraSnapshot = false;
        }

        private void OnCameraRenderTick(object sender, EventArgs e)
        {
            UpdateCameraOrientationWidgets(force: false);
        }

        private void UpdateCameraOrientationWidgets(bool force)
        {
            if (_viewerHost?.Camera == null || _viewport == null) return;

            var position = _viewerHost.Camera.Position;
            var look = _viewerHost.Camera.LookDirection;
            var up = _viewerHost.Camera.UpDirection;

            bool changed = force || !_hasCameraSnapshot
                || !IsClose(position, _lastCameraPosition)
                || !IsClose(look, _lastCameraLookDirection)
                || !IsClose(up, _lastCameraUpDirection);

            if (!changed) return;

            _lastCameraPosition = position;
            _lastCameraLookDirection = look;
            _lastCameraUpDirection = up;
            _hasCameraSnapshot = true;

            UpdateOrientationCubeVisual();
            _viewport.InvalidateVisual();
        }

        private static bool IsClose(Media3D.Point3D a, Media3D.Point3D b)
        {
            const double eps = 0.000001;
            return Math.Abs(a.X - b.X) < eps
                && Math.Abs(a.Y - b.Y) < eps
                && Math.Abs(a.Z - b.Z) < eps;
        }

        private static bool IsClose(Media3D.Vector3D a, Media3D.Vector3D b)
        {
            const double eps = 0.000001;
            return Math.Abs(a.X - b.X) < eps
                && Math.Abs(a.Y - b.Y) < eps
                && Math.Abs(a.Z - b.Z) < eps;
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

        // ── Always-on element outline (disabled) ─────────────────────────────
        // Outline overlay removed — the model renders as clean shaded geometry.
        private void RebuildOutline() { }

        /// <summary>
        /// Recursively collects all <see cref="MeshGeometry3D"/> geometry objects from
        /// the scene hierarchy, skipping line-only overlay roots.
        /// </summary>
        private void CollectMeshGeometries(GroupModel3D group, List<MeshGeometry3D> results)
        {
            foreach (var child in group.Children)
            {
                if (ReferenceEquals(child, _wireframeRoot))           continue;
                // Skip the section-plane quad — it is a 500 m helper mesh, not
                // building geometry, and would produce a giant diagonal line.
                if (_sectionMgr != null &&
                    ReferenceEquals(child, _sectionMgr.PlaneVisual)) continue;
                // Skip hidden elements — wireframe/outline should only reflect
                // what is actually visible on screen.
                if (child.Visibility != Visibility.Visible) continue;
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
            if (_settings == null) return;

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

            // Follow Revit selection
            if (_followSelectionService != null)
            {
                _followSelectionService.DebounceMilliseconds =
                    Math.Max(150, Math.Min(300, _settings.FollowSelectionDebounceMs));
                _followSelectionService.IsEnabled = _settings.FollowSelectionEnabled;
            }

            if (_viewerFocusService != null)
            {
                _viewerFocusService.DistanceMultiplier = Math.Max(
                    1.0, Math.Min(5.0, _settings.FollowSelectionDistanceMultiplier));
            }

            if (FollowSelectionToggle != null)
            {
                _applyingFollowSelectionState = true;
                try
                {
                    FollowSelectionToggle.IsChecked = _settings.FollowSelectionEnabled;
                }
                finally
                {
                    _applyingFollowSelectionState = false;
                }
            }
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

        private sealed class CachedIfcModel
        {
            public readonly IfcModel Model;
            public readonly DateTime LastWriteUtc;

            public CachedIfcModel(IfcModel model, DateTime lastWriteUtc)
            {
                Model = model;
                LastWriteUtc = lastWriteUtc;
            }
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
