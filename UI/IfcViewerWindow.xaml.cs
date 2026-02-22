using Autodesk.Revit.UI;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Viewer;
using System;
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
                    EffectsManager        = _viewerHost.EffectsManager,
                    Camera                = _viewerHost.Camera,
                    ShowCoordinateSystem  = false,
                    ShowFrameRate         = true,
                    EnableSSAO            = false,
                    MSAA                  = MSAALevel.Four,
                    Background            = new System.Windows.Media.SolidColorBrush(
                                               System.Windows.Media.Color.FromRgb(26, 26, 26)),
                    FXAALevel             = FXAALevel.Low,
                    // Camera controller needs focus to receive mouse/keyboard input
                    Focusable             = true,
                    IsTabStop             = true,
                    // Zoom around the point under the cursor
                    ZoomAroundMouseDownPoint = true,
                    // Inspect mode: left=rotate, middle=pan, right-drag/scroll=zoom
                    CameraMode            = CameraMode.Inspect,
                };

                // 3. Scene root
                _sceneRoot = new GroupModel3D();
                _viewport.Items.Add(_sceneRoot);

                // 4. Insert viewport as first child of ViewportContainer (behind the status bar)
                // Give keyboard focus to the viewport on any mouse button press so the
                // Helix CameraController receives mouse-delta events for pan/zoom/rotate.
                _viewport.MouseDown += (s, ev) => _viewport.Focus();
                ViewportContainer.Children.Insert(0, _viewport);

                // 5. Build test scene
                _viewerHost.BuildTestScene(_sceneRoot);

                UpdateStatus($"GPU Viewport Active  |  Triangles: {CountTriangles()}");
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
        private void AddIfc_Click(object sender, RoutedEventArgs e)
            => SessionLogger.Info("AddIfc clicked (Stage 2 placeholder).");

        private void RemoveIfc_Click(object sender, RoutedEventArgs e)
            => SessionLogger.Info("RemoveIfc clicked (Stage 2 placeholder).");

        private void SyncRevit_Click(object sender, RoutedEventArgs e)
            => SessionLogger.Info("SyncRevit clicked (Stage 3 placeholder).");

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
