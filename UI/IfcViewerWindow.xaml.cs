using Autodesk.Revit.UI;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Viewer;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace IfcViewer.UI
{
    public partial class IfcViewerWindow : Window
    {
        // ── P/Invoke for manual window resizing ──────────────────────────────
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTLEFT           = 10;
        private const int HTRIGHT          = 11;
        private const int HTBOTTOM         = 15;
        private const int HTBOTTOMLEFT     = 16;
        private const int HTBOTTOMRIGHT    = 17;

        // ── Helix bindings (exposed as DPs for XAML binding) ─────────────────
        public static readonly DependencyProperty EffectsManagerProperty =
            DependencyProperty.Register(nameof(EffectsManager), typeof(EffectsManager), typeof(IfcViewerWindow));

        public static readonly DependencyProperty CameraProperty =
            DependencyProperty.Register(nameof(Camera), typeof(Camera), typeof(IfcViewerWindow));

        public EffectsManager EffectsManager
        {
            get => (EffectsManager)GetValue(EffectsManagerProperty);
            set => SetValue(EffectsManagerProperty, value);
        }

        public Camera Camera
        {
            get => (Camera)GetValue(CameraProperty);
            set => SetValue(CameraProperty, value);
        }

        // ── State ─────────────────────────────────────────────────────────────
        private readonly UIApplication _uiApp;
        private readonly ViewerHost    _viewerHost;
        private bool _isDarkTheme = true;

        // Scene root that Viewport3DX will render
        private readonly GroupModel3D _sceneRoot = new GroupModel3D();

        // ── Constructor ───────────────────────────────────────────────────────
        public IfcViewerWindow(UIApplication uiApp)
        {
            _uiApp = uiApp;

            // Create the GPU resource host first so DPs are set before InitializeComponent
            _viewerHost = new ViewerHost();
            EffectsManager = _viewerHost.EffectsManager;
            Camera         = _viewerHost.Camera;

            InitializeComponent();

            // Add the scene root to the viewport after the XAML tree is built
            Loaded += OnWindowLoaded;
        }

        // ── Loaded ────────────────────────────────────────────────────────────
        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            // Attach the scene root
            Viewport.Items.Add(_sceneRoot);

            // Build the test scene (lights + grid + axis + cube)
            _viewerHost.BuildTestScene(_sceneRoot);

            UpdateStatus($"Triangles: {CountTriangles()}  |  GPU Viewport Active");
            SessionLogger.Info("IfcViewerWindow loaded — test scene rendered.");
        }

        // ── Window close → dispose DirectX resources ──────────────────────────
        private void Window_Closed(object sender, EventArgs e)
        {
            // Clear scene items before disposing to avoid use-after-free
            _sceneRoot.Children.Clear();
            Viewport.Items.Clear();
            _viewerHost.Dispose();
        }

        // ── Title-bar drag ────────────────────────────────────────────────────
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
            // Implemented in Stage 2.
            SessionLogger.Info("AddIfc clicked (Stage 2 placeholder).");
        }

        private void RemoveIfc_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 2.
            SessionLogger.Info("RemoveIfc clicked (Stage 2 placeholder).");
        }

        private void SyncRevit_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 3.
            SessionLogger.Info("SyncRevit clicked (Stage 3 placeholder).");
        }

        private void ResetCamera_Click(object sender, RoutedEventArgs e)
        {
            _viewerHost.ResetCamera();
        }

        private void Wireframe_Checked(object sender, RoutedEventArgs e)
        {
            ApplyWireframe(true);
        }

        private void Wireframe_Unchecked(object sender, RoutedEventArgs e)
        {
            ApplyWireframe(false);
        }

        /// <summary>Walk scene and toggle wireframe on all MeshGeometryModel3D nodes.</summary>
        private void ApplyWireframe(bool on)
        {
            ApplyWireframeToGroup(_sceneRoot, on);
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
            // Applied to IfcRoot in Stage 2
            if (_viewerHost == null) return;
            SetGroupOpacity(_viewerHost.IfcRoot, e.NewValue);
        }

        private void RevitOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // Applied to RevitRoot in Stage 3
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

        // ── Status helper ──────────────────────────────────────────────────────
        public void UpdateStatus(string text)
        {
            if (StatusBar != null)
                StatusBar.Text = text;
        }

        private int CountTriangles()
        {
            return CountTrianglesInGroup(_sceneRoot);
        }

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

        // ── Resize edge handlers ───────────────────────────────────────────────
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
