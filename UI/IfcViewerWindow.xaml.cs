using Autodesk.Revit.UI;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace IfcViewer.UI
{
    public partial class IfcViewerWindow : Window
    {
        // ── P/Invoke for manual window resizing (windowless chrome) ──────────
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTLEFT        = 10;
        private const int HTRIGHT       = 11;
        private const int HTBOTTOM      = 15;
        private const int HTBOTTOMLEFT  = 16;
        private const int HTBOTTOMRIGHT = 17;

        // ── State ────────────────────────────────────────────────────────────
        private readonly UIApplication _uiApp;
        private bool _isDarkTheme = true;

        // ── Constructor ──────────────────────────────────────────────────────
        public IfcViewerWindow(UIApplication uiApp)
        {
            _uiApp = uiApp;
            InitializeComponent();
            SessionLogger.Info("IfcViewerWindow initialized.");
        }

        // ── Title-bar drag ───────────────────────────────────────────────────
        private void TitleBar_Loaded(object sender, RoutedEventArgs e)
        {
            // TitleBar dragging is handled by the TitleBar UserControl itself.
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            // No-op for now; resize handles use P/Invoke.
        }

        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // No-op for now.
        }

        // ── Theme toggle ─────────────────────────────────────────────────────
        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkTheme = !_isDarkTheme;
            SessionLogger.Info($"Theme toggled → {(_isDarkTheme ? "Dark" : "Light")}");
            // Theme switching implementation deferred to Stage 1+ (colour updates).
        }

        // ── Toolbar buttons ──────────────────────────────────────────────────
        private void AddIfc_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 2.
            SessionLogger.Info("AddIfc clicked (placeholder).");
        }

        private void RemoveIfc_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 2.
            SessionLogger.Info("RemoveIfc clicked (placeholder).");
        }

        private void SyncRevit_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 3.
            SessionLogger.Info("SyncRevit clicked (placeholder).");
        }

        private void ResetCamera_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Stage 1.
            SessionLogger.Info("ResetCamera clicked (placeholder).");
        }

        // ── Opacity sliders ──────────────────────────────────────────────────
        private void IfcOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // Implemented in Stage 2.
        }

        private void RevitOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // Implemented in Stage 3.
        }

        // ── Resize edge handlers ─────────────────────────────────────────────
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
