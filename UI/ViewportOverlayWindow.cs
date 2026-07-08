using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace IfcViewer.UI
{
    /// <summary>
    /// Borderless transparent owned window that floats WPF content above the
    /// swap-chain viewport. The DXGI swap chain renders into a child HWND that
    /// covers all WPF elements in the same region (airspace), so overlays that
    /// must visually overlap the 3D view live in this window instead.
    ///
    /// The window never activates (WS_EX_NOACTIVATE) so clicks on overlay
    /// controls don't steal focus from the viewer, and it tracks the anchor
    /// element's top-right corner across owner move/resize/minimize.
    /// </summary>
    public sealed class ViewportOverlayWindow : Window
    {
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        private const int GWL_EXSTYLE      = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        private readonly Window           _ownerWindow;
        private readonly FrameworkElement _anchor;

        public ViewportOverlayWindow(Window owner, FrameworkElement anchor,
                                     UIElement content, double width, double height)
        {
            _ownerWindow = owner  ?? throw new ArgumentNullException(nameof(owner));
            _anchor      = anchor ?? throw new ArgumentNullException(nameof(anchor));

            Owner              = owner;
            WindowStyle        = WindowStyle.None;
            // Tiny window — the layered-window cost that was removed from the main
            // window is negligible at this size.
            AllowsTransparency = true;
            Background         = System.Windows.Media.Brushes.Transparent;
            ShowInTaskbar      = false;
            ShowActivated      = false;
            ResizeMode         = ResizeMode.NoResize;
            Width              = width;
            Height             = height;
            Content            = content;

            owner.LocationChanged  += OnAnchorChanged;
            owner.SizeChanged      += OnAnchorChanged;
            owner.StateChanged     += OnOwnerStateChanged;
            owner.IsVisibleChanged += OnOwnerVisibleChanged;
            anchor.SizeChanged     += OnAnchorChanged;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var hwnd = new WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_EXSTYLE,
                GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
        }

        /// <summary>Align this window's top-right corner with the anchor's.</summary>
        public void Reposition()
        {
            if (!IsVisible) return;
            if (!_anchor.IsLoaded || PresentationSource.FromVisual(_anchor) == null) return;

            try
            {
                // PointToScreen returns device pixels; window Left/Top are DIPs.
                var topRightDevice = _anchor.PointToScreen(new Point(_anchor.ActualWidth, 0));
                var source = PresentationSource.FromVisual(_ownerWindow);
                if (source?.CompositionTarget == null) return;

                var topRight = source.CompositionTarget.TransformFromDevice
                                     .Transform(topRightDevice);
                Left = topRight.X - Width;
                Top  = topRight.Y;
            }
            catch
            {
                // Anchor detached mid-layout — skip this tick; the next layout
                // change repositions again.
            }
        }

        private void OnAnchorChanged(object sender, EventArgs e) => Reposition();

        private void OnOwnerStateChanged(object sender, EventArgs e)
        {
            if (_ownerWindow.WindowState == WindowState.Minimized)
            {
                Hide();
            }
            else if (_ownerWindow.IsVisible)
            {
                Show();
                Reposition();
            }
        }

        private void OnOwnerVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_ownerWindow.IsVisible && _ownerWindow.WindowState != WindowState.Minimized)
            {
                Show();
                Reposition();
            }
            else
            {
                Hide();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _ownerWindow.LocationChanged  -= OnAnchorChanged;
            _ownerWindow.SizeChanged      -= OnAnchorChanged;
            _ownerWindow.StateChanged     -= OnOwnerStateChanged;
            _ownerWindow.IsVisibleChanged -= OnOwnerVisibleChanged;
            _anchor.SizeChanged           -= OnAnchorChanged;
            base.OnClosed(e);
        }
    }
}
