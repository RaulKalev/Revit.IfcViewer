using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace ProSchedules
{
    public class WindowResizer
    {
        private readonly Window _window;
        private ResizeDirection _resizeDirection;
        private Point _startPoint;
        private bool _isResizing;

        private const double MIN_WIDTH = 500;
        private const double MIN_HEIGHT = 400;
        public UIElement IgnoreElement { get; set; } // set this from MainWindow

        public WindowResizer(Window window)
        {
            _window = window;
        }

        public void StartResizing(MouseButtonEventArgs e, ResizeDirection direction)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                _resizeDirection = direction;
                _startPoint = e.GetPosition(null);
                _isResizing = true;
                _window.CaptureMouse();

                // Set cursor based on resize direction
                switch (_resizeDirection)
                {
                    case ResizeDirection.Left:
                    case ResizeDirection.Right:
                        Mouse.OverrideCursor = Cursors.SizeWE;
                        break;
                    case ResizeDirection.Bottom:
                        Mouse.OverrideCursor = Cursors.SizeNS;
                        break;
                    case ResizeDirection.BottomLeft:
                        Mouse.OverrideCursor = Cursors.SizeNESW;
                        break;
                    case ResizeDirection.BottomRight:
                        Mouse.OverrideCursor = Cursors.SizeNWSE;
                        break;
                }
            }
        }

        public void ResizeWindow(MouseEventArgs e)
        {
            if (IgnoreElement != null)
            {
                var mouseOver = System.Windows.Input.Mouse.DirectlyOver as DependencyObject;
                if (IsDescendantOf(mouseOver, IgnoreElement))
                    return;
            }

            if (!_isResizing) return;

            Point currentPoint = e.GetPosition(null); // Get position relative to screen
            double deltaX = currentPoint.X - _startPoint.X;
            double deltaY = currentPoint.Y - _startPoint.Y;

            switch (_resizeDirection)
            {
                case ResizeDirection.Left:
                    ResizeLeft(deltaX);
                    break;
                case ResizeDirection.Right:
                    ResizeRight(deltaX);
                    break;
                case ResizeDirection.Bottom:
                    ResizeBottom(deltaY);
                    break;
                case ResizeDirection.BottomLeft:
                    ResizeLeft(deltaX);
                    ResizeBottom(deltaY);
                    break;
                case ResizeDirection.BottomRight:
                    ResizeRight(deltaX);
                    ResizeBottom(deltaY);
                    break;
            }

            _startPoint = currentPoint;
        }
        private bool IsDescendantOf(DependencyObject child, DependencyObject parent)
        {
            while (child != null)
            {
                if (child == parent)
                    return true;

                DependencyObject newParent = null;

                if (child is Visual || child is System.Windows.Media.Media3D.Visual3D)
                {
                    newParent = VisualTreeHelper.GetParent(child);
                }

                // Fallback to logical tree if visual tree fails
                if (newParent == null)
                {
                    newParent = LogicalTreeHelper.GetParent(child);
                }

                child = newParent;
            }

            return false;
        }

        public void StopResizing()
        {
            if (_isResizing)
            {
                _isResizing = false;
                _window.ReleaseMouseCapture();
                Mouse.OverrideCursor = null; // Reset cursor
            }
        }

        private void ResizeLeft(double deltaX)
        {
            double newWidth = _window.Width - deltaX;
            if (newWidth >= MIN_WIDTH)
            {
                _window.Width = newWidth;
                _window.Left += deltaX;
            }
            else
            {
                // Prevent "pushing" behavior
                double offset = _window.Width - MIN_WIDTH;
                _window.Width = MIN_WIDTH;
                _window.Left += offset; // Adjust only the necessary offset
            }
        }

        private void ResizeRight(double deltaX)
        {
            double newWidth = _window.Width + deltaX;
            if (newWidth >= MIN_WIDTH)
            {
                _window.Width = newWidth;
            }
        }

        private void ResizeBottom(double deltaY)
        {
            double newHeight = _window.Height + deltaY;
            if (newHeight >= MIN_HEIGHT)
            {
                _window.Height = newHeight;
            }
        }
    }

    public enum ResizeDirection
    {
        Left,
        Right,
        Bottom,
        BottomLeft,
        BottomRight
    }
}

