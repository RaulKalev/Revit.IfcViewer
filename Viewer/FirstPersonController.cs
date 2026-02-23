using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Drives first-person WASD / arrow-key walk navigation on top of a Helix
    /// <see cref="PerspectiveCamera"/>.
    ///
    /// Usage:
    ///   1. Construct with the camera, the viewport (for key events) and the
    ///      window (for mouse Preview tunnel events that fire before Helix).
    ///   2. Call <see cref="Activate"/> to start; <see cref="Deactivate"/> to stop.
    /// </summary>
    public sealed class FirstPersonController : IDisposable
    {
        // ── Settings ─────────────────────────────────────────────────────────
        /// <summary>Base walk speed in metres per second.</summary>
        public double WalkSpeed { get; set; } = 5.0;

        /// <summary>Multiplier applied when Shift is held.</summary>
        public double SprintMultiplier { get; set; } = 3.0;

        /// <summary>Mouse look sensitivity (radians per pixel).</summary>
        public double MouseSensitivity { get; set; } = 0.003;

        // ── Dependencies ─────────────────────────────────────────────────────
        private readonly PerspectiveCamera _camera;
        private readonly UIElement         _keyTarget;   // viewport — keys
        private readonly UIElement         _mouseTarget; // window  — mouse Preview (fires before Helix)

        // ── State ────────────────────────────────────────────────────────────
        // CompositionTarget.Rendering fires exactly once per render frame — no
        // timer drift, no queued-up ticks when the frame is slow.
        private readonly HashSet<Key>           _keysDown = new HashSet<Key>();
        private          DateTime               _lastTick;
        private          bool                   _active;

        // Mouse-look state
        private bool                    _mouseLooking;
        private System.Windows.Point    _lastMousePos;

        // Yaw / pitch angles (radians) — derived from camera look direction on activation
        private double _yaw;
        private double _pitch;

        // ── Constructor ──────────────────────────────────────────────────────
        /// <param name="camera">The Helix PerspectiveCamera to drive.</param>
        /// <param name="keyTarget">The Viewport3DX — receives keyboard events.</param>
        /// <param name="mouseTarget">
        ///   The Window (or any ancestor of the viewport) — receives PreviewMouse
        ///   tunnel events that fire BEFORE Helix's CameraController, so mouse-look
        ///   can intercept right-drag before Helix does.
        /// </param>
        public FirstPersonController(PerspectiveCamera camera,
                                     UIElement keyTarget,
                                     UIElement mouseTarget)
        {
            _camera      = camera      ?? throw new ArgumentNullException(nameof(camera));
            _keyTarget   = keyTarget   ?? throw new ArgumentNullException(nameof(keyTarget));
            _mouseTarget = mouseTarget ?? keyTarget; // fall back to viewport if not supplied
        }

        // ── Public API ───────────────────────────────────────────────────────

        public bool IsActive => _active;

        /// <summary>Start first-person mode.</summary>
        public void Activate()
        {
            if (_active) return;
            _active = true;

            // Derive yaw/pitch from current camera look direction
            DecomposeLookDirection(_camera.LookDirection, out _yaw, out _pitch);

            _keysDown.Clear();
            _lastTick = DateTime.UtcNow;

            _keyTarget.PreviewKeyDown        += OnKeyDown;
            _keyTarget.PreviewKeyUp          += OnKeyUp;
            _keyTarget.LostFocus             += OnLostFocus;
            // Preview (tunnel) events on the window ancestor fire before Helix's
            // CameraController child can consume them — this is how we intercept
            // right-drag for mouse-look even when Helix would otherwise handle it.
            _mouseTarget.PreviewMouseDown    += OnMouseDown;
            _mouseTarget.PreviewMouseUp      += OnMouseUp;
            _mouseTarget.PreviewMouseMove    += OnMouseMove;

            // CompositionTarget.Rendering fires once per render frame (synchronized
            // with the GPU present cycle), eliminating timer drift and queued ticks.
            CompositionTarget.Rendering      += OnTick;
            _keyTarget.Focus();

            SessionLogger.Info("FirstPersonController: activated.");
        }

        /// <summary>Stop first-person mode and restore Inspect camera mode.</summary>
        public void Deactivate()
        {
            if (!_active) return;
            _active = false;

            CompositionTarget.Rendering      -= OnTick;
            _keysDown.Clear();
            StopMouseLook();

            _keyTarget.PreviewKeyDown        -= OnKeyDown;
            _keyTarget.PreviewKeyUp          -= OnKeyUp;
            _keyTarget.LostFocus             -= OnLostFocus;
            _mouseTarget.PreviewMouseDown    -= OnMouseDown;
            _mouseTarget.PreviewMouseUp      -= OnMouseUp;
            _mouseTarget.PreviewMouseMove    -= OnMouseMove;

            SessionLogger.Info("FirstPersonController: deactivated.");
        }

        public void Dispose()
        {
            Deactivate();
        }

        // ── Timer tick: move camera ───────────────────────────────────────────

        private void OnTick(object sender, EventArgs e)
        {
            if (!_active) return;

            var now   = DateTime.UtcNow;
            double dt = (now - _lastTick).TotalSeconds;
            _lastTick = now;
            dt        = Math.Min(dt, 0.1); // cap to avoid huge jumps after stall

            bool sprint = _keysDown.Contains(Key.LeftShift) || _keysDown.Contains(Key.RightShift);
            double speed = WalkSpeed * (sprint ? SprintMultiplier : 1.0) * dt;

            if (speed <= 0) return;

            // Forward / right / up vectors from current yaw+pitch
            GetVectors(_yaw, _pitch, out var forward, out var right, out var up);

            var delta = new Media3D.Vector3D(0, 0, 0);

            // Forward / back — WASD + arrows
            if (_keysDown.Contains(Key.W) || _keysDown.Contains(Key.Up))
                delta += ToWpf(forward) * speed;
            if (_keysDown.Contains(Key.S) || _keysDown.Contains(Key.Down))
                delta -= ToWpf(forward) * speed;

            // Strafe
            if (_keysDown.Contains(Key.D) || _keysDown.Contains(Key.Right))
                delta += ToWpf(right) * speed;
            if (_keysDown.Contains(Key.A) || _keysDown.Contains(Key.Left))
                delta -= ToWpf(right) * speed;

            // Vertical — Q/E or PgUp/PgDn
            if (_keysDown.Contains(Key.E) || _keysDown.Contains(Key.PageUp))
                delta += new Media3D.Vector3D(0, speed, 0);
            if (_keysDown.Contains(Key.Q) || _keysDown.Contains(Key.PageDown))
                delta -= new Media3D.Vector3D(0, speed, 0);

            if (delta.LengthSquared > 0)
            {
                var pos = _camera.Position + delta;
                _camera.Position      = pos;
                _camera.LookDirection = ToWpf(forward);
                _camera.UpDirection   = new Media3D.Vector3D(0, 1, 0);
            }
        }

        // ── Key events ───────────────────────────────────────────────────────

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (!_active) return;
            _keysDown.Add(e.Key);

            // Consume movement keys so the Helix viewport doesn't also pan/zoom.
            // This is only reached when walk mode is active (handlers are unregistered
            // on Deactivate), but the _active guard ensures no edge-case leakage.
            if (IsMovementKey(e.Key))
                e.Handled = true;
        }

        private void OnKeyUp(object sender, KeyEventArgs e)
        {
            if (!_active) return;
            _keysDown.Remove(e.Key);
        }

        private void OnLostFocus(object sender, RoutedEventArgs e)
        {
            _keysDown.Clear();
            StopMouseLook();
        }

        // ── Mouse-look ───────────────────────────────────────────────────────

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!_active) return;
            if (e.RightButton == MouseButtonState.Pressed)
            {
                _mouseLooking = true;
                _lastMousePos = e.GetPosition(_mouseTarget);
                _mouseTarget.CaptureMouse();
                e.Handled = true; // stops Helix CameraController from also rotating
            }
        }

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.RightButton == MouseButtonState.Released)
                StopMouseLook();
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_active || !_mouseLooking) return;

            var pos   = e.GetPosition(_mouseTarget);
            double dx = pos.X - _lastMousePos.X;
            double dy = pos.Y - _lastMousePos.Y;
            _lastMousePos = pos;

            _yaw   -= dx * MouseSensitivity;
            _pitch -= dy * MouseSensitivity;

            // Clamp pitch so we can't flip upside-down
            _pitch = Math.Max(-Math.PI / 2.0 + 0.01, Math.Min(Math.PI / 2.0 - 0.01, _pitch));

            GetVectors(_yaw, _pitch, out var forward, out _, out _);
            _camera.LookDirection = ToWpf(forward);
            _camera.UpDirection   = new Media3D.Vector3D(0, 1, 0);
            e.Handled = true; // stops Helix from panning/rotating on the same move
        }

        private void StopMouseLook()
        {
            if (_mouseLooking)
            {
                _mouseLooking = false;
                _mouseTarget.ReleaseMouseCapture();
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static bool IsMovementKey(Key k)
            => k == Key.W || k == Key.A || k == Key.S || k == Key.D
            || k == Key.Up || k == Key.Down || k == Key.Left || k == Key.Right
            || k == Key.Q || k == Key.E || k == Key.PageUp || k == Key.PageDown;

        /// <summary>
        /// Decompose a look direction into yaw (horizontal) and pitch (vertical) angles.
        /// Helix Y-up: yaw rotates around Y, pitch rotates around right axis.
        /// </summary>
        private static void DecomposeLookDirection(Media3D.Vector3D dir,
                                                   out double yaw, out double pitch)
        {
            var d = dir;
            d.Normalize();
            pitch = Math.Asin(Math.Max(-1, Math.Min(1, d.Y)));
            yaw   = Math.Atan2(-d.Z, d.X);
        }

        /// <summary>Return forward, right, up vectors for the given yaw/pitch (Helix Y-up).</summary>
        private static void GetVectors(double yaw, double pitch,
                                       out Vector3 forward, out Vector3 right, out Vector3 up)
        {
            float cy = (float)Math.Cos(yaw);
            float sy = (float)Math.Sin(yaw);
            float cp = (float)Math.Cos(pitch);
            float sp = (float)Math.Sin(pitch);

            forward = new Vector3(cy * cp, sp, -sy * cp);
            forward.Normalize();

            right = new Vector3(sy, 0, cy);  // perpendicular to forward in XZ
            right.Normalize();

            up = Vector3.Cross(right, forward);
            up.Normalize();
        }

        private static Media3D.Vector3D ToWpf(Vector3 v)
            => new Media3D.Vector3D(v.X, v.Y, v.Z);
    }
}
