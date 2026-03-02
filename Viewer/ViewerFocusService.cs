using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Applies camera focus operations without changing the user's current view
    /// orientation (yaw/pitch/roll).
    /// </summary>
    public sealed class ViewerFocusService
    {
        private readonly PerspectiveCamera _camera;
        private BoundingBox _lastFocusedBounds;
        private bool _hasLastFocus;

        /// <summary>Extra framing margin around the target bounds.</summary>
        public double PaddingRatio { get; set; } = 0.10;

        /// <summary>
        /// Multiplies the computed framing distance.
        /// Values greater than 1.0 place the camera farther away.
        /// </summary>
        public double DistanceMultiplier { get; set; } = 1.0;

        /// <summary>Camera animation time in milliseconds.</summary>
        public int AnimationMilliseconds { get; set; } = 250;

        public ViewerFocusService(PerspectiveCamera camera)
        {
            _camera = camera ?? throw new ArgumentNullException(nameof(camera));
        }

        public bool FocusByMesh(MeshGeometryModel3D mesh)
        {
            if (mesh?.Geometry == null) return false;
            return FocusByWorldBounds(mesh.Geometry.Bound);
        }

        public bool FocusByWorldBounds(BoundingBox bounds)
        {
            if (bounds.Maximum == bounds.Minimum) return false;
            if (IsSameAsLastFocus(bounds)) return false;

            var center = bounds.Center;
            var size = bounds.Maximum - bounds.Minimum;
            float radius = Math.Max(0.02f, size.Length() * 0.5f);
            radius *= (float)(1.0 + Math.Max(0.0, PaddingRatio));

            var lookDirection = _camera.LookDirection;
            if (lookDirection.LengthSquared < 1e-9)
                lookDirection = new Media3D.Vector3D(-1, -1, -1);
            lookDirection.Normalize();

            double halfFovRad = _camera.FieldOfView * 0.5 * Math.PI / 180.0;
            double distance = radius / Math.Max(Math.Tan(halfFovRad), 1e-5);
            distance = Math.Max(distance, radius * 1.1);
            distance *= Math.Max(0.05, DistanceMultiplier);

            var newLookDirection = lookDirection * distance;
            var target = new Media3D.Point3D(center.X, center.Y, center.Z);
            var newPosition = target - newLookDirection;

            var upDirection = _camera.UpDirection;
            if (upDirection.LengthSquared < 1e-9)
                upDirection = new Media3D.Vector3D(0, 1, 0);

            _camera.AnimateTo(newPosition, newLookDirection, upDirection, AnimationMilliseconds);

            double diagonal = size.Length();
            _camera.FarPlaneDistance = Math.Max(500.0, diagonal * 3.0);

            _lastFocusedBounds = bounds;
            _hasLastFocus = true;
            return true;
        }

        public void ResetLastFocus()
        {
            _hasLastFocus = false;
        }

        private bool IsSameAsLastFocus(BoundingBox bounds)
        {
            if (!_hasLastFocus) return false;

            var prevCenter = _lastFocusedBounds.Center;
            var nextCenter = bounds.Center;
            var prevSize = _lastFocusedBounds.Maximum - _lastFocusedBounds.Minimum;
            var nextSize = bounds.Maximum - bounds.Minimum;

            float centerDeltaSq = (prevCenter - nextCenter).LengthSquared();
            float sizeDeltaSq = (prevSize - nextSize).LengthSquared();

            return centerDeltaSq < 1e-6f && sizeDeltaSq < 1e-6f;
        }
    }
}
