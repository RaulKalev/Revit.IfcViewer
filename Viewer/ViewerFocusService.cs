using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Applies camera focus operations. When the scene's element handles are
    /// available (<see cref="HandlesProvider"/>), focusing picks a view direction
    /// with a clear line of sight to the target — preferring directions close to
    /// the user's current one — instead of blindly keeping the current
    /// orientation, which often parked the camera behind a wall.
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

        /// <summary>
        /// Supplies every element handle in the scene (all loaded models) for
        /// line-of-sight tests. Occlusion-aware focusing is skipped when null.
        /// </summary>
        public Func<IReadOnlyList<ElementHandle>> HandlesProvider { get; set; }

        public ViewerFocusService(PerspectiveCamera camera)
        {
            _camera = camera ?? throw new ArgumentNullException(nameof(camera));
        }

        /// <summary>Focus the camera on bounds, keeping the current orientation.</summary>
        public bool FocusByWorldBounds(BoundingBox bounds)
            => FocusCore(bounds, null);

        /// <summary>
        /// Focus the camera on one element, rotating to a direction from which the
        /// element is actually visible (not hidden behind other geometry).
        /// </summary>
        public bool FocusOnElement(ElementHandle target)
        {
            if (target == null) return false;
            return FocusCore(target.Bounds, target);
        }

        private bool FocusCore(BoundingBox bounds, ElementHandle target)
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

            var baseDir = new Vector3(
                (float)lookDirection.X, (float)lookDirection.Y, (float)lookDirection.Z);

            Vector3 chosenDir = baseDir;
            IReadOnlyList<ElementHandle> handles =
                target != null ? HandlesProvider?.Invoke() : null;
            if (handles != null)
            {
                chosenDir = ChooseVisibleDirection(handles, target, center, baseDir,
                                                   (float)distance, out bool clearView);

                // Fully enclosed target (e.g. inside a small room): no direction is
                // clear at framing distance, so step the camera in along the best
                // direction to just inside the nearest blocker.
                if (!clearView)
                {
                    float free = MaxFreeDistance(handles, target, center, chosenDir, (float)distance);
                    double minDistance = Math.Max(0.3, radius * 1.05);
                    distance = Math.Max(Math.Min(distance, free - 0.25), minDistance);
                }
            }

            var newLookDirection = new Media3D.Vector3D(
                chosenDir.X * distance, chosenDir.Y * distance, chosenDir.Z * distance);
            var targetPoint = new Media3D.Point3D(center.X, center.Y, center.Z);
            var newPosition = targetPoint - newLookDirection;

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

        // ── Line-of-sight direction search ────────────────────────────────────

        /// <summary>
        /// Pick the look direction whose camera position sees the target with the
        /// fewest obstructions. Candidates fan out from the current direction
        /// (least disruptive first), then an elevated 45° ring for looking over
        /// walls into rooms. Occlusion is tested conservatively against element
        /// bounding boxes, which is cheap and never misses a real blocker.
        /// </summary>
        private static Vector3 ChooseVisibleDirection(
            IReadOnlyList<ElementHandle> handles, ElementHandle target,
            Vector3 center, Vector3 baseDir, float distance, out bool clearView)
        {
            Vector3 best = baseDir;
            int bestScore = int.MaxValue;
            clearView = false;

            foreach (Vector3 dir in CandidateDirections(baseDir))
            {
                int score = CountOccluders(handles, target, center, dir, distance);
                if (score == 0)
                {
                    clearView = true;
                    return dir;
                }
                if (score < bestScore)
                {
                    bestScore = score;
                    best = dir;
                }
            }
            return best;
        }

        /// <summary>
        /// Distance from the target's centre along the camera direction to the
        /// nearest blocking element — the farthest the camera can sit while
        /// keeping an unobstructed view. Boxes that contain the centre itself
        /// (host walls, enclosing shells) are unavoidable and ignored.
        /// </summary>
        private static float MaxFreeDistance(
            IReadOnlyList<ElementHandle> handles, ElementHandle target,
            Vector3 center, Vector3 dir, float distance)
        {
            var ray = new Ray(center, -dir);
            float free = distance;
            for (int i = 0; i < handles.Count; i++)
            {
                ElementHandle h = handles[i];
                if (h.IsHidden || ReferenceEquals(h, target)) continue;
                if (h.Bounds.Intersects(ref ray, out float t) && t > 0.01f && t < free)
                    free = t;
            }
            return free;
        }

        private static IEnumerable<Vector3> CandidateDirections(Vector3 baseDir)
        {
            // Horizontal ring at the current pitch, nearest yaw offsets first.
            double[] yawOffsetsDeg = { 0, 35, -35, 70, -70, 110, -110, 145, -145, 180 };
            foreach (double deg in yawOffsetsDeg)
                yield return RotateAroundY(baseDir, deg);

            // Elevated ring: look down at ~45° — sees over walls into rooms.
            Vector3 horiz = new Vector3(baseDir.X, 0, baseDir.Z);
            if (horiz.LengthSquared() < 1e-6f) horiz = new Vector3(1, 0, 0);
            horiz.Normalize();
            const float c45 = 0.70710678f;
            var elevated = new Vector3(horiz.X * c45, -c45, horiz.Z * c45);
            foreach (double deg in yawOffsetsDeg)
                yield return RotateAroundY(elevated, deg);

            // Last resort: steep top-down view.
            yield return Vector3.Normalize(new Vector3(horiz.X * 0.25f, -1f, horiz.Z * 0.25f));
        }

        private static Vector3 RotateAroundY(Vector3 v, double degrees)
        {
            double rad = degrees * Math.PI / 180.0;
            float cos = (float)Math.Cos(rad), sin = (float)Math.Sin(rad);
            return new Vector3(v.X * cos + v.Z * sin, v.Y, -v.X * sin + v.Z * cos);
        }

        /// <summary>
        /// Number of elements whose bounds block the segment from the candidate
        /// camera position to the target's bounds. Zero means a clear view.
        /// </summary>
        private static int CountOccluders(
            IReadOnlyList<ElementHandle> handles, ElementHandle target,
            Vector3 center, Vector3 dir, float distance)
        {
            Vector3 position = center - dir * distance;
            var ray = new Ray(position, Vector3.Normalize(center - position));

            // Unobstructed sight only needs to reach the target's bounds, not its centre.
            BoundingBox targetBounds = target.Bounds;
            if (!targetBounds.Intersects(ref ray, out float tTarget))
                tTarget = distance;

            int count = 0;
            for (int i = 0; i < handles.Count; i++)
            {
                ElementHandle h = handles[i];
                if (h.IsHidden || ReferenceEquals(h, target)) continue;
                if (h.Bounds.Intersects(ref ray, out float t) && t < tTarget - 0.01f)
                    count++;
            }
            return count;
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
