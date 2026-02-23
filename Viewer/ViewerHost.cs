using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Windows;
using System.Windows.Media;
using Media3D = System.Windows.Media.Media3D;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Owns all Helix SharpDX GPU resources: EffectsManager, Camera, and scene root.
    /// Create once per window; call Dispose() when the window closes.
    /// </summary>
    public sealed class ViewerHost : IDisposable
    {
        // ── Public Helix objects bound by the XAML ──────────────────────────
        public EffectsManager EffectsManager { get; }
        public PerspectiveCamera Camera { get; }

        // ── Scene groups ────────────────────────────────────────────────────
        /// <summary>Root group for IFC geometry (added in Stage 2).</summary>
        public GroupModel3D IfcRoot { get; } = new GroupModel3D();

        /// <summary>Root group for Revit geometry (added in Stage 3).</summary>
        public GroupModel3D RevitRoot { get; } = new GroupModel3D();

        // ── Default camera position ──────────────────────────────────────────
        private static readonly Media3D.Point3D  DefaultPosition  = new Media3D.Point3D(10, 10, 10);
        private static readonly Media3D.Vector3D DefaultLookDir   = new Media3D.Vector3D(-1, -1, -1);
        private static readonly Media3D.Vector3D DefaultUpDir     = new Media3D.Vector3D(0, 1, 0);

        private bool _disposed;

        // ── Constructor ──────────────────────────────────────────────────────
        public ViewerHost()
        {
            EffectsManager = new DefaultEffectsManager();

            Camera = new PerspectiveCamera
            {
                Position          = DefaultPosition,
                LookDirection     = DefaultLookDir,
                UpDirection       = DefaultUpDir,
                FieldOfView       = 45,
                NearPlaneDistance = 0.01,
                // Use a large but finite far plane. Infinity causes depth-buffer
                // precision loss and defeats GPU early-Z rejection.
                // Updated to scene-diagonal * 3 after geometry loads via FitView().
                FarPlaneDistance  = 5000.0
            };

            SessionLogger.Info("ViewerHost created — EffectsManager and Camera initialized.");
        }

        // ── Camera helpers ───────────────────────────────────────────────────

        /// <summary>
        /// Switch near-plane distance for walk-through mode.
        /// Walk mode needs a very small near plane (0.001 m) so geometry is not
        /// clipped when the camera is close to or inside a surface.
        /// Orbit/inspect mode can use a slightly larger value (0.01 m) which gives
        /// better depth-buffer precision at a distance.
        /// </summary>
        public void SetWalkMode(bool walking)
        {
            Camera.NearPlaneDistance = walking ? 0.001 : 0.01;
            SessionLogger.Info($"NearPlane → {Camera.NearPlaneDistance} ({(walking ? "walk" : "orbit")})");
        }

        /// <summary>Animate the camera back to the default position smoothly.</summary>
        public void ResetCamera()
        {
            // AnimateTo(position, direction, upDir, animationTimeMs)
            Camera.AnimateTo(DefaultPosition, DefaultLookDir, DefaultUpDir, 300);
            SessionLogger.Info("Camera reset to default.");
        }

        /// <summary>
        /// Frame the given bounding box in the camera view.
        /// Call after loading geometry to fit the scene.
        /// </summary>
        public void FitView(BoundingBox box)
        {
            // Check for degenerate (empty) bounding box
            if (box.Maximum == box.Minimum) return;

            var center   = box.Center;
            var radius   = (box.Maximum - box.Minimum).Length() * 0.5f;
            var distance = radius / (float)Math.Tan(Camera.FieldOfView * 0.5 * Math.PI / 180.0) * 1.5f;

            var targetPos = new Media3D.Point3D(
                center.X + distance * 0.6,
                center.Y + distance * 0.6,
                center.Z + distance * 0.6);

            var lookDelta = new Media3D.Point3D(center.X, center.Y, center.Z) - targetPos;

            Camera.AnimateTo(targetPos,
                new Media3D.Vector3D(lookDelta.X, lookDelta.Y, lookDelta.Z),
                DefaultUpDir,
                400);

            // Set far plane to scene diagonal * 3 so depth buffer precision is
            // maximised while still covering the full scene from any view angle.
            double diagonal = (box.Maximum - box.Minimum).Length();
            Camera.FarPlaneDistance = Math.Max(500.0, diagonal * 3.0);

            SessionLogger.Info($"FitView — center ({center.X:F1},{center.Y:F1},{center.Z:F1})  radius {radius:F1}  farPlane {Camera.FarPlaneDistance:F0}");
        }

        // ── Scene helpers ────────────────────────────────────────────────────
        /// <summary>Initialise the scene: lights and IFC/Revit root groups.</summary>
        public void BuildTestScene(GroupModel3D sceneGroup)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // ── Directional lights ──────────────────────────────────────────
            // Sun: strong white from upper-left-front — primary key light.
            var sunLight = new DirectionalLight3D
            {
                Direction = new Media3D.Vector3D(-1, -2, -1.5),
                Color = Colors.White
            };
            // Fill: reduced to 40% of previous brightness so the sun↔shadow
            // gradient is steeper; slight blue tint differentiates it from the sun.
            var fillLight = new DirectionalLight3D
            {
                Direction = new Media3D.Vector3D(1, 1, 0.5),
                Color = System.Windows.Media.Color.FromRgb(35, 35, 50)
            };
            // Ambient: lowered from 40 to 15 (≈6% of full range).
            // With less ambient, surfaces pointing away from both lights are noticeably
            // darker, giving strong face-to-face contrast and clear shape perception.
            var ambientLight = new AmbientLight3D
            {
                Color = System.Windows.Media.Color.FromRgb(15, 15, 15)
            };
            sceneGroup.Children.Add(sunLight);
            sceneGroup.Children.Add(fillLight);
            sceneGroup.Children.Add(ambientLight);

            // ── IFC / Revit root groups ──────────────────────────────────────
            sceneGroup.Children.Add(IfcRoot);
            sceneGroup.Children.Add(RevitRoot);

            sw.Stop();
            SessionLogger.Info($"Scene initialised in {sw.ElapsedMilliseconds} ms.");
        }

        // ── IDisposable ──────────────────────────────────────────────────────
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            EffectsManager?.Dispose();
            SessionLogger.Info("ViewerHost disposed — DirectX resources released.");
        }
    }
}
