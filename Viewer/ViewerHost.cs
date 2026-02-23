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
                Position        = DefaultPosition,
                LookDirection   = DefaultLookDir,
                UpDirection     = DefaultUpDir,
                FieldOfView     = 45,
                NearPlaneDistance = 0.01,
                FarPlaneDistance  = double.PositiveInfinity
            };

            SessionLogger.Info("ViewerHost created — EffectsManager and Camera initialized.");
        }

        // ── Camera helpers ───────────────────────────────────────────────────
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

            SessionLogger.Info($"FitView — center ({center.X:F1},{center.Y:F1},{center.Z:F1})  radius {radius:F1}");
        }

        // ── Scene helpers ────────────────────────────────────────────────────
        /// <summary>Build the static test scene shown at Stage 0→1: axis lines + cube.</summary>
        public void BuildTestScene(GroupModel3D sceneGroup)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // ── Directional lights ──────────────────────────────────────────
            var sunLight = new DirectionalLight3D
            {
                Direction = new Media3D.Vector3D(-1, -2, -1.5),
                Color = Colors.White
            };
            var fillLight = new DirectionalLight3D
            {
                Direction = new Media3D.Vector3D(1, 1, 0.5),
                Color = System.Windows.Media.Color.FromRgb(80, 80, 120)
            };
            var ambientLight = new AmbientLight3D
            {
                Color = System.Windows.Media.Color.FromRgb(40, 40, 40)
            };
            sceneGroup.Children.Add(sunLight);
            sceneGroup.Children.Add(fillLight);
            sceneGroup.Children.Add(ambientLight);

            // ── Ground grid ─────────────────────────────────────────────────
            var grid = BuildGrid(20, 1f);
            sceneGroup.Children.Add(grid);

            // ── Axis indicator (X=red, Y=green, Z=blue) ─────────────────────
            sceneGroup.Children.Add(BuildAxisLine(new Vector3(0, 0, 0), new Vector3(3, 0, 0), Colors.Red));
            sceneGroup.Children.Add(BuildAxisLine(new Vector3(0, 0, 0), new Vector3(0, 3, 0), Colors.LimeGreen));
            sceneGroup.Children.Add(BuildAxisLine(new Vector3(0, 0, 0), new Vector3(0, 0, 3), Colors.DodgerBlue));

            // ── Test cube (1×1×1 at origin) ─────────────────────────────────
            var cubeMesh = new MeshGeometry3D();
            BuildUnitCube(cubeMesh);

            var cubeMat = new PhongMaterial
            {
                DiffuseColor  = new Color4(0.2f, 0.6f, 0.9f, 1f),
                SpecularColor = new Color4(0.5f, 0.5f, 0.5f, 1f),
                SpecularShininess = 32f
            };

            var cubeModel = new MeshGeometryModel3D
            {
                Geometry = cubeMesh,
                Material = cubeMat,
                Transform = new Media3D.TranslateTransform3D(0, 0.5, 0) // sit on grid
            };
            sceneGroup.Children.Add(cubeModel);

            // ── Also add the IFC / Revit root groups ────────────────────────
            sceneGroup.Children.Add(IfcRoot);
            sceneGroup.Children.Add(RevitRoot);

            sw.Stop();
            SessionLogger.Info($"Test scene built in {sw.ElapsedMilliseconds} ms.");
        }

        // ── Internal geometry helpers ────────────────────────────────────────
        private static LineGeometryModel3D BuildAxisLine(Vector3 from, Vector3 to, System.Windows.Media.Color color)
        {
            var positions = new Vector3Collection { from, to };
            var indices   = new IntCollection { 0, 1 };
            var geom = new LineGeometry3D { Positions = positions, Indices = indices };

            return new LineGeometryModel3D
            {
                Geometry  = geom,
                Color     = color,
                Thickness = 2.0
            };
        }

        private static LineGeometryModel3D BuildGrid(int halfExtent, float step)
        {
            var positions = new Vector3Collection();
            var indices   = new IntCollection();
            int idx = 0;

            var gridColor = System.Windows.Media.Color.FromArgb(100, 80, 80, 80);

            for (int i = -halfExtent; i <= halfExtent; i++)
            {
                float fi = i * step;
                float ext = halfExtent * step;

                // Lines parallel to Z
                positions.Add(new Vector3(fi, 0, -ext));
                positions.Add(new Vector3(fi, 0,  ext));
                indices.Add(idx++);
                indices.Add(idx++);

                // Lines parallel to X
                positions.Add(new Vector3(-ext, 0, fi));
                positions.Add(new Vector3( ext, 0, fi));
                indices.Add(idx++);
                indices.Add(idx++);
            }

            var geom = new LineGeometry3D { Positions = positions, Indices = indices };
            return new LineGeometryModel3D { Geometry = geom, Color = gridColor, Thickness = 0.5 };
        }

        private static void BuildUnitCube(MeshGeometry3D mesh)
        {
            // 8 vertices of a unit cube centred at origin (Y from 0 to 1 after transform)
            var p = new Vector3[]
            {
                new Vector3(-0.5f, -0.5f, -0.5f), // 0 bottom-back-left
                new Vector3( 0.5f, -0.5f, -0.5f), // 1 bottom-back-right
                new Vector3( 0.5f,  0.5f, -0.5f), // 2 top-back-right
                new Vector3(-0.5f,  0.5f, -0.5f), // 3 top-back-left
                new Vector3(-0.5f, -0.5f,  0.5f), // 4 bottom-front-left
                new Vector3( 0.5f, -0.5f,  0.5f), // 5 bottom-front-right
                new Vector3( 0.5f,  0.5f,  0.5f), // 6 top-front-right
                new Vector3(-0.5f,  0.5f,  0.5f), // 7 top-front-left
            };

            // 6 faces × 4 verts each (quads split into 2 triangles)
            int[][] faces = new[]
            {
                new[]{0,3,2,1}, // back   (-Z)
                new[]{4,5,6,7}, // front  (+Z)
                new[]{0,1,5,4}, // bottom (-Y)
                new[]{3,7,6,2}, // top    (+Y)
                new[]{0,4,7,3}, // left   (-X)
                new[]{1,2,6,5}, // right  (+X)
            };

            var positions = new Vector3Collection();
            var normals   = new Vector3Collection();
            var indices   = new IntCollection();

            int baseIdx = 0;
            Vector3[] faceNormals =
            {
                new Vector3( 0, 0,-1),
                new Vector3( 0, 0, 1),
                new Vector3( 0,-1, 0),
                new Vector3( 0, 1, 0),
                new Vector3(-1, 0, 0),
                new Vector3( 1, 0, 0),
            };

            for (int f = 0; f < faces.Length; f++)
            {
                var face = faces[f];
                var n    = faceNormals[f];
                foreach (int vi in face)
                {
                    positions.Add(p[vi]);
                    normals.Add(n);
                }
                // two triangles: 0-1-2, 0-2-3
                indices.Add(baseIdx);     indices.Add(baseIdx + 1); indices.Add(baseIdx + 2);
                indices.Add(baseIdx);     indices.Add(baseIdx + 2); indices.Add(baseIdx + 3);
                baseIdx += 4;
            }

            mesh.Positions = positions;
            mesh.Normals   = normals;
            mesh.Indices   = indices;
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
