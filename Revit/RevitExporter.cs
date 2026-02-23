using Autodesk.Revit.DB;
using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Exports tessellated geometry from the active Revit 3D view using
    /// the Revit API's <see cref="CustomExporter"/> pipeline.
    ///
    /// Threading: <see cref="Export"/> must be called on the Revit API thread
    /// (i.e. inside an IExternalCommand or ExternalEvent). All Helix WPF objects
    /// are then marshalled to <paramref name="uiDispatcher"/> so they are owned
    /// by the WPF UI thread.
    /// </summary>
    public static class RevitExporter
    {
        // ── Public API ───────────────────────────────────────────────────────

        /// <summary>
        /// Tessellates <paramref name="view"/> and returns a <see cref="RevitModel"/>
        /// whose <c>SceneGroup</c> is ready to add to <c>ViewerHost.RevitRoot</c>.
        /// </summary>
        /// <param name="doc">Active Revit document.</param>
        /// <param name="view">3D view to export (must be a View3D).</param>
        /// <param name="uiDispatcher">WPF dispatcher used to create Helix DependencyObjects.</param>
        public static RevitModel Export(Document doc, View3D view, Dispatcher uiDispatcher)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            SessionLogger.Info($"RevitExporter: exporting view '{view.Name}'");

            var context = new RevitExportContext(doc);

            // CustomExporter drives the geometry pipeline.
            // includeFaces=false (we get triangulated meshes via OnPolymesh)
            // includeLinearObjects=false
            using (var exporter = new CustomExporter(doc, context))
            {
                exporter.IncludeGeometricObjects = false; // suppress detail curves etc.
                exporter.ShouldStopOnError       = false;
                exporter.Export(view);
            }

            sw.Stop();
            SessionLogger.Info(
                $"RevitExporter: done in {sw.ElapsedMilliseconds} ms. " +
                $"Buckets={context.Buckets.Count}  Faces={context.FaceCount}");

            // ── Build Helix scene on the WPF UI thread ───────────────────────
            return uiDispatcher.Invoke(() => BuildScene(context, view.Name));
        }

        // ── Scene builder (runs on UI thread) ────────────────────────────────
        private static RevitModel BuildScene(RevitExportContext ctx, string viewName)
        {
            var sceneGroup = new GroupModel3D();
            var allBounds  = new List<BoundingBox>();
            int triCount   = 0;

            foreach (var kv in ctx.Buckets)
            {
                var bucket = kv.Value;
                if (bucket.Indices.Count == 0) continue;

                var helixGeom = new MeshGeometry3D
                {
                    Positions = new Vector3Collection(bucket.Positions),
                    Normals   = new Vector3Collection(bucket.Normals),
                    Indices   = new IntCollection(bucket.Indices)
                };

                var mat = new PhongMaterial
                {
                    DiffuseColor      = kv.Key.ToColor4(),
                    SpecularColor     = new Color4(0.15f, 0.15f, 0.15f, 1f),
                    SpecularShininess = 12f,
                };

                // CrossSectionMeshGeometryModel3D extends MeshGeometryModel3D and
                // exposes Plane1..8 for the section-plane tool.
                var mesh3d = new CrossSectionMeshGeometryModel3D
                {
                    Geometry      = helixGeom,
                    Material      = mat,
                    IsTransparent = kv.Key.Alpha < 0.99f,
                };

                sceneGroup.Children.Add(mesh3d);
                triCount += bucket.Indices.Count / 3;

                if (bucket.Positions.Count > 0)
                {
                    BoundingBox.FromPoints(bucket.Positions.ToArray(), out BoundingBox b);
                    allBounds.Add(b);
                }
            }

            BoundingBox bounds = allBounds.Count > 0
                ? allBounds.Aggregate(BoundingBox.Merge)
                : new BoundingBox();

            SessionLogger.Info($"RevitExporter: {triCount} triangles — scene ready.");
            return new RevitModel(viewName, sceneGroup, bounds, ctx.Buckets.Count, triCount);
        }
    }

    // ── IExportContext implementation ─────────────────────────────────────────

    internal sealed class RevitExportContext : IExportContext
    {
        // ── Output ───────────────────────────────────────────────────────────
        public readonly Dictionary<ColourKey, Bucket> Buckets = new Dictionary<ColourKey, Bucket>();
        public int FaceCount { get; private set; }

        // ── Per-element state ────────────────────────────────────────────────
        private readonly Document _doc;
        private readonly Stack<Transform> _transformStack = new Stack<Transform>();
        private Color4 _currentColor = new Color4(0.7f, 0.7f, 0.7f, 1f);
        private bool _skipElement;

        // Category filter: skip non-3D annotation categories
        private static readonly HashSet<BuiltInCategory> SkippedCategories = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_Cameras,
            BuiltInCategory.OST_RenderRegions,
            BuiltInCategory.OST_SectionBox,
        };

        public RevitExportContext(Document doc)
        {
            _doc = doc;
        }

        // ── IExportContext lifecycle ──────────────────────────────────────────

        public bool Start()
        {
            _transformStack.Clear();
            _transformStack.Push(Transform.Identity);
            return true;
        }

        public void Finish() { }

        public bool IsCanceled() => false;

        // ── View ─────────────────────────────────────────────────────────────

        public RenderNodeAction OnViewBegin(ViewNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnViewEnd(ElementId elementId) { }

        // ── Element ──────────────────────────────────────────────────────────

        public RenderNodeAction OnElementBegin(ElementId elementId)
        {
            _skipElement = false;

            Element elem = _doc.GetElement(elementId);
            if (elem == null) { _skipElement = true; return RenderNodeAction.Skip; }

            // Skip non-solid categories
            if (elem.Category != null)
            {
                try
                {
                    var bic = (BuiltInCategory)(int)elem.Category.Id.Value;
                    if (SkippedCategories.Contains(bic))
                    { _skipElement = true; return RenderNodeAction.Skip; }
                }
                catch { /* non-built-in category — proceed */ }
            }

            // Default colour from category material
            _currentColor = CategoryToColor(elem);
            return RenderNodeAction.Proceed;
        }

        public void OnElementEnd(ElementId elementId)
        {
            _skipElement = false;
        }

        // ── Instances (linked models, families) ───────────────────────────────

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            _transformStack.Push(CurrentTransform.Multiply(node.GetTransform()));
            return RenderNodeAction.Proceed;
        }

        public void OnInstanceEnd(InstanceNode node)
        {
            if (_transformStack.Count > 1) _transformStack.Pop();
        }

        // ── Link ─────────────────────────────────────────────────────────────

        public RenderNodeAction OnLinkBegin(LinkNode node)
        {
            _transformStack.Push(CurrentTransform.Multiply(node.GetTransform()));
            return RenderNodeAction.Proceed;
        }

        public void OnLinkEnd(LinkNode node)
        {
            if (_transformStack.Count > 1) _transformStack.Pop();
        }

        // ── Face / material ───────────────────────────────────────────────────

        public RenderNodeAction OnFaceBegin(FaceNode node)
        {
            return RenderNodeAction.Proceed;
        }

        public void OnFaceEnd(FaceNode node) { }

        // ── Material ─────────────────────────────────────────────────────────

        public void OnMaterial(MaterialNode node)
        {
            if (_skipElement) return;

            float r = (float)(node.Color.Red   / 255.0);
            float g = (float)(node.Color.Green / 255.0);
            float b = (float)(node.Color.Blue  / 255.0);
            float a = (float)(1.0 - node.Transparency);

            _currentColor = new Color4(r, g, b, a);
        }

        // ── Polymesh (triangulated faces) ─────────────────────────────────────

        public void OnPolymesh(PolymeshTopology polymesh)
        {
            if (_skipElement) return;

            var pts   = polymesh.GetPoints();
            var norms = polymesh.GetNormals();
            var facets = polymesh.GetFacets();

            Transform xform = CurrentTransform;
            var key    = new ColourKey(_currentColor);

            if (!Buckets.TryGetValue(key, out Bucket bucket))
            {
                bucket = new Bucket();
                Buckets[key] = bucket;
            }

            int baseIdx = bucket.Positions.Count;
            bool hasNormals = norms != null && norms.Count == pts.Count;

            // Revit uses feet internally; convert to metres (* 0.3048) then
            // remap Revit Z-up → Helix Y-up: (X, Y, Z) → (X, Z, -Y)
            const float FT_TO_M = 0.3048f;

            for (int i = 0; i < pts.Count; i++)
            {
                var p = pts[i];
                XYZ tp = xform.OfPoint(p);
                bucket.Positions.Add(new Vector3(
                    (float)tp.X * FT_TO_M,
                    (float)tp.Z * FT_TO_M,
                   -(float)tp.Y * FT_TO_M));

                if (hasNormals)
                {
                    var n = norms[i];
                    XYZ tn = xform.OfVector(n).Normalize();
                    bucket.Normals.Add(new Vector3(
                        (float)tn.X,
                        (float)tn.Z,
                       -(float)tn.Y));
                }
                else
                {
                    bucket.Normals.Add(new Vector3(0, 1, 0));
                }
            }

            foreach (var facet in facets)
            {
                bucket.Indices.Add(baseIdx + facet.V1);
                bucket.Indices.Add(baseIdx + facet.V2);
                bucket.Indices.Add(baseIdx + facet.V3);
                FaceCount++;
            }
        }

        // ── RPC / light / curve — ignored ────────────────────────────────────

        public void OnRPC(RPCNode node) { }
        public void OnLight(LightNode node) { }

        // ── Helpers ──────────────────────────────────────────────────────────

        private Transform CurrentTransform
            => _transformStack.Count > 0 ? _transformStack.Peek() : Transform.Identity;

        private static Color4 CategoryToColor(Element elem)
        {
            try
            {
                if (elem.Category?.Material != null)
                {
                    var mat = elem.Category.Material;
                    return new Color4(
                        mat.Color.Red   / 255f,
                        mat.Color.Green / 255f,
                        mat.Color.Blue  / 255f,
                        1f);
                }
            }
            catch { /* ignore */ }

            // Very rough fallback by built-in category
            if (elem.Category != null)
            {
                try
                {
                    var bic = (BuiltInCategory)(int)elem.Category.Id.Value;
                    switch (bic)
                    {
                        case BuiltInCategory.OST_Walls:         return new Color4(0.75f, 0.72f, 0.68f, 1f);
                        case BuiltInCategory.OST_Floors:        return new Color4(0.60f, 0.60f, 0.58f, 1f);
                        case BuiltInCategory.OST_Ceilings:      return new Color4(0.85f, 0.85f, 0.82f, 1f);
                        case BuiltInCategory.OST_Roofs:         return new Color4(0.68f, 0.45f, 0.35f, 1f);
                        case BuiltInCategory.OST_StructuralColumns:
                        case BuiltInCategory.OST_StructuralFraming:
                                                                return new Color4(0.65f, 0.62f, 0.58f, 1f);
                        case BuiltInCategory.OST_Windows:       return new Color4(0.55f, 0.78f, 0.92f, 0.35f);
                        case BuiltInCategory.OST_Doors:         return new Color4(0.82f, 0.70f, 0.55f, 1f);
                        case BuiltInCategory.OST_Stairs:        return new Color4(0.70f, 0.65f, 0.58f, 1f);
                    }
                }
                catch { /* ignore */ }
            }

            return new Color4(0.55f, 0.65f, 0.70f, 1f);
        }
    }

    // ── Supporting types ──────────────────────────────────────────────────────

    internal sealed class Bucket
    {
        public readonly List<Vector3> Positions = new List<Vector3>();
        public readonly List<Vector3> Normals   = new List<Vector3>();
        public readonly List<int>     Indices   = new List<int>();
    }

    internal readonly struct ColourKey : IEquatable<ColourKey>
    {
        private readonly float _r, _g, _b, _a;

        public float Alpha => _a;

        public ColourKey(Color4 c)
        {
            _r = Round(c.Red);
            _g = Round(c.Green);
            _b = Round(c.Blue);
            _a = Round(c.Alpha);
        }

        private static float Round(float v) => (float)Math.Round(v, 2);

        public Color4 ToColor4() => new Color4(_r, _g, _b, _a);

        public bool Equals(ColourKey other)
            => _r == other._r && _g == other._g && _b == other._b && _a == other._a;

        public override bool Equals(object obj)
            => obj is ColourKey ck && Equals(ck);

        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + _r.GetHashCode();
                h = h * 31 + _g.GetHashCode();
                h = h * 31 + _b.GetHashCode();
                h = h * 31 + _a.GetHashCode();
                return h;
            }
        }
    }
}
