using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Xbim.Common;
using Xbim.Common.Geometry;
using Xbim.Ifc;
using Xbim.ModelGeometry.Scene;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Loads an IFC file using xBIM Essentials + Geometry Engine and produces
    /// Helix SharpDX scene objects ready to insert into a GroupModel3D.
    ///
    /// Threading: <see cref="LoadAsync"/> runs the heavy xBIM work on a background
    /// thread. The returned <see cref="IfcModel"/> must be added to the live scene
    /// on the UI thread by the caller.
    /// </summary>
    public static class IfcLoader
    {
        // ── Fallback colours by IFC product type ─────────────────────────────
        private static readonly XbimColour ColWall    = new XbimColour("Wall",   0.75f, 0.72f, 0.68f, 1f);
        private static readonly XbimColour ColSlab    = new XbimColour("Slab",   0.60f, 0.60f, 0.58f, 1f);
        private static readonly XbimColour ColColumn  = new XbimColour("Col",    0.65f, 0.62f, 0.58f, 1f);
        private static readonly XbimColour ColBeam    = new XbimColour("Beam",   0.65f, 0.62f, 0.58f, 1f);
        private static readonly XbimColour ColWindow  = new XbimColour("Win",    0.55f, 0.78f, 0.92f, 0.35f);
        private static readonly XbimColour ColDoor    = new XbimColour("Door",   0.82f, 0.70f, 0.55f, 1f);
        private static readonly XbimColour ColStair   = new XbimColour("Stair",  0.70f, 0.65f, 0.58f, 1f);
        private static readonly XbimColour ColRoof    = new XbimColour("Roof",   0.68f, 0.45f, 0.35f, 1f);
        private static readonly XbimColour ColDefault = new XbimColour("Other",  0.55f, 0.65f, 0.70f, 1f);

        // ── IFC types to skip (spaces, openings, annotations) ────────────────
        private static readonly HashSet<string> IgnoredTypes = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "IfcSpace", "IfcOpeningElement", "IfcVirtualElement",
            "IfcAnnotation", "IfcGrid", "IfcSite", "IfcBuilding", "IfcBuildingStorey"
        };

        // ── Public API ───────────────────────────────────────────────────────

        /// <summary>
        /// Opens and tessellates <paramref name="filePath"/> on a background thread.
        /// Returns an <see cref="IfcModel"/> whose <c>SceneGroup</c> is populated but
        /// not yet attached to the live Helix scene.
        /// </summary>
        public static Task<IfcModel> LoadAsync(string filePath)
            => Task.Run(() => LoadCore(filePath));

        // ── Core loader (runs on background thread) ──────────────────────────
        private static IfcModel LoadCore(string filePath)
        {
            var fileName = System.IO.Path.GetFileName(filePath);
            SessionLogger.Info("IfcLoader: opening '" + fileName + "'");
            var sw = System.Diagnostics.Stopwatch.StartNew();

            using (var model = IfcStore.Open(filePath))
            {
                // Build or reuse the triangulation cache
                var context = new Xbim3DModelContext(model);
                context.CreateContext();

                // Colour map: StyleLabel → XbimColour (surface styles defined in the file)
                var colourMap = new XbimColourMap();
                colourMap.SetProductTypeColourMap(); // seed with IFC type defaults

                var sceneGroup = new GroupModel3D();
                var allBounds  = new List<BoundingBox>();
                int meshCount  = 0;
                int triCount   = 0;

                // Bucket geometry by colour to reduce Helix draw calls
                // Key = rounded RGBA, Value = (positions, normals, indices, isTransparent)
                var buckets = new Dictionary<ColourKey, Bucket>();

                foreach (XbimShapeInstance instance in context.ShapeInstances())
                {
                    // Skip unwanted product types
                    if (ShouldSkip(instance, model)) continue;

                    // Get the pre-tessellated geometry for this instance
                    XbimShapeGeometry geom;
                    try { geom = context.ShapeGeometry(instance); }
                    catch { continue; }

                    if (string.IsNullOrEmpty(geom.ShapeData)) continue;

                    // Parse the triangulated mesh (applies the instance transform)
                    var xMesh = new XbimMeshGeometry3D();
                    try
                    {
                        xMesh.Read(geom.ShapeData,
                                   (XbimMatrix3D?)instance.Transformation);
                    }
                    catch { continue; }

                    if (xMesh.PositionCount == 0) continue;

                    // Resolve colour for this shape instance
                    XbimColour colour = ResolveColour(instance, model, colourMap);
                    var key = new ColourKey(colour);

                    if (!buckets.TryGetValue(key, out Bucket bucket))
                    {
                        bucket = new Bucket(colour.IsTransparent);
                        buckets[key] = bucket;
                    }

                    int baseIdx = bucket.Positions.Count;

                    // IFC uses Z-up; convert to Helix Y-up: (X, Y, Z) → (X, Z, -Y)
                    foreach (XbimPoint3D p in xMesh.Positions)
                        bucket.Positions.Add(new Vector3((float)p.X, (float)p.Z, -(float)p.Y));

                    foreach (XbimVector3D n in xMesh.Normals)
                        bucket.Normals.Add(new Vector3((float)n.X, (float)n.Z, -(float)n.Y));

                    foreach (int i in xMesh.TriangleIndices)
                        bucket.Indices.Add(baseIdx + i);

                    meshCount++;
                }

                // Build one MeshGeometryModel3D per colour bucket
                foreach (KeyValuePair<ColourKey, Bucket> kv in buckets)
                {
                    ColourKey key    = kv.Key;
                    Bucket    bucket = kv.Value;

                    if (bucket.Indices.Count == 0) continue;

                    var helixGeom = new MeshGeometry3D
                    {
                        Positions = new Vector3Collection(bucket.Positions),
                        Normals   = new Vector3Collection(bucket.Normals),
                        Indices   = new IntCollection(bucket.Indices)
                    };

                    var mat = new PhongMaterial
                    {
                        DiffuseColor      = key.ToColor4(),
                        SpecularColor     = new Color4(0.15f, 0.15f, 0.15f, 1f),
                        SpecularShininess = 12f,
                    };

                    var mesh3d = new MeshGeometryModel3D
                    {
                        Geometry      = helixGeom,
                        Material      = mat,
                        IsTransparent = bucket.IsTransparent,
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

                sw.Stop();
                SessionLogger.Info("IfcLoader: " + meshCount + " elements, " +
                                   triCount + " triangles in " + sw.ElapsedMilliseconds + " ms.");

                return new IfcModel(filePath, sceneGroup, bounds, meshCount, triCount);
            }
        }

        // ── Type filter ──────────────────────────────────────────────────────
        private static bool ShouldSkip(XbimShapeInstance instance, IModel model)
        {
            try
            {
                IPersistEntity entity = model.Instances[instance.IfcProductLabel];
                return entity != null && IgnoredTypes.Contains(entity.ExpressType.Name);
            }
            catch { return false; }
        }

        // ── Colour resolution ────────────────────────────────────────────────
        private static XbimColour ResolveColour(XbimShapeInstance instance,
                                                IModel model,
                                                XbimColourMap colourMap)
        {
            // 1. Try the surface-style colour embedded in the IFC file
            if (instance.StyleLabel > 0)
            {
                try
                {
                    IPersistEntity styleEntity = model.Instances[instance.StyleLabel];
                    if (styleEntity is Xbim.Ifc4.Interfaces.IIfcSurfaceStyle ss)
                    {
                        foreach (var item in ss.Styles)
                        {
                            if (item is Xbim.Ifc4.Interfaces.IIfcSurfaceStyleShading shading)
                            {
                                var col = shading.SurfaceColour;
                                float alpha = 1f - (float)(shading.Transparency ?? 0);
                                return new XbimColour("s",
                                    (float)col.Red, (float)col.Green, (float)col.Blue, alpha);
                            }
                        }
                    }
                    // IFC2x3 path: same interface via adapter
                    if (styleEntity is Xbim.Ifc2x3.PresentationAppearanceResource.IfcSurfaceStyle ss2x3)
                    {
                        foreach (var item in ss2x3.Styles)
                        {
                            if (item is Xbim.Ifc2x3.PresentationAppearanceResource.IfcSurfaceStyleRendering rend)
                            {
                                var col = rend.SurfaceColour;
                                float alpha = rend.Transparency.HasValue
                                    ? 1f - (float)rend.Transparency.Value : 1f;
                                return new XbimColour("s",
                                    (float)col.Red, (float)col.Green, (float)col.Blue, alpha);
                            }
                        }
                    }
                }
                catch { /* fall through */ }
            }

            // 2. Fall back to product-type colour
            try
            {
                IPersistEntity entity = model.Instances[instance.IfcProductLabel];
                if (entity != null)
                {
                    string typeName = entity.ExpressType.Name;
                    if (colourMap.Contains(typeName))
                        return colourMap[typeName];
                    return GetFallbackColour(typeName);
                }
            }
            catch { /* fall through */ }

            return ColDefault;
        }

        private static XbimColour GetFallbackColour(string typeName)
        {
            if (typeName.IndexOf("Wall",   StringComparison.OrdinalIgnoreCase) >= 0) return ColWall;
            if (typeName.IndexOf("Slab",   StringComparison.OrdinalIgnoreCase) >= 0) return ColSlab;
            if (typeName.IndexOf("Column", StringComparison.OrdinalIgnoreCase) >= 0) return ColColumn;
            if (typeName.IndexOf("Beam",   StringComparison.OrdinalIgnoreCase) >= 0) return ColBeam;
            if (typeName.IndexOf("Window", StringComparison.OrdinalIgnoreCase) >= 0) return ColWindow;
            if (typeName.IndexOf("Door",   StringComparison.OrdinalIgnoreCase) >= 0) return ColDoor;
            if (typeName.IndexOf("Stair",  StringComparison.OrdinalIgnoreCase) >= 0) return ColStair;
            if (typeName.IndexOf("Roof",   StringComparison.OrdinalIgnoreCase) >= 0) return ColRoof;
            return ColDefault;
        }

        // ── Bucket: accumulates vertices for one colour ───────────────────────
        private sealed class Bucket
        {
            public readonly List<Vector3> Positions = new List<Vector3>();
            public readonly List<Vector3> Normals   = new List<Vector3>();
            public readonly List<int>     Indices   = new List<int>();
            public readonly bool          IsTransparent;
            public Bucket(bool transparent) { IsTransparent = transparent; }
        }

        // ── Colour key: rounded RGBA for bucketing ────────────────────────────
        private readonly struct ColourKey : IEquatable<ColourKey>
        {
            private readonly float _r, _g, _b, _a;

            public ColourKey(XbimColour c)
            {
                _r = Round(c.Red);
                _g = Round(c.Green);
                _b = Round(c.Blue);
                _a = Round(c.Alpha);
            }

            private static float Round(double v) => (float)Math.Round(v, 2);

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
}
