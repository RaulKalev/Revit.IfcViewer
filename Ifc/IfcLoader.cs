using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;
using Xbim.Common;
using Xbim.Common.Geometry;
using Xbim.Common.XbimExtensions;   // BinaryReaderExtensions.ReadShapeTriangulation
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
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr AddDllDirectory(string lpPathName);

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
        /// WPF DependencyObjects (GroupModel3D, MeshGeometryModel3D, PhongMaterial) are
        /// marshalled to <paramref name="uiDispatcher"/> so they are owned by the UI thread.
        /// Returns an <see cref="IfcModel"/> whose <c>SceneGroup</c> is ready to insert
        /// into the live Helix scene (on the UI thread).
        /// </summary>
        public static Task<IfcModel> LoadAsync(string filePath, Dispatcher uiDispatcher)
            => Task.Run(() => LoadCore(filePath, uiDispatcher));

        // ── Core loader (runs on background thread) ──────────────────────────
        private static IfcModel LoadCore(string filePath, Dispatcher uiDispatcher)
        {
            var fileName = System.IO.Path.GetFileName(filePath);
            SessionLogger.Info("IfcLoader: opening '" + fileName + "'");
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Ensure Xbim.Geometry.Engine32/64.dll (unmanaged) is findable before ANY
            // xBIM call. Must happen before IfcStore.Open — the engine is loaded lazily
            // at first geometry operation, and SetDllDirectory only affects future loads.
            string assemblyDir = Path.GetDirectoryName(typeof(IfcLoader).Assembly.Location);
            if (!string.IsNullOrEmpty(assemblyDir))
            {
                SetDllDirectory(assemblyDir);
                AddDllDirectory(assemblyDir);

                // Also prepend to PATH so LoadLibrary's fallback search finds it.
                string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                if (currentPath.IndexOf(assemblyDir, StringComparison.OrdinalIgnoreCase) < 0)
                    Environment.SetEnvironmentVariable("PATH", assemblyDir + ";" + currentPath);

                bool eng32     = File.Exists(Path.Combine(assemblyDir, "Xbim.Geometry.Engine32.dll"));
                bool eng64     = File.Exists(Path.Combine(assemblyDir, "Xbim.Geometry.Engine64.dll"));
                bool interop   = File.Exists(Path.Combine(assemblyDir, "Xbim.Geometry.Engine.Interop.dll"));
                bool esent     = File.Exists(Path.Combine(assemblyDir, "Xbim.IO.Esent.dll"));
                bool esentInterop = File.Exists(Path.Combine(assemblyDir, "Esent.Interop.dll"));
                SessionLogger.Info($"IfcLoader: DLL path → {assemblyDir}");
                SessionLogger.Info($"  Engine32={eng32} Engine64={eng64} Interop={interop} Esent={esent} EsentInterop={esentInterop}");
            }

            using (var model = IfcStore.Open(filePath))
            {
                SessionLogger.Info($"IfcLoader: model type={model.GetType().Name}  GeometryStore={model.GeometryStore?.GetType().Name ?? "null"}");

                // Tessellate the model into the in-memory geometry store.
                // Pass adjustWcs=false to keep the original coordinate system.
                // null progress = no callbacks.
                var context = new Xbim3DModelContext(model);
                bool contextOk = context.CreateContext(null, false);
                SessionLogger.Info($"IfcLoader: CreateContext={contextOk}  instances={context.ShapeInstances().Count()}  geometries={context.ShapeGeometries().Count()}");

                // Colour map: StyleLabel → XbimColour (surface styles defined in the file)
                var colourMap = new XbimColourMap();
                colourMap.SetProductTypeColourMap(); // seed with IFC type defaults

                int meshCount = 0;
                int diagTotal = 0, diagSkipped = 0, diagNoGeom = 0,
                    diagInvalid = 0, diagThrew = 0, diagEmpty = 0;

                // Bucket geometry by colour to reduce Helix draw calls.
                // All work here is plain CLR types — safe on a background thread.
                // Key = rounded RGBA, Value = (positions, normals, indices, isTransparent)
                var buckets = new Dictionary<ColourKey, Bucket>();

                // Read tessellated geometry back from the store via IGeometryStoreReader.
                // ShapeGeometryOfInstance returns IXbimShapeGeometryData whose ShapeData
                // property is byte[] — the raw XbimShapeTriangulation binary blob written
                // by CreateContext(). BinaryReaderExtensions.ReadShapeTriangulation() decodes
                // it without invoking the native geometry engine again.
                using (var storeReader = model.GeometryStore.BeginRead())
                {
                    foreach (XbimShapeInstance instance in storeReader.ShapeInstances)
                    {
                        diagTotal++;

                        // Skip unwanted product types
                        if (ShouldSkip(instance, model)) { diagSkipped++; continue; }

                        // Get the raw byte[] blob from the store
                        IXbimShapeGeometryData geomData;
                        try
                        {
                            geomData = storeReader.ShapeGeometryOfInstance(instance);
                        }
                        catch { diagNoGeom++; continue; }

                        if (geomData == null || geomData.ShapeData == null || geomData.ShapeData.Length == 0)
                        { diagInvalid++; continue; }

                        // Deserialize: byte[] → XbimShapeTriangulation (pure managed read)
                        List<float[]> triPositions;
                        List<int>     triIndices;
                        try
                        {
                            XbimShapeTriangulation tri;
                            using (var ms = new MemoryStream(geomData.ShapeData))
                            using (var br = new BinaryReader(ms))
                            {
                                tri = br.ReadShapeTriangulation();
                            }
                            // Apply the instance world transform
                            tri = tri.Transform(instance.Transformation);
                            tri.ToPointsWithNormalsAndIndices(out triPositions, out triIndices);
                        }
                        catch (Exception ex)
                        {
                            diagThrew++;
                            if (diagThrew <= 3) SessionLogger.Warn("Triangulate threw: " + ex.Message);
                            continue;
                        }

                        if (triPositions == null || triPositions.Count == 0)
                        { diagEmpty++; continue; }

                        // Resolve colour for this shape instance
                        XbimColour colour = ResolveColour(instance, model, colourMap);
                        var key = new ColourKey(colour);

                        if (!buckets.TryGetValue(key, out Bucket bucket))
                        {
                            bucket = new Bucket(colour.IsTransparent);
                            buckets[key] = bucket;
                        }

                        int baseIdx = bucket.Positions.Count;

                        // triPositions[i] = float[6]{ X, Y, Z, NX, NY, NZ } in IFC Z-up space.
                        // Remap IFC Z-up → Helix Y-up: (X, Y, Z) → (X, Z, -Y)
                        foreach (float[] v in triPositions)
                        {
                            bucket.Positions.Add(new Vector3(v[0],  v[2], -v[1]));
                            bucket.Normals  .Add(new Vector3(v[3],  v[5], -v[4]));
                        }

                        foreach (int idx in triIndices)
                            bucket.Indices.Add(baseIdx + idx);

                        meshCount++;
                    }
                } // storeReader.Dispose()

                sw.Stop();
                SessionLogger.Info(
                    $"IfcLoader: tessellation done in {sw.ElapsedMilliseconds} ms. " +
                    $"total={diagTotal} skipped={diagSkipped} noGeom={diagNoGeom} " +
                    $"invalid={diagInvalid} threw={diagThrew} emptyTri={diagEmpty} built={meshCount}");
                SessionLogger.Info("IfcLoader: xBIM tessellation done — " + meshCount +
                                   " instances in " + sw.ElapsedMilliseconds + " ms." +
                                   " Building Helix scene on UI thread…");

                // ── Build WPF/Helix objects on the UI thread ─────────────────
                // GroupModel3D, MeshGeometryModel3D and PhongMaterial are all
                // DependencyObjects and must be created/owned by the UI thread.
                return uiDispatcher.Invoke(() =>
                {
                    var sceneGroup = new GroupModel3D();
                    var allBounds  = new List<BoundingBox>();
                    int triCount   = 0;

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

                    SessionLogger.Info("IfcLoader: " + triCount + " triangles — scene ready.");
                    return new IfcModel(filePath, sceneGroup, bounds, meshCount, triCount);
                });
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
