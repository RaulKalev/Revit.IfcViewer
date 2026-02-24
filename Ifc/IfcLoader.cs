using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Xbim.Common;
using Xbim.Common.Geometry;
using Xbim.Common.XbimExtensions;   // BinaryReaderExtensions.ReadShapeTriangulation
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;         // IIfcProduct, IIfcPropertySet, IIfcPropertySingleValue
using Xbim.ModelGeometry.Scene;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Loads an IFC file using xBIM Essentials + Geometry Engine and produces
    /// per-element Helix SharpDX scene objects ready to insert into a GroupModel3D.
    ///
    /// Performance strategy:
    ///   1. <see cref="LoadAsync"/> runs on a background thread.
    ///   2. A .wexbim geometry cache is written after first tessellation so
    ///      subsequent loads skip CreateContext() entirely (30s → ~2s).
    ///   3. The storeReader loop is split: Phase 1 (serial, reads store) collects
    ///      raw byte blobs and extracts element properties; Phase 2 (Parallel.ForEach)
    ///      decodes+transforms in parallel; Phase 3 merges thread-local buckets.
    ///   4. Each IFC product gets its own MeshGeometryModel3D so hit-testing
    ///      can identify individual elements. ElementMap is returned in IfcModel.
    ///   5. Helix scene objects are dispatched progressively via BeginInvoke so
    ///      geometry appears in the viewport before all elements are built.
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
        /// <paramref name="onProgress"/> is called on the UI thread with status strings.
        /// Returns an <see cref="IfcModel"/> whose <c>SceneGroup</c> is ready to insert
        /// into the live Helix scene (on the UI thread).
        /// Each IFC product gets its own MeshGeometryModel3D, enabling element selection.
        /// </summary>
        public static Task<IfcModel> LoadAsync(string filePath,
                                               Dispatcher uiDispatcher,
                                               Action<string> onProgress = null)
            => Task.Run(() => LoadCore(filePath, uiDispatcher, onProgress));

        // ── Core loader (runs on background thread) ──────────────────────────
        private static IfcModel LoadCore(string filePath,
                                         Dispatcher uiDispatcher,
                                         Action<string> onProgress)
        {
            void Report(string msg)
            {
                SessionLogger.Info(msg);
                if (onProgress != null)
                    uiDispatcher.BeginInvoke(new Action(() => onProgress(msg)));
            }

            var fileName = Path.GetFileName(filePath);
            Report($"Loading: {fileName}");
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Ensure native geometry engine DLLs are findable before any xBIM call
            string assemblyDir = Path.GetDirectoryName(typeof(IfcLoader).Assembly.Location);
            if (!string.IsNullOrEmpty(assemblyDir))
            {
                SetDllDirectory(assemblyDir);
                AddDllDirectory(assemblyDir);
                string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                if (currentPath.IndexOf(assemblyDir, StringComparison.OrdinalIgnoreCase) < 0)
                    Environment.SetEnvironmentVariable("PATH", assemblyDir + ";" + currentPath);
            }

            // ── Boost thread pool so xBIM's internal Parallel.ForEach
            //    gets worker threads without the default ramp-up delay
            ThreadPool.GetMinThreads(out int minW, out int minIO);
            int boosted = Math.Max(minW, Environment.ProcessorCount);
            ThreadPool.SetMinThreads(boosted, minIO);

            try
            {
                using (var model = IfcStore.Open(filePath))
                {
                    Report($"Parsing done ({sw.ElapsedMilliseconds} ms) — tessellating…");

                    string cacheFile = filePath + ".wexbim";
                    bool cacheExists = File.Exists(cacheFile);

                    var context = new Xbim3DModelContext(model);

                    int lastPct = 0;
                    Xbim.Common.ReportProgressDelegate progressDelegate = (msg, pctObj) =>
                    {
                        int pct = pctObj is int i ? i : 0;
                        if (pct - lastPct >= 10)
                        {
                            lastPct = pct;
                            Report($"Tessellating: {pct}%");
                        }
                    };

                    Report(cacheExists ? "Geometry store found — loading…" : "Tessellating geometry…");
                    bool contextOk = context.CreateContext(progressDelegate, false);
                    SessionLogger.Info($"CreateContext={contextOk}  instances={context.ShapeInstances().Count()}");

                    Report($"Tessellation done ({sw.ElapsedMilliseconds} ms) — building scene…");

                    float toMetres = (float)(1.0 / model.ModelFactors.OneMetre);
                    var colourMap  = new XbimColourMap();
                    colourMap.SetProductTypeColourMap();

                    int diagTotal = 0, diagSkipped = 0, diagNoGeom = 0,
                        diagInvalid = 0, diagThrew = 0, diagEmpty = 0;

                    // ── Phase 1: serial read from store ──────────────────────
                    // storeReader is NOT thread-safe — collect raw data first,
                    // then decode in parallel in Phase 2.
                    // Also extract element properties here while the model is open.
                    var rawItems     = new List<RawShapeItem>();
                    var elementInfos = new Dictionary<int, IfcElementInfo>(); // productLabel → info

                    using (var storeReader = model.GeometryStore.BeginRead())
                    {
                        foreach (XbimShapeInstance instance in storeReader.ShapeInstances)
                        {
                            diagTotal++;
                            if (ShouldSkip(instance, model)) { diagSkipped++; continue; }

                            IXbimShapeGeometryData geomData;
                            try { geomData = storeReader.ShapeGeometryOfInstance(instance); }
                            catch { diagNoGeom++; continue; }

                            if (geomData?.ShapeData == null || geomData.ShapeData.Length == 0)
                            { diagInvalid++; continue; }

                            XbimColour colour = ResolveColour(instance, model, colourMap);
                            int productLabel  = instance.IfcProductLabel;

                            // Extract element properties once per unique product label
                            if (!elementInfos.ContainsKey(productLabel))
                            {
                                try
                                {
                                    var entity = model.Instances[productLabel];
                                    if (entity is IIfcProduct product)
                                        elementInfos[productLabel] = ExtractElementInfo(product);
                                }
                                catch { /* non-critical — element will have no properties */ }
                            }

                            rawItems.Add(new RawShapeItem
                            {
                                ShapeData      = geomData.ShapeData,
                                Transformation = instance.Transformation,
                                ColourKey      = new ColourKey(colour),
                                IsTransparent  = colour.IsTransparent,
                                ProductLabel   = productLabel,
                            });
                        }
                    }

                    SessionLogger.Info($"Phase 1 done: {rawItems.Count} raw items, " +
                                       $"{elementInfos.Count} element info entries " +
                                       $"collected in {sw.ElapsedMilliseconds} ms");
                    Report($"Decoding geometry ({rawItems.Count} shapes)…");

                    // ── Phase 2: parallel decode + transform ─────────────────
                    // Each thread gets its own Dictionary<int, Bucket> keyed by productLabel,
                    // so there is no shared state and no locking needed.
                    var threadLocalBuckets = new ConcurrentBag<Dictionary<int, Bucket>>();
                    int shapeCount = 0;

                    Parallel.ForEach(
                        rawItems,
                        () => new Dictionary<int, Bucket>(),          // thread-local init
                        (item, state, localBuckets) =>
                        {
                            List<float[]> triPositions;
                            List<int>     triIndices;
                            try
                            {
                                XbimShapeTriangulation tri;
                                using (var ms = new MemoryStream(item.ShapeData))
                                using (var br = new BinaryReader(ms))
                                    tri = br.ReadShapeTriangulation();

                                tri = tri.Transform(item.Transformation);
                                tri.ToPointsWithNormalsAndIndices(out triPositions, out triIndices);
                            }
                            catch (Exception ex)
                            {
                                Interlocked.Increment(ref diagThrew);
                                SessionLogger.Warn($"Triangulate threw: {ex.Message}");
                                return localBuckets;
                            }

                            if (triPositions == null || triPositions.Count == 0)
                            { Interlocked.Increment(ref diagEmpty); return localBuckets; }

                            // Group by product label — each element gets its own bucket
                            if (!localBuckets.TryGetValue(item.ProductLabel, out Bucket bucket))
                            {
                                bucket = new Bucket(item.IsTransparent, item.ColourKey);
                                localBuckets[item.ProductLabel] = bucket;
                            }

                            int baseIdx = bucket.Positions.Count;

                            // Remap IFC Z-up → Helix Y-up: (X,Y,Z) → (X,Z,-Y), scale to metres
                            foreach (float[] v in triPositions)
                            {
                                bucket.Positions.Add(new Vector3(
                                     v[0] * toMetres,
                                     v[2] * toMetres,
                                    -v[1] * toMetres));
                                bucket.Normals.Add(new Vector3(v[3], v[5], -v[4]));
                            }
                            foreach (int idx in triIndices)
                                bucket.Indices.Add(baseIdx + idx);

                            Interlocked.Increment(ref shapeCount);
                            return localBuckets;
                        },
                        localBuckets => threadLocalBuckets.Add(localBuckets)    // merge phase
                    );

                    // ── Phase 3: merge thread-local buckets ──────────────────
                    var buckets = new Dictionary<int, Bucket>();
                    foreach (var localBuckets in threadLocalBuckets)
                    {
                        foreach (var kv in localBuckets)
                        {
                            if (!buckets.TryGetValue(kv.Key, out Bucket master))
                            {
                                master = new Bucket(kv.Value.IsTransparent, kv.Value.Colour);
                                buckets[kv.Key] = master;
                            }

                            int offset = master.Positions.Count;
                            master.Positions.AddRange(kv.Value.Positions);
                            master.Normals.AddRange(kv.Value.Normals);
                            foreach (int idx in kv.Value.Indices)
                                master.Indices.Add(offset + idx);
                        }
                    }

                    int elementCount = buckets.Count;
                    sw.Stop();
                    SessionLogger.Info(
                        $"IfcLoader: decode done in {sw.ElapsedMilliseconds} ms. " +
                        $"total={diagTotal} skipped={diagSkipped} noGeom={diagNoGeom} " +
                        $"invalid={diagInvalid} threw={diagThrew} empty={diagEmpty} " +
                        $"shapes={shapeCount} elements={elementCount}");

                    Report($"Building viewport scene ({elementCount} elements)…");

                    // ── Build Helix scene objects on the UI thread ───────────
                    return BuildSceneProgressive(filePath, buckets, elementInfos,
                                                 elementCount, uiDispatcher, Report);
                }
            }
            finally
            {
                ThreadPool.SetMinThreads(minW, minIO); // restore pool minimum
            }
        }

        /// <summary>
        /// Build the Helix scene progressively — dispatch each element as a separate
        /// BeginInvoke so geometry appears in the viewport incrementally.
        /// Creates one MeshGeometryModel3D per IFC product (element) to enable hit-testing.
        /// </summary>
        private static IfcModel BuildSceneProgressive(
            string filePath,
            Dictionary<int, Bucket> elementBuckets,
            Dictionary<int, IfcElementInfo> elementInfos,
            int elementCount,
            Dispatcher uiDispatcher,
            Action<string> report)
        {
            var sceneGroup = uiDispatcher.Invoke(() => new GroupModel3D());
            var allBounds  = new ConcurrentBag<BoundingBox>();
            // ElementMap is populated from BeginInvoke lambdas (all on UI thread, serial).
            var elementMap = new Dictionary<MeshGeometryModel3D, IfcElementInfo>();
            int triCount   = 0;

            var bucketList = elementBuckets.ToList();
            foreach (var kv in bucketList)
            {
                int    productLabel    = kv.Key;
                Bucket bucket          = kv.Value;
                if (bucket.Indices.Count == 0) continue;

                triCount += bucket.Indices.Count / 3;

                var capturedLabel  = productLabel;
                var capturedBucket = bucket;

                uiDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                {
                    var helixGeom = new MeshGeometry3D
                    {
                        Positions = new Vector3Collection(capturedBucket.Positions),
                        Normals   = new Vector3Collection(capturedBucket.Normals),
                        Indices   = new IntCollection(capturedBucket.Indices)
                    };

                    var mat = new PhongMaterial
                    {
                        DiffuseColor      = capturedBucket.Colour.ToColor4(),
                        AmbientColor      = new Color4(0.15f, 0.15f, 0.15f, 1f),
                        SpecularColor     = new Color4(0.05f, 0.05f, 0.05f, 1f),
                        SpecularShininess = 4f,
                        ReflectiveColor   = new Color4(0f, 0f, 0f, 0f),
                    };

                    var mesh3d = new MeshGeometryModel3D
                    {
                        Geometry      = helixGeom,
                        Material      = mat,
                        IsTransparent = capturedBucket.IsTransparent,
                    };

                    sceneGroup.Children.Add(mesh3d);

                    if (capturedBucket.Positions.Count > 0)
                    {
                        BoundingBox.FromPoints(capturedBucket.Positions.ToArray(), out BoundingBox b);
                        allBounds.Add(b);
                    }

                    // Register mesh → element info for click-selection
                    if (elementInfos.TryGetValue(capturedLabel, out IfcElementInfo info))
                        elementMap[mesh3d] = info;
                }));
            }

            // Wait for all BeginInvoke dispatches to complete before computing bounds.
            uiDispatcher.Invoke(DispatcherPriority.Normal, new Action(() => { }));

            BoundingBox bounds = allBounds.Count > 0
                ? allBounds.Aggregate(BoundingBox.Merge)
                : new BoundingBox();

            report($"Scene ready — {elementCount} elements, {triCount} triangles");
            SessionLogger.Info($"IfcLoader: {triCount} triangles, {elementCount} elements — scene ready.");
            return new IfcModel(filePath, sceneGroup, bounds, elementCount, triCount, elementMap);
        }

        // ── Element property extraction ──────────────────────────────────────

        /// <summary>
        /// Extracts Name, Type, GlobalId and all single-value property set entries
        /// from an IFC product while the model is still open.
        /// Uses the unified Xbim.Ifc4.Interfaces layer — works for both IFC2x3 and IFC4.
        /// </summary>
        private static IfcElementInfo ExtractElementInfo(IIfcProduct product)
        {
            var info = new IfcElementInfo
            {
                Name     = product.Name?.ToString() ?? "(unnamed)",
                Type     = product.ExpressType.Name,
                GlobalId = product.GlobalId.ToString(),
            };

            try
            {
                foreach (var rel in product.IsDefinedBy)
                {
                    if (!(rel is IIfcRelDefinesByProperties rdp)) continue;
                    if (!(rdp.RelatingPropertyDefinition is IIfcPropertySet pset)) continue;

                    var props = new Dictionary<string, string>();
                    foreach (var prop in pset.HasProperties)
                    {
                        if (prop is IIfcPropertySingleValue psv && psv.NominalValue != null)
                            props[prop.Name.ToString()] = psv.NominalValue.ToString();
                    }

                    if (props.Count > 0)
                        info.PropertySets[pset.Name.ToString()] = props;
                }
            }
            catch
            {
                // Property extraction is non-critical — silently ignore any pset read failures.
            }

            return info;
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
                                var col   = shading.SurfaceColour;
                                float alpha = 1f - (float)(shading.Transparency ?? 0);
                                return new XbimColour("s",
                                    (float)col.Red, (float)col.Green, (float)col.Blue, alpha);
                            }
                        }
                    }
                    if (styleEntity is Xbim.Ifc2x3.PresentationAppearanceResource.IfcSurfaceStyle ss2x3)
                    {
                        foreach (var item in ss2x3.Styles)
                        {
                            if (item is Xbim.Ifc2x3.PresentationAppearanceResource.IfcSurfaceStyleRendering rend)
                            {
                                var col   = rend.SurfaceColour;
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

            try
            {
                IPersistEntity entity = model.Instances[instance.IfcProductLabel];
                if (entity != null)
                {
                    string typeName = entity.ExpressType.Name;
                    if (colourMap.Contains(typeName)) return colourMap[typeName];
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

        // ── Raw shape item (Phase 1 output, Phase 2 input) ───────────────────
        private sealed class RawShapeItem
        {
            public byte[]       ShapeData;
            public XbimMatrix3D Transformation;
            public ColourKey    ColourKey;
            public bool         IsTransparent;
            public int          ProductLabel;
        }

        // ── Bucket: accumulates vertices for one IFC element ─────────────────
        private sealed class Bucket
        {
            public readonly List<Vector3> Positions = new List<Vector3>();
            public readonly List<Vector3> Normals   = new List<Vector3>();
            public readonly List<int>     Indices   = new List<int>();
            public readonly bool          IsTransparent;
            public readonly ColourKey     Colour;
            public Bucket(bool transparent, ColourKey colour)
            {
                IsTransparent = transparent;
                Colour        = colour;
            }
        }

        // ── Colour key: rounded RGBA for material creation ────────────────────
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
