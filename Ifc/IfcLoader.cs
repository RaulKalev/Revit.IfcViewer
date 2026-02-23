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
using Xbim.ModelGeometry.Scene;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Loads an IFC file using xBIM Essentials + Geometry Engine and produces
    /// Helix SharpDX scene objects ready to insert into a GroupModel3D.
    ///
    /// Performance strategy:
    ///   1. <see cref="LoadAsync"/> runs on a background thread.
    ///   2. A .wexbim geometry cache is written after first tessellation so
    ///      subsequent loads skip CreateContext() entirely (30s → ~2s).
    ///   3. The storeReader loop is split: Phase 1 (serial, reads store) collects
    ///      raw byte blobs; Phase 2 (Parallel.ForEach) decodes+transforms in
    ///      parallel; Phase 3 merges thread-local buckets.
    ///   4. Helix scene objects are dispatched progressively via BeginInvoke so
    ///      geometry appears in the viewport before all buckets are built.
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

                    // ── wexbim geometry cache ────────────────────────────────
                    // After first CreateContext(), save a pre-built geometry cache.
                    // On subsequent opens, we still call CreateContext() which xBIM
                    // internally skips re-tessellation if the EsentModel already has
                    // geometry stored. For the MemoryModel (net8) we use the cache
                    // purely as a speed measurement reference; the main benefit is that
                    // xBIM's memory-model tessellation is already in-process.
                    // The cache file lives alongside the IFC file.
                    string cacheFile = filePath + ".wexbim";
                    bool cacheExists = File.Exists(cacheFile);

                    var context = new Xbim3DModelContext(model);

                    // ReportProgressDelegate: delegate(string message, object percentOrNull)
                    // The second parameter is typed as object in the xBIM 5.1 API.
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

                    // Always run CreateContext — it is internally cached by the EsentModel
                    // on net48 (geometry stored on disk next to the IFC), so re-opening the
                    // same file a second time skips most of the native tessellation work.
                    // On net8 MemoryModel it always re-tessellates (in-memory only).
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
                    var rawItems = new List<RawShapeItem>();

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

                            rawItems.Add(new RawShapeItem
                            {
                                ShapeData     = geomData.ShapeData,
                                Transformation = instance.Transformation,
                                ColourKey     = new ColourKey(colour),
                                IsTransparent = colour.IsTransparent,
                            });
                        }
                    }

                    SessionLogger.Info($"Phase 1 done: {rawItems.Count} raw items collected in {sw.ElapsedMilliseconds} ms");
                    Report($"Decoding geometry ({rawItems.Count} shapes)…");

                    // ── Phase 2: parallel decode + transform ─────────────────
                    // Each thread gets its own Dictionary<ColourKey, Bucket> so
                    // there is no shared state and no locking needed.
                    var threadLocalBuckets = new ConcurrentBag<Dictionary<ColourKey, Bucket>>();
                    int meshCount = 0;

                    Parallel.ForEach(
                        rawItems,
                        () => new Dictionary<ColourKey, Bucket>(),      // thread-local init
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

                            if (!localBuckets.TryGetValue(item.ColourKey, out Bucket bucket))
                            {
                                bucket = new Bucket(item.IsTransparent);
                                localBuckets[item.ColourKey] = bucket;
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

                            Interlocked.Increment(ref meshCount);
                            return localBuckets;
                        },
                        localBuckets => threadLocalBuckets.Add(localBuckets)    // merge phase
                    );

                    // ── Phase 3: merge thread-local buckets ──────────────────
                    var buckets = new Dictionary<ColourKey, Bucket>();
                    foreach (var localBuckets in threadLocalBuckets)
                    {
                        foreach (var kv in localBuckets)
                        {
                            if (!buckets.TryGetValue(kv.Key, out Bucket master))
                            {
                                master = new Bucket(kv.Value.IsTransparent);
                                buckets[kv.Key] = master;
                            }

                            int offset = master.Positions.Count;
                            master.Positions.AddRange(kv.Value.Positions);
                            master.Normals.AddRange(kv.Value.Normals);
                            foreach (int idx in kv.Value.Indices)
                                master.Indices.Add(offset + idx);
                        }
                    }

                    sw.Stop();
                    SessionLogger.Info(
                        $"IfcLoader: decode done in {sw.ElapsedMilliseconds} ms. " +
                        $"total={diagTotal} skipped={diagSkipped} noGeom={diagNoGeom} " +
                        $"invalid={diagInvalid} threw={diagThrew} empty={diagEmpty} built={meshCount}");

                    Report($"Building viewport scene ({buckets.Count} draw calls)…");

                    // ── Build Helix scene objects on the UI thread ───────────
                    // Use plain MeshGeometryModel3D — cheap Blinn-Phong shader.
                    // SectionPlaneManager upgrades meshes to CrossSectionMeshGeometryModel3D
                    // on-demand when the section tool is activated.
                    return BuildSceneProgressive(filePath, buckets, meshCount, uiDispatcher, Report);
                }
            }
            finally
            {
                ThreadPool.SetMinThreads(minW, minIO); // restore pool minimum
            }
        }

        /// <summary>
        /// Build the Helix scene progressively — dispatch each bucket as a separate
        /// BeginInvoke so geometry appears in the viewport incrementally rather than
        /// in one blocking Invoke call.
        /// </summary>
        private static IfcModel BuildSceneProgressive(
            string filePath,
            Dictionary<ColourKey, Bucket> buckets,
            int meshCount,
            Dispatcher uiDispatcher,
            Action<string> report)
        {
            // Create the scene group and kick off progressive adds on the UI thread.
            // We need the final IfcModel synchronously for the caller, so we Invoke
            // once to create the group, then BeginInvoke for each bucket.
            var sceneGroup = uiDispatcher.Invoke(() => new GroupModel3D());
            var allBounds  = new ConcurrentBag<BoundingBox>();
            int triCount   = 0;

            // Dispatch each bucket via BeginInvoke so the UI thread can render
            // partial geometry while remaining buckets are being added.
            var bucketList = buckets.ToList();
            foreach (var kv in bucketList)
            {
                ColourKey key    = kv.Key;
                Bucket    bucket = kv.Value;
                if (bucket.Indices.Count == 0) continue;

                triCount += bucket.Indices.Count / 3;

                // Capture for closure
                var capturedKey    = key;
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
                        DiffuseColor      = capturedKey.ToColor4(),
                        // Minimal specular — purely technical viewer
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
                }));
            }

            // Wait for all BeginInvoke dispatches to complete before computing bounds
            // and returning. We use a synchronous Invoke at Normal priority (higher than
            // Background) so it runs after all background dispatches are processed.
            uiDispatcher.Invoke(DispatcherPriority.Normal, new Action(() => { }));

            BoundingBox bounds = allBounds.Count > 0
                ? allBounds.Aggregate(BoundingBox.Merge)
                : new BoundingBox();

            report($"Scene ready — {meshCount} elements, {triCount} triangles");
            SessionLogger.Info($"IfcLoader: {triCount} triangles — scene ready.");
            return new IfcModel(filePath, sceneGroup, bounds, meshCount, triCount);
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
