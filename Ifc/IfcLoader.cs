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
using Xbim.Ifc4.Interfaces;         // IIfcProduct + relation/property interfaces
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

        // ── IFC types to skip (spatial containers, void/features, annotations) ──
        // Matched against the IFC entity CLR type hierarchy and ExpressType names.
        private static readonly HashSet<string> IgnoredTypeNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            // Spatial hierarchy / non-physical space containers
            "IfcSpatialStructureElement",
            "IfcSpatialElement",
            "IfcSpace",
            "IfcSpatialZone",
            "IfcExternalSpatialElement",
            "IfcSite",
            "IfcBuilding",
            "IfcBuildingStorey",

            // Voids and feature modifiers
            "IfcOpeningElement",
            "IfcOpeningStandardCase",
            "IfcFeatureElement",
            "IfcFeatureElementSubtraction",
            "IfcVoidingFeature",
            "IfcSurfaceFeature",
            "IfcVirtualElement",

            // Reference / annotation constructs
            "IfcAnnotation",
            "IfcGrid",
            "IfcGridAxis",
        };

        private static readonly HashSet<string> OpeningTypeNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "IfcOpeningElement",
            "IfcOpeningStandardCase",
            "IfcVoidingFeature",
            "IfcFeatureElementSubtraction",
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

            // ── Fast path: skip all xBIM work if processed cache exists ──────
            // The cache stores the final per-element geometry + element infos
            // (after wall-layer merging and opening cuts), so loading it is just
            // a binary file read — typically <100 ms vs 10-30 s from scratch.
            if (TryLoadCache(filePath, out var cachedBuckets, out var cachedInfos))
            {
                sw.Stop();
                SessionLogger.Info($"IfcLoader: cache hit in {sw.ElapsedMilliseconds} ms ({cachedBuckets.Count} elements).");
                Report($"Loaded from cache ({cachedBuckets.Count} elements).");
                return BuildSceneProgressive(filePath, cachedBuckets, cachedInfos,
                                             cachedBuckets.Count, uiDispatcher, Report);
            }

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
                    var openingItems = new List<OpeningShapeItem>();
                    var elementInfos = new Dictionary<int, IfcElementInfo>(); // productLabel → info
                    var wallLayerRelationGroups = BuildWallLayerRelationGroups(model); // wall host label -> layered children
                    var voidRelations = BuildVoidRelationMap(model, wallLayerRelationGroups); // opening label → set of host+child product labels
                    var productsWithIncludedGeom = new HashSet<int>();
                    HashSet<int> productsWithEngineCuts = null;

                    using (var storeReader = model.GeometryStore.BeginRead())
                    {
                        // Some products can have multiple geometry representations:
                        //  - OpeningsAndAdditionsIncluded  (engine-cut, with voids applied)
                        //  - OpeningsAndAdditionsExcluded  (raw body, no boolean cuts)
                        //  - OpeningsAndAdditionsOnly      (feature solids, not product body)
                        // Prefer "Included" per product so the engine's boolean cuts are used.
                        //
                        // productsWithIncludedGeom: ALL products with at least one "Included"
                        //   shape — used by ShouldSkipByRepresentation to suppress "Excluded"
                        //   duplicates.
                        // productsWithEngineCuts: products that have BOTH "Included" AND
                        //   "Excluded" representations — the engine performed actual boolean
                        //   operations.  Used to skip the AABB triangle-cull fallback and to
                        //   discard layer parts (the engine-cut host is authoritative).
                        var productsWithExcludedGeom = new HashSet<int>();
                        foreach (XbimShapeInstance instance in storeReader.ShapeInstances)
                        {
                            if (instance.RepresentationType
                                == XbimGeometryRepresentationType.OpeningsAndAdditionsIncluded)
                                productsWithIncludedGeom.Add(instance.IfcProductLabel);
                            else if (instance.RepresentationType
                                == XbimGeometryRepresentationType.OpeningsAndAdditionsExcluded)
                                productsWithExcludedGeom.Add(instance.IfcProductLabel);
                        }
                        productsWithEngineCuts = new HashSet<int>(productsWithIncludedGeom);
                        productsWithEngineCuts.IntersectWith(productsWithExcludedGeom);

                        foreach (XbimShapeInstance instance in storeReader.ShapeInstances)
                        {
                            diagTotal++;
                            IPersistEntity entity = null;
                            try { entity = model.Instances[instance.IfcProductLabel]; }
                            catch { }

                            if (IsOpeningLikeEntity(entity))
                            {
                                IXbimShapeGeometryData openingGeom;
                                try { openingGeom = storeReader.ShapeGeometryOfInstance(instance); }
                                catch { diagNoGeom++; continue; }
                                if (openingGeom?.ShapeData == null || openingGeom.ShapeData.Length == 0)
                                { diagInvalid++; continue; }

                                // Resolve which host product(s) this opening cuts.
                                HashSet<int> hostLabels = null;
                                voidRelations.TryGetValue(instance.IfcProductLabel, out hostLabels);

                                openingItems.Add(new OpeningShapeItem
                                {
                                    ShapeData = openingGeom.ShapeData,
                                    Transformation = instance.Transformation,
                                    HostProductLabels = hostLabels,
                                });
                                continue;
                            }

                            if (ShouldSkip(instance, entity, productsWithIncludedGeom))
                            {
                                diagSkipped++;
                                continue;
                            }

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
                                       $"{openingItems.Count} opening items, " +
                                       $"{elementInfos.Count} element info entries, " +
                                       $"{wallLayerRelationGroups.Count} wall relation groups, " +
                                       $"{productsWithEngineCuts.Count} engine-cut products " +
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

                    int mergedWallLayersByRelations = CollapseWallLayersByRelationGroups(
                        buckets,
                        elementInfos,
                        wallLayerRelationGroups,
                        productsWithEngineCuts);
                    int mergedWallLayersByGeometry = CollapseWallLayersIntoCombinedWalls(
                        buckets,
                        elementInfos);
                    int mergedWallLayers = mergedWallLayersByRelations + mergedWallLayersByGeometry;

                    var openingVolumes = BuildOpeningVolumes(openingItems, toMetres);
                    var layerBounds = GetLayerCandidateBounds(buckets, elementInfos);
                    var layeredOpeningVolumes = ExpandOpeningVolumesForLayers(
                        openingVolumes,
                        layerBounds);
                    int cutWallCount = ApplyOpeningVolumesToLayers(
                        buckets, elementInfos, layeredOpeningVolumes,
                        wallLayerRelationGroups);

                    int elementCount = buckets.Count;
                    sw.Stop();
                    SessionLogger.Info(
                        $"IfcLoader: decode done in {sw.ElapsedMilliseconds} ms. " +
                        $"total={diagTotal} skipped={diagSkipped} noGeom={diagNoGeom} " +
                        $"invalid={diagInvalid} threw={diagThrew} empty={diagEmpty} " +
                        $"mergedWallLayers={mergedWallLayers} " +
                        $"(relations={mergedWallLayersByRelations}, geometry={mergedWallLayersByGeometry}) " +
                        $"openings={openingVolumes.Count} layeredOpenings={layeredOpeningVolumes.Count} layeredCuts={cutWallCount} " +
                        $"shapes={shapeCount} elements={elementCount}");

                    Report($"Building viewport scene ({elementCount} elements)…");

                    // Persist processed geometry so subsequent opens use the fast path.
                    TrySaveCacheAsync(filePath, buckets, elementInfos);

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

            // Wait for all BeginInvoke(Background) dispatches to complete.
            // ContextIdle priority (3) is LOWER than Background (4), so this only
            // executes after all queued Background callbacks have finished — ensuring
            // both allBounds and elementMap are fully populated before we return.
            uiDispatcher.Invoke(DispatcherPriority.ContextIdle, new Action(() => { }));

            BoundingBox bounds = allBounds.Count > 0
                ? allBounds.Aggregate(BoundingBox.Merge)
                : new BoundingBox();

            report($"Scene ready — {elementCount} elements, {triCount} triangles");
            SessionLogger.Info($"IfcLoader: {triCount} triangles, {elementCount} elements — scene ready.");
            return new IfcModel(filePath, sceneGroup, bounds, elementCount, triCount, elementMap);
        }

        // ── Processed geometry cache ──────────────────────────────────────────
        //
        // After a full IFC load (tessellate + wall-layer merge + opening cuts),
        // the final per-element geometry and element-info tables are written to a
        // compact binary file in %LocalAppData%.  Subsequent opens of the same
        // file skip all xBIM work and read the binary data directly.
        //
        // Cache invalidation: the file name encodes a FNV-1a hash of the IFC path
        // plus the file's last-write ticks.  A changed file → different ticks →
        // different name → automatic miss.  Bumping CacheFormatVersion forces a
        // rebuild of all caches regardless of timestamps.
        // ─────────────────────────────────────────────────────────────────────
        private const int CacheFormatVersion = 7;
        private static readonly byte[] CacheMagic
            = System.Text.Encoding.ASCII.GetBytes("IFCVC");

        private static string GetCacheDir()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RKTools", "IfcViewer", "Cache");
            Directory.CreateDirectory(dir);
            return dir;
        }

        /// <summary>
        /// FNV-1a 32-bit hash of the normalised IFC file path.
        /// Stable across processes (no framework randomisation).
        /// </summary>
        private static uint ComputePathHash(string ifcFilePath)
        {
            string normalised = Path.GetFullPath(ifcFilePath).ToLowerInvariant();
            uint h = 2166136261u;
            foreach (char c in normalised) { h ^= (byte)c; h *= 16777619u; }
            return h;
        }

        /// <summary>
        /// Returns the full path of the binary cache file for the given IFC file.
        /// The name encodes a FNV-1a hash of the normalised source path plus the
        /// file's last-write ticks, so it changes automatically when the file is
        /// updated.
        /// </summary>
        private static string GetCachePath(string ifcFilePath)
        {
            uint h     = ComputePathHash(ifcFilePath);
            long ticks = File.GetLastWriteTimeUtc(ifcFilePath).Ticks;
            return Path.Combine(GetCacheDir(), $"{h:X8}_{ticks:X16}.ifcvcache");
        }

        /// <summary>
        /// Deletes all cached versions of the given IFC file from the cache
        /// directory.  Call this before a forced reload so the next
        /// <see cref="LoadAsync"/> call performs a full rebuild.
        /// </summary>
        public static void InvalidateCache(string ifcFilePath)
        {
            try
            {
                uint   h      = ComputePathHash(ifcFilePath);
                string prefix = $"{h:X8}_";
                string dir    = GetCacheDir();
                foreach (string f in Directory.GetFiles(dir, prefix + "*.ifcvcache"))
                {
                    try { File.Delete(f); }
                    catch { /* best-effort */ }
                }
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Attempts to read processed geometry and element infos from the binary
        /// cache.  Returns false on any miss, version mismatch, or read error.
        /// </summary>
        private static bool TryLoadCache(
            string ifcFilePath,
            out Dictionary<int, Bucket> buckets,
            out Dictionary<int, IfcElementInfo> elementInfos)
        {
            buckets      = null;
            elementInfos = null;
            try
            {
                string path = GetCachePath(ifcFilePath);
                if (!File.Exists(path)) return false;

                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var br = new BinaryReader(fs, System.Text.Encoding.UTF8))
                {
                    // Validate magic + version
                    byte[] magic = br.ReadBytes(5);
                    for (int i = 0; i < 5; i++)
                        if (magic[i] != CacheMagic[i]) return false;
                    if (br.ReadInt32() != CacheFormatVersion) return false;

                    // ── Geometry ────────────────────────────────────────────
                    int elementCount = br.ReadInt32();
                    buckets = new Dictionary<int, Bucket>(elementCount);

                    for (int e = 0; e < elementCount; e++)
                    {
                        int   label       = br.ReadInt32();
                        float r = br.ReadSingle(), g = br.ReadSingle(),
                              b = br.ReadSingle(), a = br.ReadSingle();
                        bool  transparent = br.ReadBoolean();

                        var bucket = new Bucket(transparent, new ColourKey(r, g, b, a));

                        int posCount = br.ReadInt32();
                        bucket.Positions.Capacity = posCount;
                        bucket.Normals.Capacity   = posCount;
                        for (int i = 0; i < posCount; i++)
                            bucket.Positions.Add(new Vector3(
                                br.ReadSingle(), br.ReadSingle(), br.ReadSingle()));
                        for (int i = 0; i < posCount; i++)
                            bucket.Normals.Add(new Vector3(
                                br.ReadSingle(), br.ReadSingle(), br.ReadSingle()));

                        int idxCount = br.ReadInt32();
                        bucket.Indices.Capacity = idxCount;
                        for (int i = 0; i < idxCount; i++)
                            bucket.Indices.Add(br.ReadInt32());

                        buckets[label] = bucket;
                    }

                    // ── Element infos ────────────────────────────────────────
                    int infoCount = br.ReadInt32();
                    elementInfos = new Dictionary<int, IfcElementInfo>(infoCount);

                    for (int e = 0; e < infoCount; e++)
                    {
                        int label = br.ReadInt32();
                        var info  = new IfcElementInfo
                        {
                            Name     = br.ReadString(),
                            Type     = br.ReadString(),
                            GlobalId = br.ReadString(),
                        };

                        int psetCount = br.ReadInt32();
                        for (int p = 0; p < psetCount; p++)
                        {
                            string psetName  = br.ReadString();
                            int    propCount = br.ReadInt32();
                            var    props     = new Dictionary<string, string>(propCount);
                            for (int i = 0; i < propCount; i++)
                                props[br.ReadString()] = br.ReadString();
                            info.PropertySets[psetName] = props;
                        }
                        elementInfos[label] = info;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                SessionLogger.Warn($"IfcLoader: cache read failed ({ex.Message}) — rebuilding.");
                buckets      = null;
                elementInfos = null;
                return false;
            }
        }

        /// <summary>
        /// Serialises processed geometry and element infos to the binary cache
        /// file on a background thread.  Before writing, any stale cache files
        /// that share the same path-hash prefix are deleted.
        /// </summary>
        private static void TrySaveCacheAsync(
            string ifcFilePath,
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos)
        {
            // Compute target path on the caller thread (cheap + avoids race with
            // the IFC file being modified before the background thread runs).
            string targetPath = GetCachePath(ifcFilePath);

            Task.Run(() =>
            {
                try
                {
                    string cacheDir   = GetCacheDir();
                    string hashPrefix = Path.GetFileName(targetPath).Substring(0, 8);

                    // Remove stale entries for this source file before writing.
                    foreach (string old in Directory.GetFiles(
                        cacheDir, $"{hashPrefix}_*.ifcvcache"))
                    {
                        if (!string.Equals(old, targetPath,
                                StringComparison.OrdinalIgnoreCase))
                            File.Delete(old);
                    }

                    using (var fs = new FileStream(
                        targetPath, FileMode.Create,
                        FileAccess.Write, FileShare.None))
                    using (var bw = new BinaryWriter(fs, System.Text.Encoding.UTF8))
                    {
                        bw.Write(CacheMagic);
                        bw.Write(CacheFormatVersion);

                        // ── Geometry ────────────────────────────────────────
                        bw.Write(buckets.Count);
                        foreach (var kv in buckets)
                        {
                            bw.Write(kv.Key);
                            Color4 c = kv.Value.Colour.ToColor4();
                            bw.Write(c.Red); bw.Write(c.Green);
                            bw.Write(c.Blue); bw.Write(c.Alpha);
                            bw.Write(kv.Value.IsTransparent);

                            bw.Write(kv.Value.Positions.Count);
                            foreach (var p in kv.Value.Positions)
                            { bw.Write(p.X); bw.Write(p.Y); bw.Write(p.Z); }
                            foreach (var n in kv.Value.Normals)
                            { bw.Write(n.X); bw.Write(n.Y); bw.Write(n.Z); }

                            bw.Write(kv.Value.Indices.Count);
                            foreach (int idx in kv.Value.Indices)
                                bw.Write(idx);
                        }

                        // ── Element infos ─────────────────────────────────
                        bw.Write(elementInfos.Count);
                        foreach (var kv in elementInfos)
                        {
                            bw.Write(kv.Key);
                            bw.Write(kv.Value.Name     ?? "");
                            bw.Write(kv.Value.Type     ?? "");
                            bw.Write(kv.Value.GlobalId ?? "");

                            bw.Write(kv.Value.PropertySets.Count);
                            foreach (var pset in kv.Value.PropertySets)
                            {
                                bw.Write(pset.Key);
                                bw.Write(pset.Value.Count);
                                foreach (var prop in pset.Value)
                                {
                                    bw.Write(prop.Key);
                                    bw.Write(prop.Value ?? "");
                                }
                            }
                        }

                        long sizeKb = fs.Length / 1024;
                        SessionLogger.Info(
                            $"IfcLoader: cache saved ({sizeKb} KB) → " +
                            $"{Path.GetFileName(targetPath)}");
                    }
                }
                catch (Exception ex)
                {
                    SessionLogger.Warn($"IfcLoader: cache write failed: {ex.Message}");
                }
            });
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
        private static bool ShouldSkip(
            XbimShapeInstance instance,
            IPersistEntity entity,
            HashSet<int> productsWithIncludedGeom)
        {
            try
            {
                if (ShouldSkipByRepresentation(instance, productsWithIncludedGeom))
                    return true;

                return IsIgnoredEntity(entity);
            }
            catch { return false; }
        }

        private static bool ShouldSkipByRepresentation(
            XbimShapeInstance instance,
            HashSet<int> productsWithIncludedGeom)
        {
            var repType = instance.RepresentationType;

            // "Only" geometry is just feature solids (voids/additions), not product body.
            if (repType == XbimGeometryRepresentationType.OpeningsAndAdditionsOnly)
                return true;

            // If "Included" (engine-cut) geometry exists for this product, suppress
            // the "Excluded" (uncut) duplicate so openings stay visible.
            if (repType == XbimGeometryRepresentationType.OpeningsAndAdditionsExcluded
                && productsWithIncludedGeom != null
                && productsWithIncludedGeom.Contains(instance.IfcProductLabel))
                return true;

            return false;
        }

        private static bool IsOpeningLikeEntity(IPersistEntity entity)
        {
            return HasTypeInHierarchy(entity, OpeningTypeNames);
        }

        private static bool IsIgnoredEntity(IPersistEntity entity)
        {
            return HasTypeInHierarchy(entity, IgnoredTypeNames);
        }

        private static bool HasTypeInHierarchy(
            IPersistEntity entity,
            HashSet<string> typeNames)
        {
            if (entity == null || typeNames == null || typeNames.Count == 0) return false;

            // Check CLR type hierarchy first (catches derived IFC entities).
            Type clrType = entity.GetType();
            while (clrType != null && clrType != typeof(object))
            {
                if (typeNames.Contains(clrType.Name))
                    return true;
                clrType = clrType.BaseType;
            }

            // Fallback check using xBIM Express type metadata.
            string expressName = entity.ExpressType?.Name;
            if (!string.IsNullOrEmpty(expressName) && typeNames.Contains(expressName))
                return true;

            return false;
        }

        private static Dictionary<int, HashSet<int>> BuildWallLayerRelationGroups(IModel model)
        {
            var groups = new Dictionary<int, HashSet<int>>();
            if (model == null) return groups;

            try
            {
                foreach (IIfcRelAggregates rel in model.Instances.OfType<IIfcRelAggregates>())
                    RegisterWallLayerRelation(rel?.RelatingObject, rel?.RelatedObjects, groups);
            }
            catch (Exception ex)
            {
                SessionLogger.Warn($"IfcLoader: relation scan (IfcRelAggregates) failed: {ex.Message}");
            }

            try
            {
                foreach (IIfcRelNests rel in model.Instances.OfType<IIfcRelNests>())
                    RegisterWallLayerRelation(rel?.RelatingObject, rel?.RelatedObjects, groups);
            }
            catch (Exception ex)
            {
                SessionLogger.Warn($"IfcLoader: relation scan (IfcRelNests) failed: {ex.Message}");
            }

            return groups;
        }

        /// <summary>
        /// Builds a map from opening element labels to the set of host product
        /// labels they void.  For decomposed walls, the opening's host set
        /// includes both the host wall and its child layer parts.
        /// This enables targeted AABB fallback cuts (only cut the walls that
        /// actually have an opening, not all nearby walls).
        /// </summary>
        private static Dictionary<int, HashSet<int>> BuildVoidRelationMap(
            IModel model,
            Dictionary<int, HashSet<int>> wallLayerRelationGroups)
        {
            var map = new Dictionary<int, HashSet<int>>();
            if (model == null) return map;

            try
            {
                foreach (IIfcRelVoidsElement rel in model.Instances.OfType<IIfcRelVoidsElement>())
                {
                    if (rel?.RelatingBuildingElement == null
                        || rel.RelatedOpeningElement == null)
                        continue;

                    int openingLabel = (rel.RelatedOpeningElement as IPersistEntity)?.EntityLabel ?? 0;
                    int hostLabel    = (rel.RelatingBuildingElement as IPersistEntity)?.EntityLabel ?? 0;
                    if (openingLabel <= 0 || hostLabel <= 0) continue;

                    if (!map.TryGetValue(openingLabel, out HashSet<int> hosts))
                    {
                        hosts = new HashSet<int>();
                        map[openingLabel] = hosts;
                    }

                    hosts.Add(hostLabel);

                    // Also add all decomposed children of the host wall — the
                    // opening should cut through layer parts as well.
                    if (wallLayerRelationGroups != null
                        && wallLayerRelationGroups.TryGetValue(hostLabel, out HashSet<int> children))
                    {
                        foreach (int child in children)
                            hosts.Add(child);
                    }
                }
            }
            catch (Exception ex)
            {
                SessionLogger.Warn($"IfcLoader: IfcRelVoidsElement scan failed: {ex.Message}");
            }

            return map;
        }

        private static void RegisterWallLayerRelation(
            IIfcObjectDefinition relatingObject,
            IEnumerable<IIfcObjectDefinition> relatedObjects,
            Dictionary<int, HashSet<int>> groups)
        {
            if (groups == null || relatedObjects == null) return;

            int wallHostLabel = ResolveWallHostLabel(relatingObject, 6);
            if (wallHostLabel <= 0) return;

            if (!groups.TryGetValue(wallHostLabel, out HashSet<int> children))
            {
                children = new HashSet<int>();
                groups[wallHostLabel] = children;
            }

            foreach (IIfcObjectDefinition related in relatedObjects)
            {
                if (!(related is IPersistEntity relatedEntity)) continue;

                int childLabel = relatedEntity.EntityLabel;
                if (childLabel <= 0 || childLabel == wallHostLabel) continue;
                if (!ShouldApplyLayeredCut(relatedEntity.ExpressType?.Name)) continue;

                children.Add(childLabel);
            }
        }

        private static int ResolveWallHostLabel(IIfcObjectDefinition objectDefinition, int maxDepth)
        {
            if (!(objectDefinition is IPersistEntity)) return 0;

            var visited = new HashSet<int>();
            IIfcObjectDefinition current = objectDefinition;
            int depth = 0;

            while (current != null && depth <= maxDepth)
            {
                if (!(current is IPersistEntity currentEntity)) return 0;

                int label = currentEntity.EntityLabel;
                if (label <= 0 || !visited.Add(label)) return 0;

                if (IsWallLikeType(currentEntity.ExpressType?.Name))
                    return label;

                IIfcObjectDefinition next = null;

                if (current.Decomposes != null)
                {
                    foreach (IIfcRelDecomposes rel in current.Decomposes)
                    {
                        if (!TryGetRelatingObjectFromDecomposes(rel, out IIfcObjectDefinition relatingObject))
                            continue;
                        next = relatingObject;
                        break;
                    }
                }

                if (next == null && current.Nests != null)
                {
                    foreach (IIfcRelNests rel in current.Nests)
                    {
                        if (rel?.RelatingObject == null) continue;
                        next = rel.RelatingObject;
                        break;
                    }
                }

                current = next;
                depth++;
            }

            return 0;
        }

        private static bool IsWallLikeType(string ifcType)
        {
            return !string.IsNullOrWhiteSpace(ifcType)
                && ifcType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryGetRelatingObjectFromDecomposes(
            IIfcRelDecomposes rel,
            out IIfcObjectDefinition relatingObject)
        {
            relatingObject = null;
            if (rel == null) return false;

            if (rel is IIfcRelAggregates aggregates && aggregates.RelatingObject != null)
            {
                relatingObject = aggregates.RelatingObject;
                return true;
            }

            if (rel is IIfcRelNests nests && nests.RelatingObject != null)
            {
                relatingObject = nests.RelatingObject;
                return true;
            }

            return false;
        }

        private static int CollapseWallLayersByRelationGroups(
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos,
            Dictionary<int, HashSet<int>> wallLayerGroups,
            HashSet<int> productsWithEngineCuts)
        {
            if (buckets == null || buckets.Count == 0) return 0;
            if (wallLayerGroups == null || wallLayerGroups.Count == 0) return 0;

            int mergedCount = 0;
            foreach (var kv in wallLayerGroups)
            {
                int hostLabel = kv.Key;
                HashSet<int> childLabels = kv.Value;
                if (childLabels == null || childLabels.Count == 0) continue;

                // When the host wall has engine-cut geometry (both "Included"
                // AND "Excluded" representations), discard the layer parts —
                // the host already carries correct boolean cuts from the
                // XbimGeometryEngine and merging uncut layers would fill openings.
                if (productsWithEngineCuts != null
                    && productsWithEngineCuts.Contains(hostLabel)
                    && buckets.ContainsKey(hostLabel))
                {
                    foreach (int childLabel in childLabels)
                    {
                        if (!buckets.ContainsKey(childLabel)) continue;
                        buckets.Remove(childLabel);
                        elementInfos.Remove(childLabel);
                        mergedCount++;
                    }
                    continue;
                }

                // Fallback: host has no engine-cut geometry — merge layers
                // together so the AABB opening cutter can process the combined shape.
                //
                // xBIM produces geometry for BOTH the host wall AND the individual
                // layer parts (IfcBuildingElementPart).  These overlap — the host is
                // the full wall body, the children are individual layers that together
                // form the same shape.  To avoid doubled/overlapping geometry:
                //   - If children have geometry, discard the host body (children are
                //     the detailed decomposition and may carry individual engine cuts).
                //   - Merge the children into a single bucket for AABB opening fallback.
                var childrenInBuckets = new List<int>();
                foreach (int childLabel in childLabels)
                {
                    if (buckets.ContainsKey(childLabel))
                        childrenInBuckets.Add(childLabel);
                }

                bool hostHasGeom = buckets.ContainsKey(hostLabel);

                // When children have geometry, discard the host to avoid doubled
                // geometry (the host body overlaps the sum of its layer parts).
                if (childrenInBuckets.Count > 0 && hostHasGeom)
                {
                    buckets.Remove(hostLabel);
                    // Keep hostInfo in elementInfos — it will be copied to primary below.
                    mergedCount++;
                }

                var groupLabels = new List<int>(childrenInBuckets);
                // Only include host if it still has geometry (children were absent).
                if (buckets.ContainsKey(hostLabel))
                    groupLabels.Add(hostLabel);

                if (groupLabels.Count < 2) continue;

                int primaryLabel = SelectPrimaryWallLayerLabel(groupLabels, hostLabel, buckets, elementInfos);
                if (!buckets.TryGetValue(primaryLabel, out Bucket primaryBucket))
                    continue;

                foreach (int label in groupLabels)
                {
                    if (label == primaryLabel) continue;
                    if (!buckets.TryGetValue(label, out Bucket secondaryBucket)) continue;

                    MergeBucket(primaryBucket, secondaryBucket);
                    buckets.Remove(label);
                    elementInfos.Remove(label);
                    mergedCount++;
                }

                // Carry the host wall's identity to the merged bucket so properties
                // panel shows the wall info (not a layer part).
                if (primaryLabel != hostLabel
                    && elementInfos.TryGetValue(hostLabel, out IfcElementInfo hostInfo)
                    && IsWallLikeType(hostInfo?.Type))
                {
                    elementInfos[primaryLabel] = hostInfo;
                }
            }

            return mergedCount;
        }

        private static int SelectPrimaryWallLayerLabel(
            List<int> labels,
            int hostLabel,
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos)
        {
            if (labels == null || labels.Count == 0) return hostLabel;
            if (hostLabel > 0 && labels.Contains(hostLabel) && buckets.ContainsKey(hostLabel))
                return hostLabel;

            int bestLabel = labels[0];
            int bestPriority = GetWallTypePriority(GetIfcTypeName(bestLabel, elementInfos));
            float bestVolume = GetBucketVolume(bestLabel, buckets);

            for (int i = 1; i < labels.Count; i++)
            {
                int currentLabel = labels[i];
                int priority = GetWallTypePriority(GetIfcTypeName(currentLabel, elementInfos));
                float volume = GetBucketVolume(currentLabel, buckets);

                if (priority > bestPriority
                    || (priority == bestPriority && volume > bestVolume))
                {
                    bestLabel = currentLabel;
                    bestPriority = priority;
                    bestVolume = volume;
                }
            }

            return bestLabel;
        }

        private static string GetIfcTypeName(int label, Dictionary<int, IfcElementInfo> elementInfos)
        {
            if (elementInfos == null) return null;
            return elementInfos.TryGetValue(label, out IfcElementInfo info) ? info?.Type : null;
        }

        private static float GetBucketVolume(int label, Dictionary<int, Bucket> buckets)
        {
            if (buckets == null) return 0f;
            if (!buckets.TryGetValue(label, out Bucket bucket)) return 0f;
            if (bucket == null || bucket.Positions.Count == 0) return 0f;
            return GetVolume(GetBounds(bucket.Positions));
        }

        private static int CollapseWallLayersIntoCombinedWalls(
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos)
        {
            if (buckets == null || buckets.Count == 0) return 0;
            if (elementInfos == null || elementInfos.Count == 0) return 0;

            const float minWallHeight = 1.20f;

            var candidates = new List<WallLayerCandidate>();
            foreach (var kv in buckets)
            {
                if (!elementInfos.TryGetValue(kv.Key, out IfcElementInfo info)) continue;
                if (!ShouldApplyLayeredCut(info?.Type)) continue;
                if (kv.Value == null || kv.Value.Positions.Count == 0) continue;

                BoundingBox bounds = GetBounds(kv.Value.Positions);
                Vector3 size = bounds.Maximum - bounds.Minimum;

                // Keep likely wall-like vertical layers only.
                if (size.Y < minWallHeight) continue;

                candidates.Add(new WallLayerCandidate
                {
                    ProductLabel = kv.Key,
                    Bounds = bounds,
                    Bucket = kv.Value,
                    TypeName = info.Type,
                    Name = info.Name,
                    DominantAxis = GetDominantPlanAxis(size.X, size.Z),
                });
            }

            if (candidates.Count < 2) return 0;

            int n = candidates.Count;
            var parent = new int[n];
            for (int i = 0; i < n; i++) parent[i] = i;

            for (int i = 0; i < n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if (ShouldMergeWallLayerCandidates(candidates[i], candidates[j]))
                        Union(parent, i, j);
                }
            }

            var groups = new Dictionary<int, List<int>>();
            for (int i = 0; i < n; i++)
            {
                int root = Find(parent, i);
                if (!groups.TryGetValue(root, out List<int> list))
                {
                    list = new List<int>();
                    groups[root] = list;
                }
                list.Add(i);
            }

            int mergedCount = 0;
            foreach (List<int> group in groups.Values)
            {
                if (group.Count < 2) continue;

                int primaryIndex = SelectPrimaryWallLayerCandidate(group, candidates);
                WallLayerCandidate primary = candidates[primaryIndex];

                foreach (int idx in group)
                {
                    if (idx == primaryIndex) continue;

                    WallLayerCandidate secondary = candidates[idx];
                    MergeBucket(primary.Bucket, secondary.Bucket);
                    buckets.Remove(secondary.ProductLabel);
                    elementInfos.Remove(secondary.ProductLabel);
                    mergedCount++;
                }
            }

            return mergedCount;
        }

        private static bool ShouldMergeWallLayerCandidates(
            WallLayerCandidate a,
            WallLayerCandidate b)
        {
            const float layerGap = 0.45f;          // max separation between adjacent layers
            const float minHeightOverlapRatio = 0.60f;
            const float minLinearOverlapRatio = 0.45f;
            const float minLinearOverlapAbs = 0.60f;

            // Infer the effective horizontal run-axis for each candidate, even when
            // DominantAxis == -1 (nearly-square plan footprint).  This ensures the
            // perpendicular-wall guard fires regardless of which side is ambiguous,
            // preventing diagonal corner stubs from being merged into straight walls.
            float aXSpan = a.Bounds.Maximum.X - a.Bounds.Minimum.X;
            float aZSpan = a.Bounds.Maximum.Z - a.Bounds.Minimum.Z;
            float bXSpan = b.Bounds.Maximum.X - b.Bounds.Minimum.X;
            float bZSpan = b.Bounds.Maximum.Z - b.Bounds.Minimum.Z;

            int axisA = a.DominantAxis >= 0 ? a.DominantAxis : (aXSpan >= aZSpan ? 0 : 1);
            int axisB = b.DominantAxis >= 0 ? b.DominantAxis : (bXSpan >= bZSpan ? 0 : 1);

            // Avoid merging perpendicular walls at junctions.
            if (axisA != axisB)
                return false;

            int axis = axisA; // both axes agree at this point

            float aHeight = Math.Max(0.001f, a.Bounds.Maximum.Y - a.Bounds.Minimum.Y);
            float bHeight = Math.Max(0.001f, b.Bounds.Maximum.Y - b.Bounds.Minimum.Y);
            float overlapY = GetIntervalOverlapSize(
                a.Bounds.Minimum.Y, a.Bounds.Maximum.Y,
                b.Bounds.Minimum.Y, b.Bounds.Maximum.Y);
            if (overlapY / Math.Min(aHeight, bHeight) < minHeightOverlapRatio)
                return false;

            // If exporter split one wall into layer elements with same name,
            // allow a direct merge when they are adjacent in the cross-axis direction
            // (layers are stacked perpendicular to the wall's run direction).
            // Only check the cross-axis gap — not both axes — to avoid merging
            // same-named elements at building corners that happen to be near in both X and Z.
            if (!string.IsNullOrWhiteSpace(a.Name)
                && string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))
            {
                float crossGap = axis == 0
                    ? IntervalGap(a.Bounds.Minimum.Z, a.Bounds.Maximum.Z,
                                  b.Bounds.Minimum.Z, b.Bounds.Maximum.Z)
                    : IntervalGap(a.Bounds.Minimum.X, a.Bounds.Maximum.X,
                                  b.Bounds.Minimum.X, b.Bounds.Maximum.X);
                if (crossGap <= 0.80f)
                    return true;
            }

            float aX = Math.Max(0.001f, aXSpan);
            float aZ = Math.Max(0.001f, aZSpan);
            float bX = Math.Max(0.001f, bXSpan);
            float bZ = Math.Max(0.001f, bZSpan);

            float overlapX = GetIntervalOverlapSize(
                a.Bounds.Minimum.X, a.Bounds.Maximum.X,
                b.Bounds.Minimum.X, b.Bounds.Maximum.X);
            float overlapZ = GetIntervalOverlapSize(
                a.Bounds.Minimum.Z, a.Bounds.Maximum.Z,
                b.Bounds.Minimum.Z, b.Bounds.Maximum.Z);

            float gapX = IntervalGap(
                a.Bounds.Minimum.X, a.Bounds.Maximum.X,
                b.Bounds.Minimum.X, b.Bounds.Maximum.X);
            float gapZ = IntervalGap(
                a.Bounds.Minimum.Z, a.Bounds.Maximum.Z,
                b.Bounds.Minimum.Z, b.Bounds.Maximum.Z);

            if (axis == 0)
            {
                float minOverlap = Math.Max(minLinearOverlapAbs, Math.Min(aX, bX) * minLinearOverlapRatio);
                return overlapX >= minOverlap && gapZ <= layerGap;
            }

            {
                float minOverlap = Math.Max(minLinearOverlapAbs, Math.Min(aZ, bZ) * minLinearOverlapRatio);
                return overlapZ >= minOverlap && gapX <= layerGap;
            }
        }

        private static int GetDominantPlanAxis(float sizeX, float sizeZ)
        {
            const float axisRatio = 1.20f;
            if (sizeX >= sizeZ * axisRatio) return 0; // X-dominant
            if (sizeZ >= sizeX * axisRatio) return 1; // Z-dominant
            return -1; // ambiguous / diagonal
        }

        private static float GetIntervalOverlapSize(
            float minA,
            float maxA,
            float minB,
            float maxB)
        {
            float overlap = Math.Min(maxA, maxB) - Math.Max(minA, minB);
            return overlap > 0 ? overlap : 0f;
        }

        private static int SelectPrimaryWallLayerCandidate(
            List<int> group,
            List<WallLayerCandidate> candidates)
        {
            int best = group[0];
            float bestVolume = GetVolume(candidates[best].Bounds);
            int bestPriority = GetWallTypePriority(candidates[best].TypeName);

            for (int i = 1; i < group.Count; i++)
            {
                int current = group[i];
                int priority = GetWallTypePriority(candidates[current].TypeName);
                float volume = GetVolume(candidates[current].Bounds);

                if (priority > bestPriority
                    || (priority == bestPriority && volume > bestVolume))
                {
                    best = current;
                    bestPriority = priority;
                    bestVolume = volume;
                }
            }

            return best;
        }

        private static int GetWallTypePriority(string ifcType)
        {
            if (string.IsNullOrWhiteSpace(ifcType)) return 0;
            if (ifcType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0) return 2;
            if (string.Equals(ifcType, "IfcCovering", StringComparison.OrdinalIgnoreCase)) return 1;
            return 0;
        }

        private static float GetVolume(BoundingBox bounds)
        {
            Vector3 size = bounds.Maximum - bounds.Minimum;
            return Math.Max(0f, size.X) * Math.Max(0f, size.Y) * Math.Max(0f, size.Z);
        }

        private static void MergeBucket(Bucket target, Bucket source)
        {
            if (target == null || source == null || source.Positions.Count == 0) return;

            int offset = target.Positions.Count;
            target.Positions.AddRange(source.Positions);
            target.Normals.AddRange(source.Normals);
            foreach (int idx in source.Indices)
                target.Indices.Add(offset + idx);
        }

        private static int Find(int[] parent, int i)
        {
            while (parent[i] != i)
            {
                parent[i] = parent[parent[i]];
                i = parent[i];
            }
            return i;
        }

        private static void Union(int[] parent, int a, int b)
        {
            int ra = Find(parent, a);
            int rb = Find(parent, b);
            if (ra != rb) parent[rb] = ra;
        }

        private static List<OpeningVolume> BuildOpeningVolumes(
            List<OpeningShapeItem> openingItems,
            float toMetres)
        {
            var volumes = new List<OpeningVolume>();
            if (openingItems == null || openingItems.Count == 0) return volumes;

            const float epsilon = 0.002f;           // 2 mm tolerance

            foreach (var item in openingItems)
            {
                if (item?.ShapeData == null || item.ShapeData.Length == 0) continue;

                try
                {
                    XbimShapeTriangulation tri;
                    using (var ms = new MemoryStream(item.ShapeData))
                    using (var br = new BinaryReader(ms))
                        tri = br.ReadShapeTriangulation();

                    tri = tri.Transform(item.Transformation);

                    List<float[]> triPositions;
                    List<int> triIndices;
                    tri.ToPointsWithNormalsAndIndices(out triPositions, out triIndices);
                    if (triPositions == null || triPositions.Count == 0) continue;

                    Vector3 min = new Vector3(float.MaxValue, float.MaxValue, float.MaxValue);
                    Vector3 max = new Vector3(float.MinValue, float.MinValue, float.MinValue);

                    foreach (float[] v in triPositions)
                    {
                        var p = new Vector3(
                            v[0] * toMetres,
                            v[2] * toMetres,
                           -v[1] * toMetres);

                        if (p.X < min.X) min.X = p.X;
                        if (p.Y < min.Y) min.Y = p.Y;
                        if (p.Z < min.Z) min.Z = p.Z;
                        if (p.X > max.X) max.X = p.X;
                        if (p.Y > max.Y) max.Y = p.Y;
                        if (p.Z > max.Z) max.Z = p.Z;
                    }

                    var tol = new Vector3(epsilon, epsilon, epsilon);
                    volumes.Add(new OpeningVolume
                    {
                        Bounds = new BoundingBox(min - tol, max + tol),
                        HostProductLabels = item.HostProductLabels,
                    });
                }
                catch
                {
                    // Opening extraction is best-effort. Skip malformed opening geometry.
                }
            }

            return volumes;
        }

        private static List<BoundingBox> GetLayerCandidateBounds(
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos)
        {
            var bounds = new List<BoundingBox>();
            if (buckets == null || buckets.Count == 0) return bounds;
            if (elementInfos == null || elementInfos.Count == 0) return bounds;

            foreach (var kv in buckets)
            {
                if (!elementInfos.TryGetValue(kv.Key, out IfcElementInfo info)) continue;
                if (!ShouldApplyLayeredCut(info?.Type)) continue;
                if (kv.Value == null || kv.Value.Positions.Count == 0) continue;

                bounds.Add(GetBounds(kv.Value.Positions));
            }

            return bounds;
        }

        private static List<OpeningVolume> ExpandOpeningVolumesForLayers(
            List<OpeningVolume> openingVolumes,
            List<BoundingBox> layerBounds)
        {
            if (openingVolumes == null || openingVolumes.Count == 0)
                return new List<OpeningVolume>();
            if (layerBounds == null || layerBounds.Count == 0)
                return openingVolumes;

            const float bridgeGap = 0.30f;        // max gap between wall layers (300 mm)
            const float maxExtraDepth = 2.00f;    // clamp extension to avoid runaway cuts
            const float planarTolerance = 0.01f;  // 10 mm

            var expanded = new List<OpeningVolume>(openingVolumes.Count);

            foreach (OpeningVolume ov in openingVolumes)
            {
                BoundingBox opening = ov.Bounds;
                Vector3 min = opening.Minimum;
                Vector3 max = opening.Maximum;
                Vector3 size = max - min;
                int depthAxis = GetSmallestAxis(size);

                // A slab/floor opening has Y as its thinnest axis (vertical in scene space).
                // Expanding it along Y would stretch the cut box through the full height of
                // adjacent vertical wall layers and corrupt them.  Skip expansion and keep
                // the original bounding box as-is; the narrow box is harmless to walls.
                if (depthAxis == 1)
                {
                    expanded.Add(ov);
                    continue;
                }

                float startMin = GetAxis(min, depthAxis);
                float startMax = GetAxis(max, depthAxis);
                float depthMin = startMin;
                float depthMax = startMax;
                float maxAllowedSpan = (startMax - startMin) + maxExtraDepth;

                bool changed;
                int guard = 0;
                do
                {
                    changed = false;
                    guard++;

                    foreach (BoundingBox layer in layerBounds)
                    {
                        if (!OverlapsPlanar(opening, layer, depthAxis, planarTolerance)) continue;

                        float layerMin = GetAxis(layer.Minimum, depthAxis);
                        float layerMax = GetAxis(layer.Maximum, depthAxis);
                        if (IntervalGap(depthMin, depthMax, layerMin, layerMax) > bridgeGap)
                            continue;

                        float newMin = Math.Min(depthMin, layerMin);
                        float newMax = Math.Max(depthMax, layerMax);
                        if ((newMax - newMin) > maxAllowedSpan)
                            continue;

                        if (newMin < depthMin || newMax > depthMax)
                        {
                            depthMin = newMin;
                            depthMax = newMax;
                            changed = true;
                        }
                    }
                } while (changed && guard < 16);

                SetAxis(ref min, depthAxis, depthMin);
                SetAxis(ref max, depthAxis, depthMax);
                expanded.Add(new OpeningVolume
                {
                    Bounds = new BoundingBox(min, max),
                    HostProductLabels = ov.HostProductLabels,
                });
            }

            return expanded;
        }

        private static int GetSmallestAxis(Vector3 size)
        {
            if (size.X <= size.Y && size.X <= size.Z) return 0;
            if (size.Y <= size.X && size.Y <= size.Z) return 1;
            return 2;
        }

        private static float GetAxis(Vector3 v, int axis)
        {
            if (axis == 0) return v.X;
            if (axis == 1) return v.Y;
            return v.Z;
        }

        private static void SetAxis(ref Vector3 v, int axis, float value)
        {
            if (axis == 0)
            {
                v.X = value;
                return;
            }

            if (axis == 1)
            {
                v.Y = value;
                return;
            }

            v.Z = value;
        }

        private static float IntervalGap(float aMin, float aMax, float bMin, float bMax)
        {
            if (aMax < bMin) return bMin - aMax;
            if (bMax < aMin) return aMin - bMax;
            return 0f;
        }

        private static bool OverlapsPlanar(
            BoundingBox opening,
            BoundingBox layer,
            int depthAxis,
            float tolerance)
        {
            if (depthAxis == 0)
            {
                return OverlapsRange(opening.Minimum.Y, opening.Maximum.Y, layer.Minimum.Y, layer.Maximum.Y, tolerance)
                    && OverlapsRange(opening.Minimum.Z, opening.Maximum.Z, layer.Minimum.Z, layer.Maximum.Z, tolerance);
            }

            if (depthAxis == 1)
            {
                return OverlapsRange(opening.Minimum.X, opening.Maximum.X, layer.Minimum.X, layer.Maximum.X, tolerance)
                    && OverlapsRange(opening.Minimum.Z, opening.Maximum.Z, layer.Minimum.Z, layer.Maximum.Z, tolerance);
            }

            return OverlapsRange(opening.Minimum.X, opening.Maximum.X, layer.Minimum.X, layer.Maximum.X, tolerance)
                && OverlapsRange(opening.Minimum.Y, opening.Maximum.Y, layer.Minimum.Y, layer.Maximum.Y, tolerance);
        }

        private static bool OverlapsRange(
            float minA,
            float maxA,
            float minB,
            float maxB,
            float tolerance)
        {
            return !(maxA + tolerance < minB || minA - tolerance > maxB);
        }

        private static int ApplyOpeningVolumesToLayers(
            Dictionary<int, Bucket> buckets,
            Dictionary<int, IfcElementInfo> elementInfos,
            List<OpeningVolume> openingVolumes,
            Dictionary<int, HashSet<int>> wallLayerRelationGroups)
        {
            if (buckets == null || buckets.Count == 0) return 0;
            if (openingVolumes == null || openingVolumes.Count == 0) return 0;

            // Pre-extract all opening bounding boxes for spatial matching.
            var allOpeningBounds = new List<BoundingBox>(openingVolumes.Count);
            foreach (var ov in openingVolumes)
                allOpeningBounds.Add(ov.Bounds);

            int cutCount = 0;
            foreach (var kv in buckets)
            {
                if (!elementInfos.TryGetValue(kv.Key, out IfcElementInfo info)) continue;
                if (!ShouldApplyLayeredCut(info?.Type)) continue;

                // Use spatial proximity: apply every opening whose AABB overlaps
                // this bucket's bounds.  Label-based targeting is unreliable after
                // wall-layer merging shuffles bucket keys.  The downstream centroid-
                // inside-shrunk-opening test prevents false-positive cuts on adjacent
                // walls at T-junctions and corners.
                BoundingBox bucketBounds = GetBounds(kv.Value.Positions);
                var relevant = new List<BoundingBox>();
                foreach (BoundingBox ob in allOpeningBounds)
                {
                    if (Intersects(bucketBounds, ob))
                        relevant.Add(ob);
                }

                if (relevant.Count == 0) continue;

                if (ApplyOpeningCutsToBucket(kv.Value, relevant))
                    cutCount++;
            }

            return cutCount;
        }

        private static bool ShouldApplyLayeredCut(string ifcType)
        {
            if (string.IsNullOrWhiteSpace(ifcType)) return false;
            if (ifcType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (string.Equals(ifcType, "IfcBuildingElementPart", StringComparison.OrdinalIgnoreCase)) return true;
            return string.Equals(ifcType, "IfcCovering", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Cuts opening volumes from a wall bucket by splitting triangles along the
        /// exact AABB face planes of each opening.  Triangles that cross an opening
        /// boundary are split into sub-triangles whose edges lie exactly on the box
        /// faces, producing clean straight edges.  Sub-triangles whose centroids fall
        /// inside the (slightly shrunk) opening are discarded.
        /// </summary>
        private static bool ApplyOpeningCutsToBucket(Bucket bucket, List<BoundingBox> openingVolumes)
        {
            if (bucket == null || bucket.Indices.Count < 3 || bucket.Positions.Count == 0)
                return false;

            BoundingBox bucketBounds = GetBounds(bucket.Positions);
            var candidates = openingVolumes.Where(v => Intersects(bucketBounds, v)).ToList();
            if (candidates.Count == 0) return false;

            // Shrink opening volumes 2 mm inward for the centroid inside/outside
            // test — protects edge triangles on geometry the engine already cut.
            const float edgeTolerance = 0.002f;
            var shrunkCandidates = new List<BoundingBox>(candidates.Count);
            foreach (var box in candidates)
            {
                var tol = new Vector3(edgeTolerance, edgeTolerance, edgeTolerance);
                shrunkCandidates.Add(new BoundingBox(
                    box.Minimum + tol, box.Maximum - tol));
            }

            // Extract all triangles from the bucket.
            var triangles = new List<SubTri>();
            for (int i = 0; i + 2 < bucket.Indices.Count; i += 3)
            {
                int i0 = bucket.Indices[i];
                int i1 = bucket.Indices[i + 1];
                int i2 = bucket.Indices[i + 2];
                if (i0 < 0 || i1 < 0 || i2 < 0
                    || i0 >= bucket.Positions.Count
                    || i1 >= bucket.Positions.Count
                    || i2 >= bucket.Positions.Count)
                    continue;
                triangles.Add(new SubTri
                {
                    P0 = bucket.Positions[i0], P1 = bucket.Positions[i1], P2 = bucket.Positions[i2],
                    N0 = bucket.Normals[i0],   N1 = bucket.Normals[i1],   N2 = bucket.Normals[i2],
                });
            }

            bool modified = false;

            // For each opening, split intersecting triangles along the opening's
            // 6 AABB face planes and discard pieces whose centroids are inside.
            for (int ci = 0; ci < candidates.Count; ci++)
            {
                BoundingBox box  = candidates[ci];
                BoundingBox test = shrunkCandidates[ci];
                var next = new List<SubTri>(triangles.Count);

                foreach (SubTri tri in triangles)
                {
                    if (!TriIntersectsBox(tri, box))
                    {
                        next.Add(tri);
                        continue;
                    }

                    // Split the triangle against all 6 face planes of the AABB.
                    var pieces = SplitTriAgainstAABB(tri, box);

                    // Collect kept pieces and check if any were discarded.
                    var kept = new List<SubTri>(pieces.Count);
                    bool anyDiscarded = false;
                    foreach (SubTri piece in pieces)
                    {
                        Vector3 centroid = (piece.P0 + piece.P1 + piece.P2) / 3f;
                        if (Contains(test, centroid))
                            anyDiscarded = true;
                        else
                            kept.Add(piece);
                    }

                    if (anyDiscarded)
                    {
                        // Real cut — keep the split pieces.
                        modified = true;
                        next.AddRange(kept);
                    }
                    else
                    {
                        // No pieces discarded — split was useless, keep original.
                        next.Add(tri);
                    }
                }

                triangles = next;
            }

            if (!modified) return false;

            // Rebuild bucket from surviving triangles.
            bucket.Positions.Clear();
            bucket.Normals.Clear();
            bucket.Indices.Clear();

            foreach (SubTri tri in triangles)
            {
                int baseIdx = bucket.Positions.Count;
                bucket.Positions.Add(tri.P0);
                bucket.Positions.Add(tri.P1);
                bucket.Positions.Add(tri.P2);
                bucket.Normals.Add(tri.N0);
                bucket.Normals.Add(tri.N1);
                bucket.Normals.Add(tri.N2);
                bucket.Indices.Add(baseIdx);
                bucket.Indices.Add(baseIdx + 1);
                bucket.Indices.Add(baseIdx + 2);
            }

            return true;
        }

        // ── Triangle-AABB plane splitting ────────────────────────────────────
        //
        // Splits triangles along the exact face planes of an AABB so that the
        // resulting sub-triangle edges align perfectly with the box boundaries.
        // This produces clean straight edges at opening cuts, regardless of how
        // coarsely the original wall mesh was tessellated.

        private struct SubTri
        {
            public Vector3 P0, P1, P2;
            public Vector3 N0, N1, N2;
        }

        private static bool TriIntersectsBox(SubTri tri, BoundingBox box)
        {
            float minX = Math.Min(tri.P0.X, Math.Min(tri.P1.X, tri.P2.X));
            float minY = Math.Min(tri.P0.Y, Math.Min(tri.P1.Y, tri.P2.Y));
            float minZ = Math.Min(tri.P0.Z, Math.Min(tri.P1.Z, tri.P2.Z));
            float maxX = Math.Max(tri.P0.X, Math.Max(tri.P1.X, tri.P2.X));
            float maxY = Math.Max(tri.P0.Y, Math.Max(tri.P1.Y, tri.P2.Y));
            float maxZ = Math.Max(tri.P0.Z, Math.Max(tri.P1.Z, tri.P2.Z));
            return !(maxX < box.Minimum.X || minX > box.Maximum.X
                  || maxY < box.Minimum.Y || minY > box.Maximum.Y
                  || maxZ < box.Minimum.Z || minZ > box.Maximum.Z);
        }

        /// <summary>
        /// Splits a triangle against all 6 face planes of an AABB.
        /// Returns sub-triangles whose edges align with the box faces.
        /// </summary>
        private static List<SubTri> SplitTriAgainstAABB(SubTri tri, BoundingBox box)
        {
            var current = new List<SubTri>(8) { tri };

            // Split against each axis-aligned face plane (6 planes total).
            for (int axis = 0; axis < 3; axis++)
            {
                float minVal = (axis == 0) ? box.Minimum.X
                             : (axis == 1) ? box.Minimum.Y
                             : box.Minimum.Z;
                float maxVal = (axis == 0) ? box.Maximum.X
                             : (axis == 1) ? box.Maximum.Y
                             : box.Maximum.Z;

                current = SplitListByPlane(current, axis, minVal);
                current = SplitListByPlane(current, axis, maxVal);
            }

            return current;
        }

        private static List<SubTri> SplitListByPlane(List<SubTri> tris, int axis, float splitValue)
        {
            var result = new List<SubTri>(tris.Count + 4);
            foreach (SubTri tri in tris)
                SplitOneTriByPlane(tri, axis, splitValue, result);
            return result;
        }

        /// <summary>
        /// Splits one triangle by an axis-aligned plane at <paramref name="splitValue"/>.
        /// Produces 1 triangle (no split needed) or 3 triangles (one vertex isolated
        /// on one side → 1 triangle on that side + quad on the other side split into 2).
        /// New vertices are created by linear interpolation along the crossing edges,
        /// ensuring edges align exactly with the plane.
        /// </summary>
        private static void SplitOneTriByPlane(SubTri tri, int axis, float splitValue,
                                                List<SubTri> result)
        {
            const float eps = 0.0005f; // 0.5 mm snap tolerance

            float d0 = AxisOf(tri.P0, axis) - splitValue;
            float d1 = AxisOf(tri.P1, axis) - splitValue;
            float d2 = AxisOf(tri.P2, axis) - splitValue;

            // Snap near-zero distances to zero to avoid degenerate slivers.
            if (Math.Abs(d0) < eps) d0 = 0f;
            if (Math.Abs(d1) < eps) d1 = 0f;
            if (Math.Abs(d2) < eps) d2 = 0f;

            int s0 = Math.Sign(d0);
            int s1 = Math.Sign(d1);
            int s2 = Math.Sign(d2);

            // All vertices on one side (or on the plane) → no split.
            if (s0 >= 0 && s1 >= 0 && s2 >= 0) { result.Add(tri); return; }
            if (s0 <= 0 && s1 <= 0 && s2 <= 0) { result.Add(tri); return; }

            // Find the isolated vertex (on the opposite side from the other two).
            // Reorder: A = isolated, B and C = same side.
            Vector3 pA, pB, pC, nA, nB, nC;
            float dA, dB, dC;

            if ((s0 > 0 && s1 <= 0 && s2 <= 0) || (s0 < 0 && s1 >= 0 && s2 >= 0))
            {
                pA = tri.P0; pB = tri.P1; pC = tri.P2;
                nA = tri.N0; nB = tri.N1; nC = tri.N2;
                dA = d0; dB = d1; dC = d2;
            }
            else if ((s1 > 0 && s0 <= 0 && s2 <= 0) || (s1 < 0 && s0 >= 0 && s2 >= 0))
            {
                pA = tri.P1; pB = tri.P0; pC = tri.P2;
                nA = tri.N1; nB = tri.N0; nC = tri.N2;
                dA = d1; dB = d0; dC = d2;
            }
            else
            {
                pA = tri.P2; pB = tri.P0; pC = tri.P1;
                nA = tri.N2; nB = tri.N0; nC = tri.N1;
                dA = d2; dB = d0; dC = d1;
            }

            // Intersection of edges A→B and A→C with the plane.
            float tAB = dA / (dA - dB);
            float tAC = dA / (dA - dC);
            tAB = Math.Max(0f, Math.Min(1f, tAB));
            tAC = Math.Max(0f, Math.Min(1f, tAC));

            Vector3 pAB = LerpV(pA, pB, tAB);
            Vector3 pAC = LerpV(pA, pC, tAC);
            Vector3 nAB = SafeNormalize(LerpV(nA, nB, tAB));
            Vector3 nAC = SafeNormalize(LerpV(nA, nC, tAC));

            // Triangle on A's side: A → AB → AC
            result.Add(new SubTri { P0 = pA, P1 = pAB, P2 = pAC, N0 = nA, N1 = nAB, N2 = nAC });

            // Quad on B/C's side: AB → B → C → AC → split into 2 triangles.
            result.Add(new SubTri { P0 = pAB, P1 = pB, P2 = pC, N0 = nAB, N1 = nB, N2 = nC });
            result.Add(new SubTri { P0 = pAB, P1 = pC, P2 = pAC, N0 = nAB, N1 = nC, N2 = nAC });
        }

        private static float AxisOf(Vector3 v, int axis)
        {
            if (axis == 0) return v.X;
            if (axis == 1) return v.Y;
            return v.Z;
        }

        private static Vector3 LerpV(Vector3 a, Vector3 b, float t)
        {
            return new Vector3(
                a.X + (b.X - a.X) * t,
                a.Y + (b.Y - a.Y) * t,
                a.Z + (b.Z - a.Z) * t);
        }

        private static Vector3 SafeNormalize(Vector3 v)
        {
            float len = v.Length();
            return len > 1e-8f ? v / len : v;
        }

        private static BoundingBox GetBounds(List<Vector3> points)
        {
            Vector3 min = points[0];
            Vector3 max = points[0];

            for (int i = 1; i < points.Count; i++)
            {
                Vector3 p = points[i];
                if (p.X < min.X) min.X = p.X;
                if (p.Y < min.Y) min.Y = p.Y;
                if (p.Z < min.Z) min.Z = p.Z;
                if (p.X > max.X) max.X = p.X;
                if (p.Y > max.Y) max.Y = p.Y;
                if (p.Z > max.Z) max.Z = p.Z;
            }

            return new BoundingBox(min, max);
        }

        private static bool Contains(BoundingBox box, Vector3 point)
        {
            return point.X >= box.Minimum.X && point.X <= box.Maximum.X
                && point.Y >= box.Minimum.Y && point.Y <= box.Maximum.Y
                && point.Z >= box.Minimum.Z && point.Z <= box.Maximum.Z;
        }

        private static bool Intersects(BoundingBox a, BoundingBox b)
        {
            return !(a.Maximum.X < b.Minimum.X || a.Minimum.X > b.Maximum.X
                  || a.Maximum.Y < b.Minimum.Y || a.Minimum.Y > b.Maximum.Y
                  || a.Maximum.Z < b.Minimum.Z || a.Minimum.Z > b.Maximum.Z);
        }

        private static bool Intersects(Vector3 min, Vector3 max, BoundingBox box)
        {
            return !(max.X < box.Minimum.X || min.X > box.Maximum.X
                  || max.Y < box.Minimum.Y || min.Y > box.Maximum.Y
                  || max.Z < box.Minimum.Z || min.Z > box.Maximum.Z);
        }

        // ── Colour resolution ────────────────────────────────────────────────
        private static XbimColour ResolveColour(XbimShapeInstance instance,
                                                IModel model,
                                                XbimColourMap colourMap)
        {
            // StyleLabel > 0: direct surface style on the shape representation item.
            // StyleLabel == 0: no style assigned by xBIM engine's IIfcStyledItem scan.
            //   → try material-level style (IfcRelAssociatesMaterial → IfcMaterialDefinitionRepresentation)
            //   → fall through to type-based fallback.
            if (instance.StyleLabel > 0)
            {
                var colour = ExtractColourFromSurfaceStyle(model, instance.StyleLabel);
                if (colour != null) return colour;
            }

            // Try material-based style when no direct style exists.
            try
            {
                IPersistEntity entity = model.Instances[instance.IfcProductLabel];
                if (entity is IIfcProduct product)
                {
                    var materialColour = ResolveMaterialColour(product, model);
                    if (materialColour != null) return materialColour;

                    string typeName = product.ExpressType.Name;
                    if (colourMap.Contains(typeName)) return colourMap[typeName];
                    return GetFallbackColour(typeName);
                }
            }
            catch { /* fall through */ }

            return ColDefault;
        }

        /// <summary>
        /// Extracts colour from an IfcSurfaceStyle entity by label.
        /// Handles both IFC2x3 and IFC4 through the unified interface layer.
        /// Prefers DiffuseColour (when it's an explicit IIfcColourRgb) over SurfaceColour.
        /// </summary>
        private static XbimColour ExtractColourFromSurfaceStyle(IModel model, int styleLabel)
        {
            try
            {
                IPersistEntity styleEntity = model.Instances[styleLabel];
                if (!(styleEntity is IIfcSurfaceStyle ss)) return null;

                foreach (var item in ss.Styles)
                {
                    if (!(item is IIfcSurfaceStyleShading shading)) continue;

                    float alpha = 1f - (float)(shading.Transparency ?? 0);

                    // Prefer explicit DiffuseColour (IIfcColourRgb) when present —
                    // it is the author-intended render colour.  SurfaceColour is
                    // the base reflectance that DiffuseColour overrides.
                    if (shading is IIfcSurfaceStyleRendering rendering
                        && rendering.DiffuseColour is IIfcColourRgb diffuseRgb)
                    {
                        return new XbimColour("s",
                            (float)diffuseRgb.Red, (float)diffuseRgb.Green,
                            (float)diffuseRgb.Blue, alpha);
                    }

                    var col = shading.SurfaceColour;
                    return new XbimColour("s",
                        (float)col.Red, (float)col.Green, (float)col.Blue, alpha);
                }
            }
            catch { /* fall through */ }

            return null;
        }

        /// <summary>
        /// Resolves colour from the product's associated material when no direct
        /// surface style exists (StyleLabel == 0).  Traverses:
        ///   IIfcProduct → IIfcRelAssociatesMaterial → material
        ///   → IIfcMaterialDefinitionRepresentation → IIfcStyledRepresentation
        ///   → IIfcStyledItem → IIfcSurfaceStyle
        /// Also checks individual material layers for layered materials.
        /// </summary>
        private static XbimColour ResolveMaterialColour(IIfcProduct product, IModel model)
        {
            try
            {
                foreach (var rel in product.HasAssociations)
                {
                    if (!(rel is IIfcRelAssociatesMaterial matRel)) continue;
                    var material = matRel.RelatingMaterial;
                    if (material == null) continue;

                    // Direct material with representation
                    if (material is IIfcMaterial singleMat)
                    {
                        var c = ExtractColourFromMaterial(singleMat, model);
                        if (c != null) return c;
                    }

                    // Material layer set usage → layer set → individual layers
                    if (material is IIfcMaterialLayerSetUsage usage)
                    {
                        var c = ExtractColourFromLayerSet(usage.ForLayerSet, model);
                        if (c != null) return c;
                    }

                    // Direct material layer set
                    if (material is IIfcMaterialLayerSet layerSet)
                    {
                        var c = ExtractColourFromLayerSet(layerSet, model);
                        if (c != null) return c;
                    }

                    // Material constituent set
                    if (material is IIfcMaterialConstituentSet constituentSet)
                    {
                        foreach (var constituent in constituentSet.MaterialConstituents)
                        {
                            if (constituent?.Material == null) continue;
                            var c = ExtractColourFromMaterial(constituent.Material, model);
                            if (c != null) return c;
                        }
                    }

                    // Material profile set
                    if (material is IIfcMaterialProfileSet profileSet)
                    {
                        foreach (var profile in profileSet.MaterialProfiles)
                        {
                            if (profile?.Material == null) continue;
                            var c = ExtractColourFromMaterial(profile.Material, model);
                            if (c != null) return c;
                        }
                    }

                    if (material is IIfcMaterialProfileSetUsage profileUsage
                        && profileUsage.ForProfileSet != null)
                    {
                        foreach (var profile in profileUsage.ForProfileSet.MaterialProfiles)
                        {
                            if (profile?.Material == null) continue;
                            var c = ExtractColourFromMaterial(profile.Material, model);
                            if (c != null) return c;
                        }
                    }
                }
            }
            catch { /* non-critical */ }

            return null;
        }

        private static XbimColour ExtractColourFromLayerSet(
            IIfcMaterialLayerSet layerSet, IModel model)
        {
            if (layerSet?.MaterialLayers == null) return null;
            foreach (var layer in layerSet.MaterialLayers)
            {
                if (layer?.Material == null) continue;
                var c = ExtractColourFromMaterial(layer.Material, model);
                if (c != null) return c;
            }
            return null;
        }

        private static XbimColour ExtractColourFromMaterial(
            IIfcMaterial material, IModel model)
        {
            if (material == null) return null;
            try
            {
                foreach (var rep in material.HasRepresentation)
                {
                    if (!(rep is IIfcMaterialDefinitionRepresentation matDefRep)) continue;
                    foreach (var styledRep in matDefRep.Representations)
                    {
                        if (!(styledRep is IIfcStyledRepresentation sr)) continue;
                        foreach (var styledItem in sr.Items)
                        {
                            if (!(styledItem is IIfcStyledItem si)) continue;
                            foreach (var styleAssign in si.Styles)
                            {
                                foreach (var surfStyle in styleAssign.SurfaceStyles)
                                {
                                    if (!(surfStyle is IPersistEntity pe)) continue;
                                    var c = ExtractColourFromSurfaceStyle(
                                        model, pe.EntityLabel);
                                    if (c != null) return c;
                                }
                            }
                        }
                    }
                }
            }
            catch { /* non-critical */ }
            return null;
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

        // ── Opening shape (used as a clipping volume for layered wall cuts) ──
        private sealed class OpeningShapeItem
        {
            public byte[]       ShapeData;
            public XbimMatrix3D Transformation;
            /// <summary>
            /// Product labels of the host element(s) this opening cuts.
            /// Includes the host wall and its decomposed children (from IfcRelVoidsElement).
            /// Null/empty when the relationship could not be resolved.
            /// </summary>
            public HashSet<int> HostProductLabels;
        }

        // ── Opening volume with host association ────────────────────────────
        private sealed class OpeningVolume
        {
            public BoundingBox  Bounds;
            /// <summary>
            /// Product labels this opening is associated with.
            /// Null/empty = untargeted (apply to all nearby wall-like elements).
            /// </summary>
            public HashSet<int> HostProductLabels;
        }

        // ── Wall-layer merge candidate (temporary, loader-only) ──────────────
        private sealed class WallLayerCandidate
        {
            public int ProductLabel;
            public BoundingBox Bounds;
            public Bucket Bucket;
            public string TypeName;
            public string Name;
            public int DominantAxis;
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

            // Cache-deserialisation path — values already in rounded form.
            internal ColourKey(float r, float g, float b, float a)
            {
                _r = r; _g = g; _b = b; _a = a;
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
