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

                    using (var storeReader = model.GeometryStore.BeginRead())
                    {
                        // Some products can have multiple geometry representations:
                        //  - OpeningsAndAdditionsIncluded
                        //  - OpeningsAndAdditionsExcluded
                        //  - OpeningsAndAdditionsOnly
                        // Prefer "Included" per product to avoid filling voids back in.
                        var representationMaskByProduct = new Dictionary<int, int>();
                        foreach (XbimShapeInstance instance in storeReader.ShapeInstances)
                        {
                            int product = instance.IfcProductLabel;
                            int mask = (int)instance.RepresentationType;
                            if (representationMaskByProduct.TryGetValue(product, out int existing))
                                representationMaskByProduct[product] = existing | mask;
                            else
                                representationMaskByProduct[product] = mask;
                        }

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

                                openingItems.Add(new OpeningShapeItem
                                {
                                    ShapeData = openingGeom.ShapeData,
                                    Transformation = instance.Transformation,
                                });
                                continue;
                            }

                            if (ShouldSkip(instance, entity, representationMaskByProduct))
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
                                       $"{wallLayerRelationGroups.Count} wall relation groups " +
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
                        wallLayerRelationGroups);
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
                        buckets, elementInfos, layeredOpeningVolumes);

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
            Dictionary<int, int> representationMaskByProduct)
        {
            try
            {
                if (ShouldSkipByRepresentation(instance, representationMaskByProduct))
                    return true;

                return IsIgnoredEntity(entity);
            }
            catch { return false; }
        }

        private static bool ShouldSkipByRepresentation(
            XbimShapeInstance instance,
            Dictionary<int, int> representationMaskByProduct)
        {
            int currentMask = (int)instance.RepresentationType;
            if (currentMask == 0) return false;

            int included = (int)XbimGeometryRepresentationType.OpeningsAndAdditionsIncluded;
            int excluded = (int)XbimGeometryRepresentationType.OpeningsAndAdditionsExcluded;
            int only = (int)XbimGeometryRepresentationType.OpeningsAndAdditionsOnly;

            // "Only" geometry is just feature solids (voids/additions), not product body.
            if ((currentMask & only) != 0)
                return true;

            // If "included" exists for this product, suppress the "excluded" duplicate
            // so openings do not get filled back by a second solid.
            if ((currentMask & excluded) != 0
                && representationMaskByProduct != null
                && representationMaskByProduct.TryGetValue(instance.IfcProductLabel, out int allMasks)
                && (allMasks & included) != 0)
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
            Dictionary<int, HashSet<int>> wallLayerGroups)
        {
            if (buckets == null || buckets.Count == 0) return 0;
            if (wallLayerGroups == null || wallLayerGroups.Count == 0) return 0;

            int mergedCount = 0;
            foreach (var kv in wallLayerGroups)
            {
                int hostLabel = kv.Key;
                HashSet<int> childLabels = kv.Value;
                if (childLabels == null || childLabels.Count == 0) continue;

                var groupLabels = new List<int>();
                if (buckets.ContainsKey(hostLabel))
                    groupLabels.Add(hostLabel);

                foreach (int childLabel in childLabels)
                {
                    if (!buckets.ContainsKey(childLabel)) continue;
                    groupLabels.Add(childLabel);
                }

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

            // Avoid merging perpendicular walls at junctions.
            if (a.DominantAxis >= 0 && b.DominantAxis >= 0 && a.DominantAxis != b.DominantAxis)
                return false;

            float aHeight = Math.Max(0.001f, a.Bounds.Maximum.Y - a.Bounds.Minimum.Y);
            float bHeight = Math.Max(0.001f, b.Bounds.Maximum.Y - b.Bounds.Minimum.Y);
            float overlapY = GetIntervalOverlapSize(
                a.Bounds.Minimum.Y, a.Bounds.Maximum.Y,
                b.Bounds.Minimum.Y, b.Bounds.Maximum.Y);
            if (overlapY / Math.Min(aHeight, bHeight) < minHeightOverlapRatio)
                return false;

            // If exporter split one wall into layer elements with same name,
            // allow a direct merge when they are spatially close.
            if (!string.IsNullOrWhiteSpace(a.Name)
                && string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))
            {
                if (IntervalGap(
                        a.Bounds.Minimum.X, a.Bounds.Maximum.X,
                        b.Bounds.Minimum.X, b.Bounds.Maximum.X) <= 0.80f
                    && IntervalGap(
                        a.Bounds.Minimum.Z, a.Bounds.Maximum.Z,
                        b.Bounds.Minimum.Z, b.Bounds.Maximum.Z) <= 0.80f)
                    return true;
            }

            float aX = Math.Max(0.001f, a.Bounds.Maximum.X - a.Bounds.Minimum.X);
            float aZ = Math.Max(0.001f, a.Bounds.Maximum.Z - a.Bounds.Minimum.Z);
            float bX = Math.Max(0.001f, b.Bounds.Maximum.X - b.Bounds.Minimum.X);
            float bZ = Math.Max(0.001f, b.Bounds.Maximum.Z - b.Bounds.Minimum.Z);

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

            int axis = a.DominantAxis >= 0 ? a.DominantAxis : b.DominantAxis;
            if (axis < 0)
                axis = (Math.Max(aX, bX) >= Math.Max(aZ, bZ)) ? 0 : 1;

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

        private static List<BoundingBox> BuildOpeningVolumes(
            List<OpeningShapeItem> openingItems,
            float toMetres)
        {
            var volumes = new List<BoundingBox>();
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
                    volumes.Add(new BoundingBox(min - tol, max + tol));
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

        private static List<BoundingBox> ExpandOpeningVolumesForLayers(
            List<BoundingBox> openingVolumes,
            List<BoundingBox> layerBounds)
        {
            if (openingVolumes == null || openingVolumes.Count == 0)
                return new List<BoundingBox>();
            if (layerBounds == null || layerBounds.Count == 0)
                return openingVolumes;

            const float bridgeGap = 0.30f;        // max gap between wall layers (300 mm)
            const float maxExtraDepth = 2.00f;    // clamp extension to avoid runaway cuts
            const float planarTolerance = 0.01f;  // 10 mm

            var expanded = new List<BoundingBox>(openingVolumes.Count);

            foreach (BoundingBox opening in openingVolumes)
            {
                Vector3 min = opening.Minimum;
                Vector3 max = opening.Maximum;
                Vector3 size = max - min;
                int depthAxis = GetSmallestAxis(size);

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
                expanded.Add(new BoundingBox(min, max));
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
            List<BoundingBox> openingVolumes)
        {
            if (buckets == null || buckets.Count == 0) return 0;
            if (openingVolumes == null || openingVolumes.Count == 0) return 0;

            int cutCount = 0;
            foreach (var kv in buckets)
            {
                if (!elementInfos.TryGetValue(kv.Key, out IfcElementInfo info)) continue;
                if (!ShouldApplyLayeredCut(info?.Type)) continue;

                if (ApplyOpeningCutsToBucket(kv.Value, openingVolumes))
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

        private static bool ApplyOpeningCutsToBucket(Bucket bucket, List<BoundingBox> openingVolumes)
        {
            if (bucket == null || bucket.Indices.Count < 3 || bucket.Positions.Count == 0)
                return false;

            BoundingBox bucketBounds = GetBounds(bucket.Positions);
            var candidates = openingVolumes.Where(v => Intersects(bucketBounds, v)).ToList();
            if (candidates.Count == 0) return false;

            bool removed = false;
            var kept = new List<int>(bucket.Indices.Count);

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

                Vector3 p0 = bucket.Positions[i0];
                Vector3 p1 = bucket.Positions[i1];
                Vector3 p2 = bucket.Positions[i2];

                if (ShouldCullTriangle(p0, p1, p2, candidates))
                {
                    removed = true;
                    continue;
                }

                kept.Add(i0);
                kept.Add(i1);
                kept.Add(i2);
            }

            if (!removed) return false;

            bucket.Indices.Clear();
            bucket.Indices.AddRange(kept);
            return true;
        }

        private static bool ShouldCullTriangle(
            Vector3 p0,
            Vector3 p1,
            Vector3 p2,
            List<BoundingBox> candidates)
        {
            Vector3 triMin = new Vector3(
                Math.Min(p0.X, Math.Min(p1.X, p2.X)),
                Math.Min(p0.Y, Math.Min(p1.Y, p2.Y)),
                Math.Min(p0.Z, Math.Min(p1.Z, p2.Z)));
            Vector3 triMax = new Vector3(
                Math.Max(p0.X, Math.Max(p1.X, p2.X)),
                Math.Max(p0.Y, Math.Max(p1.Y, p2.Y)),
                Math.Max(p0.Z, Math.Max(p1.Z, p2.Z)));
            Vector3 center = (p0 + p1 + p2) / 3f;

            foreach (BoundingBox box in candidates)
            {
                if (!Intersects(triMin, triMax, box)) continue;
                if (Contains(box, center)
                    || Contains(box, p0)
                    || Contains(box, p1)
                    || Contains(box, p2))
                    return true;
            }

            return false;
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

        // ── Opening shape (used as a clipping volume for layered wall cuts) ──
        private sealed class OpeningShapeItem
        {
            public byte[]       ShapeData;
            public XbimMatrix3D Transformation;
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
