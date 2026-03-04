using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xbim.Common;
using Xbim.Common.Geometry;
using Xbim.Common.XbimExtensions;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;
using Xbim.ModelGeometry.Scene;

namespace IfcCore
{
    /// <summary>
    /// Production implementation of <see cref="IIfcViewerService"/>.
    /// 
    /// xBIM lifecycle rules enforced:
    ///   • <see cref="IfcStore.Open"/> inside a <c>using</c> block (or explicit dispose)
    ///   • Read-only access — no transactions needed
    ///   • Single-threaded geometry generation (xBIM is not thread-safe for concurrent mutation)
    ///   • <see cref="CancellationToken"/> checked at key checkpoints
    ///   • <see cref="IDisposable"/> — cleans up model on dispose
    /// </summary>
    public sealed class IfcViewerService : IIfcViewerService
    {
        // ── Loaded model state ────────────────────────────────────────────────
        private readonly object _loadLock = new object();
        private Dictionary<string, IfcElementData> _globalIdIndex;
        private Dictionary<int, IfcGeometryData> _geometryByLabel;
        private Dictionary<string, int> _globalIdToLabel;
        private bool _disposed;

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

        // ── IFC types to skip ────────────────────────────────────────────────
        private static readonly HashSet<string> IgnoredTypeNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "IfcSpatialStructureElement", "IfcSpatialElement", "IfcSpace",
            "IfcSpatialZone", "IfcExternalSpatialElement",
            "IfcSite", "IfcBuilding", "IfcBuildingStorey",
            "IfcOpeningElement", "IfcOpeningStandardCase",
            "IfcFeatureElement", "IfcFeatureElementSubtraction",
            "IfcVoidingFeature", "IfcSurfaceFeature", "IfcVirtualElement",
            "IfcAnnotation", "IfcGrid", "IfcGridAxis",
        };

        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        //  IIfcViewerService.LoadAsync
        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        public Task<IfcLoadResult> LoadAsync(string filePath,
                                             CancellationToken cancellationToken = default)
        {
            return Task.Run(() => LoadCore(filePath, cancellationToken), cancellationToken);
        }

        private IfcLoadResult LoadCore(string filePath, CancellationToken token)
        {
            var sw = Stopwatch.StartNew();
            var result = new IfcLoadResult { FilePath = filePath };

            try
            {
                token.ThrowIfCancellationRequested();

                IfcCoreLogger.Info($"IfcViewerService: opening {Path.GetFileName(filePath)}");

                // ── Open model (read-only, in-memory provider) ───────────────
                using (var model = IfcStore.Open(filePath))
                {
                    token.ThrowIfCancellationRequested();

                    // Log schema + provider info
                    result.SchemaVersion = model.SchemaVersion.ToString();
                    result.StorageProvider = model.GeometryStore?.GetType().Name ?? "Unknown";
                    result.EntityCount = (int)model.Instances.Count;

                    IfcCoreLogger.Info($"  Schema: {result.SchemaVersion}");
                    IfcCoreLogger.Info($"  Provider: {result.StorageProvider}");
                    IfcCoreLogger.Info($"  Entities: {result.EntityCount}");

                    // ── Tessellate geometry ───────────────────────────────────
                    var context = new Xbim3DModelContext(model);
                    bool contextOk = false;
                    try
                    {
                        contextOk = context.CreateContext();
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"Geometry tessellation failed: {ex.Message}");
                        IfcCoreLogger.Warn($"CreateContext failed: {ex.Message}");
                    }

                    token.ThrowIfCancellationRequested();

                    // ── Extract element metadata ─────────────────────────────
                    var elementInfo = new Dictionary<int, IfcElementData>();
                    foreach (var product in model.Instances.OfType<IIfcProduct>())
                    {
                        if (ShouldSkipType(product)) continue;

                        var data = new IfcElementData
                        {
                            Name = product.Name?.ToString() ?? "(unnamed)",
                            Type = product.ExpressType.Name,
                            GlobalId = product.GlobalId.ToString(),
                            ProductLabel = product.EntityLabel,
                        };

                        // Extract property sets
                        try
                        {
                            foreach (var rel in product.IsDefinedBy)
                            {
                                if (!(rel is IIfcRelDefinesByProperties rdp)) continue;
                                if (!(rdp.RelatingPropertyDefinition is IIfcPropertySet pset)) continue;

                                var props = new Dictionary<string, string>();
                                foreach (var prop in pset.HasProperties)
                                {
                                    if (prop is IIfcPropertySingleValue sv)
                                    {
                                        string val = sv.NominalValue?.ToString() ?? "";
                                        props[prop.Name.ToString()] = val;
                                    }
                                }
                                if (props.Count > 0)
                                    data.PropertySets[pset.Name.ToString()] = props;
                            }
                        }
                        catch { /* Non-critical */ }

                        elementInfo[product.EntityLabel] = data;
                    }

                    token.ThrowIfCancellationRequested();

                    // ── Extract geometry per element ─────────────────────────
                    var geometry = new Dictionary<int, IfcGeometryData>();

                    if (contextOk)
                    {
                        float toMetres = (float)(1.0 / model.ModelFactors.OneMetre);
                        var colourMap = new XbimColourMap();
                        colourMap.SetProductTypeColourMap();

                        // Track which products have engine-cut geometry
                        var productsWithIncluded = new HashSet<int>();
                        using (var reader = model.GeometryStore.BeginRead())
                        {
                            foreach (XbimShapeInstance inst in reader.ShapeInstances)
                            {
                                if (inst.RepresentationType ==
                                    XbimGeometryRepresentationType.OpeningsAndAdditionsIncluded)
                                    productsWithIncluded.Add(inst.IfcProductLabel);
                            }

                            token.ThrowIfCancellationRequested();

                            foreach (XbimShapeInstance inst in reader.ShapeInstances)
                            {
                                token.ThrowIfCancellationRequested();

                                // Skip spatial/void elements
                                IPersistEntity entity = null;
                                try { entity = model.Instances[inst.IfcProductLabel]; }
                                catch { continue; }
                                if (entity == null) continue;
                                if (ShouldSkipType(entity)) continue;

                                // If product has "Included" geometry, skip "Excluded" duplicates
                                if (productsWithIncluded.Contains(inst.IfcProductLabel)
                                    && inst.RepresentationType ==
                                       XbimGeometryRepresentationType.OpeningsAndAdditionsExcluded)
                                    continue;

                                // Skip opening-only representations
                                if (inst.RepresentationType ==
                                    XbimGeometryRepresentationType.OpeningsAndAdditionsOnly)
                                    continue;

                                // Read shape data
                                IXbimShapeGeometryData geomData;
                                try { geomData = reader.ShapeGeometryOfInstance(inst); }
                                catch { continue; }
                                if (geomData?.ShapeData == null || geomData.ShapeData.Length == 0)
                                    continue;

                                // Decode triangulation
                                List<float[]> triPositions;
                                List<int> triIndices;
                                try
                                {
                                    XbimShapeTriangulation tri;
                                    using (var ms = new MemoryStream(geomData.ShapeData))
                                    using (var br = new BinaryReader(ms))
                                        tri = br.ReadShapeTriangulation();

                                    tri = tri.Transform(inst.Transformation);
                                    tri.ToPointsWithNormalsAndIndices(
                                        out triPositions, out triIndices);
                                }
                                catch (Exception ex)
                                {
                                    IfcCoreLogger.Warn($"Triangulate threw: {ex.Message}");
                                    continue;
                                }

                                if (triPositions == null || triPositions.Count == 0)
                                    continue;

                                // Resolve colour
                                var colour = ResolveColour(inst, model, colourMap);
                                int label = inst.IfcProductLabel;

                                // Accumulate into geometry bucket
                                if (!geometry.TryGetValue(label, out var gd))
                                {
                                    gd = new IfcGeometryData
                                    {
                                        ColourR = colour.Red,
                                        ColourG = colour.Green,
                                        ColourB = colour.Blue,
                                        ColourA = colour.Alpha,
                                        IsTransparent = colour.IsTransparent,
                                    };
                                    geometry[label] = gd;
                                }
                                else if (colour.IsTransparent && !gd.IsTransparent)
                                {
                                    // A transparent shape (e.g. glass pane) sharing the same
                                    // product as an opaque shape (e.g. window frame) — adopt
                                    // the transparent colour so the glass is visible.
                                    gd.ColourR = colour.Red;
                                    gd.ColourG = colour.Green;
                                    gd.ColourB = colour.Blue;
                                    gd.ColourA = colour.Alpha;
                                    gd.IsTransparent = true;
                                }

                                // Append positions/normals/indices
                                AppendGeometry(gd, triPositions, triIndices, toMetres);
                            }
                        }
                    }

                    // ── Populate result ───────────────────────────────────────
                    result.Elements = geometry;
                    result.ElementInfo = elementInfo;
                    result.ProductCount = geometry.Count;
                    result.Success = true;

                } // IfcStore disposed here

                // ── Update internal lookup indices ───────────────────────────
                lock (_loadLock)
                {
                    _geometryByLabel = result.Elements;

                    _globalIdIndex = new Dictionary<string, IfcElementData>(
                        StringComparer.OrdinalIgnoreCase);
                    _globalIdToLabel = new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);

                    foreach (var kv in result.ElementInfo)
                    {
                        string gid = kv.Value.GlobalId;
                        if (!string.IsNullOrEmpty(gid))
                        {
                            _globalIdIndex[gid] = kv.Value;
                            _globalIdToLabel[gid] = kv.Key;
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                result.Success = false;
                result.ErrorMessage = "Load cancelled.";
                IfcCoreLogger.Info("IfcViewerService: load cancelled.");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                IfcCoreLogger.Error("IfcViewerService: load failed.", ex);
            }
            finally
            {
                sw.Stop();
                result.Duration = sw.Elapsed;
                IfcCoreLogger.LogLoadSummary(result);
            }

            return result;
        }

        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        //  IIfcViewerService.GetElementByGlobalId
        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        public IfcElementData GetElementByGlobalId(string globalId)
        {
            if (string.IsNullOrEmpty(globalId)) return null;

            lock (_loadLock)
            {
                if (_globalIdIndex == null) return null;
                _globalIdIndex.TryGetValue(globalId, out var data);
                return data;
            }
        }

        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        //  IIfcViewerService.GetBoundingBox
        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        public IfcBoundingBox GetBoundingBox(string globalId)
        {
            if (string.IsNullOrEmpty(globalId)) return null;

            lock (_loadLock)
            {
                if (_globalIdToLabel == null || _geometryByLabel == null)
                    return null;
                if (!_globalIdToLabel.TryGetValue(globalId, out int label))
                    return null;
                if (!_geometryByLabel.TryGetValue(label, out var gd))
                    return null;
                if (gd.Positions == null || gd.Positions.Length < 3)
                    return null;

                float minX = float.MaxValue, minY = float.MaxValue, minZ = float.MaxValue;
                float maxX = float.MinValue, maxY = float.MinValue, maxZ = float.MinValue;

                for (int i = 0; i < gd.Positions.Length; i += 3)
                {
                    float x = gd.Positions[i], y = gd.Positions[i + 1], z = gd.Positions[i + 2];
                    if (x < minX) minX = x; if (x > maxX) maxX = x;
                    if (y < minY) minY = y; if (y > maxY) maxY = y;
                    if (z < minZ) minZ = z; if (z > maxZ) maxZ = z;
                }

                return new IfcBoundingBox(minX, minY, minZ, maxX, maxY, maxZ);
            }
        }

        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        //  IIfcViewerService.DisposeModel
        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        public void DisposeModel()
        {
            lock (_loadLock)
            {
                _globalIdIndex = null;
                _globalIdToLabel = null;
                _geometryByLabel = null;
            }
            IfcCoreLogger.Info("IfcViewerService: model disposed.");
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            DisposeModel();
        }

        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        //  Private helpers
        // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        private static bool ShouldSkipType(IPersistEntity entity)
        {
            if (entity == null) return true;
            var typeName = entity.ExpressType?.Name;
            if (typeName == null) return true;

            // Check the entity's type hierarchy
            if (IgnoredTypeNames.Contains(typeName)) return true;

            var superType = entity.ExpressType?.SuperType;
            while (superType != null)
            {
                if (IgnoredTypeNames.Contains(superType.Name)) return true;
                superType = superType.SuperType;
            }
            return false;
        }

        private static XbimColour ResolveColour(
            XbimShapeInstance instance, IModel model, XbimColourMap colourMap)
        {
            // Try surface style from geometry store
            if (instance.StyleLabel > 0)
            {
                try
                {
                    var entity = model.Instances[instance.StyleLabel];
                    if (entity is IIfcSurfaceStyle surfStyle)
                    {
                        foreach (var style in surfStyle.Styles)
                        {
                            if (style is IIfcSurfaceStyleRendering rendering)
                            {
                                var sc = rendering.SurfaceColour;
                                float a = 1f;
                                if (rendering.Transparency.HasValue)
                                    a = 1f - (float)rendering.Transparency.Value;
                                // Prefer explicit DiffuseColour (IIfcColourRgb) when present —
                                // it is the author-intended render colour.  SurfaceColour is
                                // the base reflectance that DiffuseColour overrides.
                                if (rendering.DiffuseColour is IIfcColourRgb diffuseRgb)
                                    return new XbimColour("Style",
                                        (float)diffuseRgb.Red, (float)diffuseRgb.Green,
                                        (float)diffuseRgb.Blue, a);
                                return new XbimColour("Style",
                                    (float)sc.Red, (float)sc.Green, (float)sc.Blue, a);
                            }
                            if (style is IIfcSurfaceStyleShading shading)
                            {
                                var sc = shading.SurfaceColour;
                                float a = 1f;
                                if (shading.Transparency.HasValue)
                                    a = 1f - (float)shading.Transparency.Value;
                                return new XbimColour("Style",
                                    (float)sc.Red, (float)sc.Green, (float)sc.Blue, a);
                            }
                        }
                    }
                }
                catch { /* Fall through to fallback */ }
            }

            // Fallback: colour by IFC type
            try
            {
                var product = model.Instances[instance.IfcProductLabel];
                return GetFallbackColour(product);
            }
            catch
            {
                return ColDefault;
            }
        }

        private static XbimColour GetFallbackColour(IPersistEntity entity)
        {
            if (entity == null) return ColDefault;
            var name = entity.ExpressType?.Name ?? "";

            if (name.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0) return ColWall;
            if (name.IndexOf("Slab", StringComparison.OrdinalIgnoreCase) >= 0) return ColSlab;
            if (name.IndexOf("Column", StringComparison.OrdinalIgnoreCase) >= 0) return ColColumn;
            if (name.IndexOf("Beam", StringComparison.OrdinalIgnoreCase) >= 0) return ColBeam;
            if (name.IndexOf("Window", StringComparison.OrdinalIgnoreCase) >= 0) return ColWindow;
            if (name.IndexOf("Door", StringComparison.OrdinalIgnoreCase) >= 0) return ColDoor;
            if (name.IndexOf("Stair", StringComparison.OrdinalIgnoreCase) >= 0) return ColStair;
            if (name.IndexOf("Roof", StringComparison.OrdinalIgnoreCase) >= 0) return ColRoof;
            return ColDefault;
        }

        /// <summary>
        /// Appends decoded triangulation data to an existing <see cref="IfcGeometryData"/>.
        /// Converts from IFC Z-up to Helix Y-up and scales to metres.
        /// </summary>
        private static void AppendGeometry(
            IfcGeometryData gd,
            List<float[]> triPositions,
            List<int> triIndices,
            float toMetres)
        {
            int existingVertexCount = (gd.Positions?.Length ?? 0) / 3;

            // Build new positions and normals
            var newPositions = new float[triPositions.Count * 3];
            var newNormals = new float[triPositions.Count * 3];
            for (int i = 0; i < triPositions.Count; i++)
            {
                float[] v = triPositions[i];
                // Remap IFC Z-up → Helix Y-up: (X,Y,Z) → (X,Z,-Y)
                newPositions[i * 3]     =  v[0] * toMetres;
                newPositions[i * 3 + 1] =  v[2] * toMetres;
                newPositions[i * 3 + 2] = -v[1] * toMetres;
                newNormals[i * 3]       =  v[3];
                newNormals[i * 3 + 1]   =  v[5];
                newNormals[i * 3 + 2]   = -v[4];
            }

            // Build new indices (offset by existing vertex count)
            var newIndices = new int[triIndices.Count];
            for (int i = 0; i < triIndices.Count; i++)
                newIndices[i] = triIndices[i] + existingVertexCount;

            // Merge with existing arrays
            gd.Positions = MergeArrays(gd.Positions, newPositions);
            gd.Normals = MergeArrays(gd.Normals, newNormals);
            gd.Indices = MergeArrays(gd.Indices, newIndices);
        }

        private static T[] MergeArrays<T>(T[] existing, T[] addition)
        {
            if (existing == null || existing.Length == 0) return addition;
            if (addition == null || addition.Length == 0) return existing;
            var merged = new T[existing.Length + addition.Length];
            Array.Copy(existing, 0, merged, 0, existing.Length);
            Array.Copy(addition, 0, merged, existing.Length, addition.Length);
            return merged;
        }
    }
}
