using Autodesk.Revit.DB;
using HelixToolkit.Wpf.SharpDX;
using IfcViewer.Viewer;
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
    /// Geometry is merged into one mesh per colour for fast rendering; per-element
    /// identity is kept via <see cref="ElementHandle"/> vertex ranges. Incremental
    /// updates via <see cref="ExportIncremental"/> re-tessellate only dirty elements
    /// and re-merge them with the retained buckets of unchanged elements.
    ///
    /// Linked Revit models visible in the view are exported too (full exports only):
    /// their elements render, pick, and show properties like host elements, but they
    /// are static — incremental syncs carry the previous linked geometry over
    /// unchanged rather than re-tessellating it.
    ///
    /// Threading: <see cref="Export"/> / <see cref="ExportIncremental"/> must be called
    /// on the Revit API thread. Helix WPF objects are marshalled to
    /// <paramref name="uiDispatcher"/>.
    /// </summary>
    public static class RevitExporter
    {
        // ── Public API ───────────────────────────────────────────────────────

        /// <summary>
        /// Full export — tessellates every visible element in <paramref name="view"/>.
        /// </summary>
        public static RevitModel Export(Document doc, View3D view, Dispatcher uiDispatcher)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            SessionLogger.Info($"RevitExporter: exporting view '{view.Name}'");

            var context = new RevitExportContext(doc, filterIds: null);
            using (var exporter = new CustomExporter(doc, context))
            {
                exporter.IncludeGeometricObjects = false;
                exporter.ShouldStopOnError       = false;
                exporter.Export(view);
            }

            sw.Stop();
            SessionLogger.Info(
                $"RevitExporter: done in {sw.ElapsedMilliseconds} ms. " +
                $"Elements={context.ElementBuckets.Count}  Linked={context.LinkedBuckets.Count}  " +
                $"Faces={context.FaceCount}");

            return uiDispatcher.Invoke(() => BuildScene(context, view.Name));
        }

        /// <summary>
        /// Incremental export — re-tessellates only <paramref name="dirtyIds"/> and removes
        /// <paramref name="deletedIds"/> from the live <paramref name="previous"/> scene.
        /// Falls back to a full <see cref="Export"/> if <paramref name="previous"/> is null.
        /// </summary>
        public static RevitModel ExportIncremental(
            Document        doc,
            View3D          view,
            ISet<ElementId> dirtyIds,
            ISet<ElementId> deletedIds,
            RevitModel      previous,
            Dispatcher      uiDispatcher)
        {
            if (previous == null || dirtyIds == null)
                return Export(doc, view, uiDispatcher);

            var sw = System.Diagnostics.Stopwatch.StartNew();
            SessionLogger.Info(
                $"RevitExporter: incremental — dirty={dirtyIds.Count}  deleted={deletedIds?.Count ?? 0}");

            // Export only dirty elements; the context skips all others via OnElementBegin.
            var context = new RevitExportContext(doc, filterIds: dirtyIds);
            using (var exporter = new CustomExporter(doc, context))
            {
                exporter.IncludeGeometricObjects = false;
                exporter.ShouldStopOnError       = false;
                exporter.Export(view);
            }

            sw.Stop();
            SessionLogger.Info(
                $"RevitExporter: incremental done in {sw.ElapsedMilliseconds} ms. " +
                $"NewBuckets={context.ElementBuckets.Count}");

            return uiDispatcher.Invoke(() => PatchScene(context, deletedIds, previous));
        }

        // ── Full scene builder (UI thread) ────────────────────────────────────

        private static RevitModel BuildScene(RevitExportContext ctx, string viewName)
        {
            var buckets = new Dictionary<ElementId, ElementBucket>(ctx.ElementBuckets);
            var infos   = new Dictionary<ElementId, RevitElementInfo>(ctx.ElementInfos);
            return BuildMergedModel(viewName, buckets, infos, ctx.LinkedBuckets, ctx.LinkedInfos);
        }

        // ── Incremental scene patcher (UI thread) ─────────────────────────────
        //
        // With merged-by-colour rendering there is no per-element mesh to swap, so an
        // incremental sync re-merges the retained buckets (deleted removed, dirty
        // replaced). Merging is pure buffer copying — the expensive tessellation was
        // still limited to the dirty elements.

        private static RevitModel PatchScene(
            RevitExportContext ctx,
            ISet<ElementId>    deletedIds,
            RevitModel         previous)
        {
            var buckets = new Dictionary<ElementId, ElementBucket>(previous.Buckets);
            var infos   = new Dictionary<ElementId, RevitElementInfo>(
                (IDictionary<ElementId, RevitElementInfo>)previous.ElementInfos);

            if (deletedIds != null)
            {
                foreach (var id in deletedIds)
                {
                    buckets.Remove(id);
                    infos.Remove(id);
                }
            }

            foreach (var kv in ctx.ElementBuckets)
            {
                if (kv.Value.Indices.Count == 0) { buckets.Remove(kv.Key); continue; }
                buckets[kv.Key] = kv.Value;
                if (ctx.ElementInfos.TryGetValue(kv.Key, out var info))
                    infos[kv.Key] = info;
            }

            // Linked models are static: incremental exports skip them entirely, so
            // the previous export's linked geometry is carried over unchanged.
            return BuildMergedModel(previous.DisplayName, buckets, infos,
                                    previous.LinkedBuckets, previous.LinkedInfos);
        }

        // ── Shared merged-scene factory (UI thread) ───────────────────────────

        private static RevitModel BuildMergedModel(
            string displayName,
            Dictionary<ElementId, ElementBucket>    buckets,
            Dictionary<ElementId, RevitElementInfo> infos,
            Dictionary<LinkedElementKey, ElementBucket>    linkedBuckets,
            Dictionary<LinkedElementKey, RevitElementInfo> linkedInfos)
        {
            var builder = new MergedSceneBuilder();
            var handles = new Dictionary<ElementId, ElementHandle>();
            int triCount = 0;

            foreach (var kv in buckets)
            {
                var bucket = kv.Value;
                if (bucket.Indices.Count == 0) continue;

                infos.TryGetValue(kv.Key, out RevitElementInfo info);
                var handle = new ElementHandle { Info = info };
                handles[kv.Key] = handle;

                builder.AddElement(handle, bucket.Colour, bucket.Colour.Alpha < 0.99f,
                                   bucket.Positions, bucket.Normals, bucket.Indices);
                triCount += bucket.Indices.Count / 3;
            }

            // Linked-model elements render and pick like host elements, but their
            // handles stay out of the ElementId-keyed map (follow-selection and
            // incremental sync are host-only; linked ids could collide with host ids).
            int linkedCount = 0;
            foreach (var kv in linkedBuckets)
            {
                var bucket = kv.Value;
                if (bucket.Indices.Count == 0) continue;

                linkedInfos.TryGetValue(kv.Key, out RevitElementInfo info);
                var handle = new ElementHandle { Info = info };

                builder.AddElement(handle, bucket.Colour, bucket.Colour.Alpha < 0.99f,
                                   bucket.Positions, bucket.Normals, bucket.Indices);
                triCount += bucket.Indices.Count / 3;
                linkedCount++;
            }

            List<MergedMeshInfo> merged = builder.BuildGeometries();

            var sceneGroup = new GroupModel3D();
            BoundingBox bounds = new BoundingBox();
            bool hasBounds = false;
            foreach (var mi in merged)
            {
                sceneGroup.Children.Add(MergedSceneBuilder.CreateMeshNode(mi));
                bounds = hasBounds ? BoundingBox.Merge(bounds, mi.Geometry.Bound) : mi.Geometry.Bound;
                hasBounds = true;
            }

            SessionLogger.Info(
                $"RevitExporter: {triCount} triangles, {handles.Count} elements " +
                $"(+{linkedCount} linked), {merged.Count} merged meshes — scene ready.");
            return new RevitModel(displayName, sceneGroup, bounds,
                handles.Count + linkedCount, triCount, handles, infos, buckets,
                linkedBuckets, linkedInfos);
        }
    }

    // ── IExportContext implementation ─────────────────────────────────────────

    internal sealed class RevitExportContext : IExportContext
    {
        // ── Output ───────────────────────────────────────────────────────────
        public readonly Dictionary<ElementId, ElementBucket>     ElementBuckets
            = new Dictionary<ElementId, ElementBucket>();
        public readonly Dictionary<ElementId, RevitElementInfo>  ElementInfos
            = new Dictionary<ElementId, RevitElementInfo>();
        public readonly Dictionary<LinkedElementKey, ElementBucket>    LinkedBuckets
            = new Dictionary<LinkedElementKey, ElementBucket>();
        public readonly Dictionary<LinkedElementKey, RevitElementInfo> LinkedInfos
            = new Dictionary<LinkedElementKey, RevitElementInfo>();
        public int FaceCount { get; private set; }

        // ── Per-element state ────────────────────────────────────────────────
        private readonly Document        _doc;
        private readonly ISet<ElementId> _filterIds; // null = export all
        private readonly Stack<Transform> _transformStack = new Stack<Transform>();
        private Color4    _currentColor     = new Color4(0.7f, 0.7f, 0.7f, 1f);
        private bool      _skipElement;
        private ElementId _currentElementId = ElementId.InvalidElementId;

        // ── Linked-model state ───────────────────────────────────────────────
        // Element callbacks inside OnLinkBegin/OnLinkEnd carry ids from the LINKED
        // document, so they must be resolved against it — and kept in separate
        // buckets because linked ids can collide with host ids. Each link traversal
        // gets a unique visit number so the same document placed as several link
        // instances (each with its own transform) stays distinct.
        private readonly Stack<Document> _linkDocStack   = new Stack<Document>();
        private readonly Stack<int>      _linkVisitStack = new Stack<int>();
        private int              _linkVisitCounter;
        private LinkedElementKey _currentLinkedKey;
        private bool InLink => _linkDocStack.Count > 0;

        private static readonly HashSet<BuiltInCategory> SkippedCategories
            = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_Cameras,
            BuiltInCategory.OST_RenderRegions,
            BuiltInCategory.OST_SectionBox,
        };

        public RevitExportContext(Document doc, ISet<ElementId> filterIds)
        {
            _doc       = doc;
            _filterIds = filterIds;
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

        public RenderNodeAction OnViewBegin(ViewNode node) => RenderNodeAction.Proceed;
        public void OnViewEnd(ElementId elementId) { }

        // ── Element ──────────────────────────────────────────────────────────

        public RenderNodeAction OnElementBegin(ElementId elementId)
        {
            _skipElement      = false;
            _currentElementId = elementId;

            if (InLink)
                return OnLinkedElementBegin(elementId);

            // Incremental filter: skip elements not in the dirty set
            if (_filterIds != null && !_filterIds.Contains(elementId))
            {
                _skipElement = true;
                return RenderNodeAction.Skip;
            }

            Element elem = _doc.GetElement(elementId);
            if (elem == null) { _skipElement = true; return RenderNodeAction.Skip; }

            if (IsSkippedCategory(elem)) { _skipElement = true; return RenderNodeAction.Skip; }

            _currentColor = CategoryToColor(elem);

            // Extract element properties for the selection panel.
            ElementInfos[elementId] = ExtractElementInfo(elem, elementId, _doc);

            return RenderNodeAction.Proceed;
        }

        private RenderNodeAction OnLinkedElementBegin(ElementId elementId)
        {
            // Linked models are static in the viewer: exported on full syncs only,
            // carried over unchanged through incremental patches — so skip them
            // entirely when an incremental dirty-set filter is active.
            Document linkDoc = _linkDocStack.Peek();
            if (_filterIds != null || linkDoc == null)
            { _skipElement = true; return RenderNodeAction.Skip; }

            Element elem = linkDoc.GetElement(elementId);
            if (elem == null) { _skipElement = true; return RenderNodeAction.Skip; }

            if (IsSkippedCategory(elem)) { _skipElement = true; return RenderNodeAction.Skip; }

            _currentColor     = CategoryToColor(elem);
            _currentLinkedKey = new LinkedElementKey(_linkVisitStack.Peek(), elementId.Value);

            var info = ExtractElementInfo(elem, elementId, linkDoc);
            try
            {
                info.PropertySets["Linked Model"] = new Dictionary<string, string>
                { ["Source"] = linkDoc.Title };
            }
            catch { /* provenance is best-effort */ }
            LinkedInfos[_currentLinkedKey] = info;

            return RenderNodeAction.Proceed;
        }

        private static bool IsSkippedCategory(Element elem)
        {
            if (elem.Category == null) return false;
            try
            {
                var bic = (BuiltInCategory)(int)elem.Category.Id.Value;
                return SkippedCategories.Contains(bic);
            }
            catch { return false; /* non-built-in category — proceed */ }
        }

        public void OnElementEnd(ElementId elementId) { _skipElement = false; }

        // ── Instances / links ─────────────────────────────────────────────────

        public RenderNodeAction OnInstanceBegin(InstanceNode node)
        {
            _transformStack.Push(CurrentTransform.Multiply(node.GetTransform()));
            return RenderNodeAction.Proceed;
        }

        public void OnInstanceEnd(InstanceNode node)
        {
            if (_transformStack.Count > 1) _transformStack.Pop();
        }

        public RenderNodeAction OnLinkBegin(LinkNode node)
        {
            _transformStack.Push(CurrentTransform.Multiply(node.GetTransform()));
            Document linkDoc = null;
            try { linkDoc = node.GetDocument(); } catch { /* unloaded link */ }
            _linkDocStack.Push(linkDoc);
            _linkVisitStack.Push(++_linkVisitCounter);
            return RenderNodeAction.Proceed;
        }

        public void OnLinkEnd(LinkNode node)
        {
            if (_transformStack.Count > 1) _transformStack.Pop();
            if (_linkDocStack.Count > 0)
            {
                _linkDocStack.Pop();
                _linkVisitStack.Pop();
            }
        }

        // ── Face / material ───────────────────────────────────────────────────

        public RenderNodeAction OnFaceBegin(FaceNode node) => RenderNodeAction.Proceed;
        public void OnFaceEnd(FaceNode node) { }

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

            var pts    = polymesh.GetPoints();
            var norms  = polymesh.GetNormals();
            var facets = polymesh.GetFacets();

            Transform xform = CurrentTransform;

            ElementBucket bucket;
            if (InLink)
            {
                if (!LinkedBuckets.TryGetValue(_currentLinkedKey, out bucket))
                {
                    bucket = new ElementBucket { Colour = _currentColor };
                    LinkedBuckets[_currentLinkedKey] = bucket;
                }
            }
            else if (!ElementBuckets.TryGetValue(_currentElementId, out bucket))
            {
                bucket = new ElementBucket { Colour = _currentColor };
                ElementBuckets[_currentElementId] = bucket;
            }

            // Colour is updated to the most recently seen material on this element.
            bucket.Colour = _currentColor;

            int  baseIdx    = bucket.Positions.Count;
            bool hasNormals = norms != null && norms.Count == pts.Count;

            // Revit uses feet internally; convert to metres (* 0.3048) then
            // remap Revit Z-up → Helix Y-up: (X, Y, Z) → (X, Z, -Y)
            const float FT_TO_M = 0.3048f;

            for (int i = 0; i < pts.Count; i++)
            {
                XYZ tp = xform.OfPoint(pts[i]);
                bucket.Positions.Add(new Vector3(
                    (float)tp.X * FT_TO_M,
                    (float)tp.Z * FT_TO_M,
                   -(float)tp.Y * FT_TO_M));

                if (hasNormals)
                {
                    XYZ tn = xform.OfVector(norms[i]).Normalize();
                    bucket.Normals.Add(new Vector3((float)tn.X, (float)tn.Z, -(float)tn.Y));
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

        // ── Parameter-group label lookup ──────────────────────────────────────
        // Maps the ForgeTypeId key (after ":", version stripped) to the exact
        // Revit UI label as shown in the Properties panel (English).
        private static readonly Dictionary<string, string> _groupLabelMap
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["constraints"]              = "Constraints",
            ["construction"]             = "Construction",
            ["coupler"]                  = "Coupler",
            ["data"]                     = "Data",
            ["dimensions"]               = "Dimensions",
            ["division"]                 = "Division Geometry",
            ["electrical"]               = "Electrical",
            ["electrical-analysis"]      = "Electrical - Analysis",
            ["electrical-circuiting"]    = "Electrical - Circuiting",
            ["electrical-lighting"]      = "Electrical - Lighting",
            ["electrical-loads"]         = "Electrical - Loads",
            ["energy-analysis"]          = "Energy Analysis",
            ["fire-protection"]          = "Fire Protection",
            ["general"]                  = "General",
            ["geometry"]                 = "Geometry",
            ["ifc"]                      = "IFC Parameters",
            ["ifc-export-element"]       = "IFC Parameters",
            ["ifc-properties"]           = "IFC Parameters",
            ["identity-data"]            = "Identity Data",
            ["infrastructure"]           = "Infrastructure",
            ["layers"]                   = "Layers",
            ["materials"]                = "Materials and Finishes",
            ["mechanical"]               = "Mechanical",
            ["mechanical-airflow"]       = "Mechanical - Airflow",
            ["mechanical-loads"]         = "Mechanical - Loads",
            ["phasing"]                  = "Phasing",
            ["plumbing"]                 = "Plumbing",
            ["primary-end"]              = "Structural Analysis - End 1",
            ["secondary-end"]            = "Structural Analysis - End 2",
            ["secondary"]                = "Structural Analysis - End 2",
            ["structural"]               = "Structural",
            ["structural-analysis"]      = "Structural Analysis",
            ["structural-section"]       = "Structural - Section Geometry",
            ["tags"]                     = "Tags",
            ["text"]                     = "Text",
            ["title"]                    = "Title",
            ["visibility"]               = "Visibility",
        };

        /// <summary>
        /// Strips the ForgeTypeId prefix and version suffix, returning just the key.
        /// e.g. "autodesk.parameter.group:identity-data-1.0.0" → "identity-data"
        /// </summary>
        private static string ForgeGroupKey(ForgeTypeId typeId)
        {
            if (typeId == null) return "";
            var s = typeId.TypeId ?? "";

            // Strip namespace prefix up to and including ":"
            var colon = s.LastIndexOf(':');
            if (colon >= 0) s = s.Substring(colon + 1);

            // Strip trailing version like "-1.0.0" or "-2" by walking backwards
            // over digits and dots until we hit '-'
            int i = s.Length - 1;
            while (i >= 0 && (char.IsDigit(s[i]) || s[i] == '.')) i--;
            if (i >= 0 && s[i] == '-') s = s.Substring(0, i);

            return s.ToLowerInvariant(); // e.g. "identity-data"
        }

        private static RevitElementInfo ExtractElementInfo(Element elem, ElementId elementId, Document doc)
        {
            var info = new RevitElementInfo
            {
                Name      = elem.Name ?? "(unnamed)",
                Category  = elem.Category?.Name ?? "",
                ElementId = elementId.ToString(),
            };

            // Family / type names
            try
            {
                if (elem is FamilyInstance fi)
                    info.FamilyName = fi.Symbol?.FamilyName ?? "";

                var typeId = elem.GetTypeId();
                if (typeId != null && typeId != ElementId.InvalidElementId)
                    info.TypeName = doc.GetElement(typeId)?.Name ?? "";
            }
            catch { /* non-critical */ }

            // Parameters — group by their parameter group label
            try
            {
                foreach (Parameter param in elem.Parameters)
                {
                    if (param?.Definition == null) continue;

                    // Resolve the formatted value
                    string value = null;
                    try
                    {
                        value = param.AsValueString();
                        if (value == null)
                        {
                            switch (param.StorageType)
                            {
                                case StorageType.String:
                                    value = param.AsString();
                                    break;
                                case StorageType.Integer:
                                    value = param.AsInteger().ToString();
                                    break;
                                case StorageType.Double:
                                    value = param.AsDouble().ToString("G6");
                                    break;
                                case StorageType.ElementId:
                                    var refEl = doc.GetElement(param.AsElementId());
                                    value = refEl?.Name ?? param.AsElementId()?.ToString();
                                    break;
                            }
                        }
                    }
                    catch { continue; }

                    if (string.IsNullOrWhiteSpace(value)) continue;

                    // Resolve the group label via the ForgeTypeId lookup map.
                    string groupLabel = "Other";
                    try
                    {
                        var key = ForgeGroupKey(param.Definition.GetGroupTypeId());
                        if (!_groupLabelMap.TryGetValue(key, out groupLabel) || string.IsNullOrWhiteSpace(groupLabel))
                            groupLabel = "Other";
                    }
                    catch { }

                    if (!info.PropertySets.TryGetValue(groupLabel, out var group))
                    {
                        group = new Dictionary<string, string>();
                        info.PropertySets[groupLabel] = group;
                    }

                    // Last-write wins when duplicate parameter names exist in the same group.
                    group[param.Definition.Name] = value;
                }
            }
            catch { /* ignore entire parameter block on unexpected error */ }

            return info;
        }

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

    internal sealed class ElementBucket
    {
        public Color4 Colour;
        public readonly List<Vector3> Positions = new List<Vector3>();
        public readonly List<Vector3> Normals   = new List<Vector3>();
        public readonly List<int>     Indices   = new List<int>();
    }

    /// <summary>
    /// Identity of one element inside a linked Revit model: the link-traversal
    /// number (unique per link instance per export) plus the element's id in the
    /// linked document. Linked ids can collide with host ids, so linked elements
    /// are never keyed by bare <see cref="ElementId"/>.
    /// </summary>
    internal readonly struct LinkedElementKey : IEquatable<LinkedElementKey>
    {
        public readonly int  LinkVisit;
        public readonly long ElementId;

        public LinkedElementKey(int linkVisit, long elementId)
        {
            LinkVisit = linkVisit;
            ElementId = elementId;
        }

        public bool Equals(LinkedElementKey other)
            => LinkVisit == other.LinkVisit && ElementId == other.ElementId;
        public override bool Equals(object obj)
            => obj is LinkedElementKey other && Equals(other);
        public override int GetHashCode()
            => (LinkVisit * 397) ^ ElementId.GetHashCode();
    }
}
