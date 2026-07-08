using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;
using Media3D = System.Windows.Media.Media3D;
using WpfPoint = System.Windows.Point;

namespace IfcViewer.Viewer
{
    /// <summary>Cursor shape requested by the section-plane workflow.</summary>
    public enum SectionPlaneCursor { Default, Crosshair, Move }

    /// <summary>
    /// BIMcollab-Zoom-style interactive section planes on top of
    /// <see cref="SectionPlaneManager"/> (which owns the actual GPU clipping).
    /// Supports up to <see cref="SectionPlaneManager.MaxPlanes"/> planes.
    ///
    /// Workflow:
    ///   BeginPick (Add button) — a small rectangle glued to the face under the
    ///        cursor previews the cut orientation; left-click places a new plane.
    ///   Idle with planes — each plane shows a rectangle sized to the model's
    ///        cross-section; hovering highlights it, clicking grabs it, a right-
    ///        click asks the host to open a Flip/Delete menu for it.
    ///   Moving — the grabbed plane follows the mouse along its normal with live
    ///        clipping. Click drops it, Esc cancels, right-click flips it.
    ///
    /// The window feeds input in viewport-relative DIP coordinates:
    ///   OnMouseMove (from the swap-chain bridge), OnClick / OnRightClick (from
    ///   the WPF click handlers; element selection must be skipped when they
    ///   return true) and OnEscape.
    /// </summary>
    public sealed class SectionPlaneController
    {
        public enum SectionState { Idle, Picking, Moving }

        // Cut plane sits this far inside the picked face so the clicked surface
        // is sliced off (instant feedback) and the surviving geometry stays clear
        // of the rectangle visual (no z-fight).
        private const float PlaneEpsilon = 0.05f;
        // Minimum interval between hover hit-tests (mouse-move arrives far faster
        // than the scene needs re-picking).
        private const int HoverThrottleMs = 15;
        // A plane can be dragged this far (m) past the model bounds.
        private const float TravelPadding = 1.0f;

        // ── One placed plane ──────────────────────────────────────────────────
        private sealed class PlacedPlane
        {
            public Vector3 Normal;   // points INTO the removed half-space
            public Vector3 Point;    // point on the cut plane

            // In-plane basis and rectangle extents, measured about the scene
            // bbox centre. U/V extents never change while sliding along N.
            public Vector3 AxisU, AxisV;
            public Vector3 BbCentre;
            public float MinU, MaxU, MinV, MaxV;
            public float MinN, MaxN;   // travel range along the normal

            public MeshGeometryModel3D Fill;
            public LineGeometryModel3D Border;
        }

        // ── Collaborators ─────────────────────────────────────────────────────
        private readonly Viewport3DX                _viewport;
        private readonly GroupModel3D               _sceneRoot;
        private readonly SectionPlaneManager        _mgr;
        private readonly Func<BoundingBox?>         _boundsProvider;
        private readonly Action<string>             _status;
        private readonly Action<SectionPlaneCursor> _setCursor;
        private readonly Action                     _pickEnded;     // picking finished/cancelled — host syncs its Add toggle
        private readonly Action<object>             _showPlaneMenu; // host opens Flip/Delete menu for the plane token

        // ── Visuals ───────────────────────────────────────────────────────────
        private readonly GroupModel3D _visualRoot = new GroupModel3D();
        private MeshGeometryModel3D _previewFill;   // small hover rect (Picking)
        private LineGeometryModel3D _previewBorder;
        private PhongMaterial _rectMatIdle, _rectMatHover, _previewMat;
        private static readonly System.Windows.Media.Color BorderIdleColor
            = System.Windows.Media.Color.FromRgb(0x4D, 0xA6, 0xFF);
        private static readonly System.Windows.Media.Color BorderHoverColor
            = System.Windows.Media.Color.FromRgb(0x8A, 0xD4, 0xFF);

        // ── State ─────────────────────────────────────────────────────────────
        private readonly List<PlacedPlane> _planes = new List<PlacedPlane>(SectionPlaneManager.MaxPlanes);
        private SectionState _state = SectionState.Idle;
        private bool _visualsVisible = true;
        private PlacedPlane _hoverPlane;

        // ── Move interaction ──────────────────────────────────────────────────
        private PlacedPlane _movePlane;
        private Vector3 _grabBase;     // plane point when the rect was grabbed
        private Vector3 _preGrabPoint; // for Esc-cancel
        private float   _grabT;        // axis parameter under the mouse at grab time
        private int _lastHoverTick;

        public SectionPlaneController(
            Viewport3DX viewport,
            GroupModel3D sceneRoot,
            SectionPlaneManager mgr,
            Func<BoundingBox?> boundsProvider,
            Action<string> status,
            Action<SectionPlaneCursor> setCursor,
            Action pickEnded,
            Action<object> showPlaneMenu)
        {
            _viewport       = viewport;
            _sceneRoot      = sceneRoot;
            _mgr            = mgr;
            _boundsProvider = boundsProvider;
            _status         = status        ?? (_ => { });
            _setCursor      = setCursor     ?? (_ => { });
            _pickEnded      = pickEnded     ?? (() => { });
            _showPlaneMenu  = showPlaneMenu ?? (_ => { });

            BuildSharedVisuals();
            _sceneRoot.Children.Add(_visualRoot);
        }

        // ── Public surface ────────────────────────────────────────────────────

        public SectionState State => _state;

        public int PlaneCount => _planes.Count;

        /// <summary>True while the tool wants mouse-move hover updates.</summary>
        public bool IsInteractive
            => _state != SectionState.Idle || (_planes.Count > 0 && _visualsVisible);

        /// <summary>Overlay group holding all section visuals — callers (wireframe
        /// extraction etc.) skip this subtree when collecting scene meshes.</summary>
        public GroupModel3D VisualRoot => _visualRoot;

        /// <summary>
        /// Arm face picking for a new plane. Returns false when the plane limit
        /// is already reached (the host should pop its Add toggle back out).
        /// </summary>
        public bool BeginPick()
        {
            if (_state == SectionState.Picking) return true;
            if (_planes.Count >= SectionPlaneManager.MaxPlanes)
            {
                _status($"Maximum of {SectionPlaneManager.MaxPlanes} section planes reached — delete one first.");
                return false;
            }
            if (_state == SectionState.Moving) EndMove();

            _state = SectionState.Picking;
            _setCursor(SectionPlaneCursor.Crosshair);
            _status($"Add section plane ({_planes.Count + 1}/{SectionPlaneManager.MaxPlanes}): hover a face, click to place. Esc to cancel.");
            SessionLogger.Info("Section pick armed.");
            return true;
        }

        /// <summary>Leave picking mode without placing (Add toggle popped out).</summary>
        public void CancelPick()
        {
            if (_state != SectionState.Picking) return;
            _state = SectionState.Idle;
            HidePreview();
            _setCursor(SectionPlaneCursor.Default);
            _status(_planes.Count > 0
                ? "Section planes active — click a rectangle to move it, right-click for options."
                : "Section plane cancelled.");
        }

        /// <summary>Show or hide all plane rectangles; the clipping stays active.</summary>
        public void SetVisualsVisible(bool visible)
        {
            if (_visualsVisible == visible) return;
            _visualsVisible = visible;

            // Hiding mid-drag: drop the plane where it is.
            if (!visible && _state == SectionState.Moving) EndMove();
            if (!visible) SetHovered(null);

            var vis = visible ? System.Windows.Visibility.Visible
                              : System.Windows.Visibility.Collapsed;
            foreach (var p in _planes)
            {
                p.Fill.Visibility   = vis;
                p.Border.Visibility = vis;
            }

            if (_planes.Count > 0)
                _status(visible ? "Section planes visible."
                                : "Section planes hidden — the cut stays active.");
        }

        /// <summary>Flip which half-space a plane removes. Token comes from the
        /// right-click menu request.</summary>
        public void FlipPlane(object planeToken)
        {
            var plane = planeToken as PlacedPlane;
            if (plane == null || !_planes.Contains(plane)) return;

            plane.Normal = -plane.Normal;

            // Mid-move: the axis parameter is measured along the normal, so the
            // grab reference must negate with it or the plane would jump.
            if (_state == SectionState.Moving && ReferenceEquals(plane, _movePlane))
                _grabT = -_grabT;

            ComputeRectExtents(plane); // rebuild basis + travel range for the new normal
            PushPlanesToManager();
            UpdateRectTransform(plane);
            _status("Section plane flipped — the other side is now cut away.");
            SessionLogger.Info($"Section plane flipped — normal {plane.Normal}.");
        }

        /// <summary>Delete one plane. Token comes from the right-click menu request.</summary>
        public void DeletePlane(object planeToken)
        {
            var plane = planeToken as PlacedPlane;
            if (plane == null || !_planes.Contains(plane)) return;

            if (_state == SectionState.Moving && ReferenceEquals(plane, _movePlane))
            {
                _movePlane = null;
                _state = SectionState.Idle;
                _setCursor(SectionPlaneCursor.Default);
            }
            if (ReferenceEquals(plane, _hoverPlane)) SetHovered(null);

            _visualRoot.Children.Remove(plane.Fill);
            _visualRoot.Children.Remove(plane.Border);
            _planes.Remove(plane);
            PushPlanesToManager();

            _status(_planes.Count > 0
                ? $"Section plane deleted — {_planes.Count} left."
                : "Section plane deleted — no planes active.");
            SessionLogger.Info($"Section plane deleted — {_planes.Count} remain.");
        }

        /// <summary>Delete every section plane; the model is fully restored.</summary>
        public void DeleteAll()
        {
            if (_state == SectionState.Moving)
            {
                _movePlane = null;
                _state = SectionState.Idle;
                _setCursor(SectionPlaneCursor.Default);
            }
            SetHovered(null);

            foreach (var p in _planes)
            {
                _visualRoot.Children.Remove(p.Fill);
                _visualRoot.Children.Remove(p.Border);
            }
            _planes.Clear();
            PushPlanesToManager();

            _status("All section planes deleted.");
            SessionLogger.Info("All section planes deleted.");
        }

        /// <summary>Hover update. <paramref name="pos"/> is viewport-relative DIPs.</summary>
        public void OnMouseMove(WpfPoint pos)
        {
            int now = Environment.TickCount;
            if (now - _lastHoverTick < HoverThrottleMs) return;
            _lastHoverTick = now;

            switch (_state)
            {
                case SectionState.Picking:
                    UpdatePickPreview(pos);
                    break;
                case SectionState.Moving:
                    UpdateMove(pos);
                    break;
                case SectionState.Idle:
                    if (_planes.Count > 0 && _visualsVisible) UpdateHover(pos);
                    break;
            }
        }

        /// <summary>
        /// Viewport left-click. Returns true when the click was consumed by the
        /// section workflow (the caller must then skip element selection).
        /// </summary>
        public bool OnClick(WpfPoint pos)
        {
            switch (_state)
            {
                case SectionState.Picking:
                    PlaceFromPick(pos);
                    return true; // picking consumes every click, hit or miss

                case SectionState.Moving:
                    EndMove();
                    return true;

                default:
                    if (_planes.Count == 0 || !_visualsVisible) return false;
                    FindFirstHit(pos, out PlacedPlane hitPlane);
                    if (hitPlane == null) return false; // let selection handle it
                    BeginMove(hitPlane, pos);
                    return true;
            }
        }

        /// <summary>
        /// Viewport right-click (no drag). On a plane rectangle it asks the host
        /// to open the Flip/Delete menu; while moving it flips the grabbed plane.
        /// Returns true when consumed.
        /// </summary>
        public bool OnRightClick(WpfPoint pos)
        {
            switch (_state)
            {
                case SectionState.Moving:
                    FlipPlane(_movePlane);
                    return true;

                case SectionState.Idle:
                    if (_planes.Count == 0 || !_visualsVisible) return false;
                    FindFirstHit(pos, out PlacedPlane hitPlane);
                    if (hitPlane == null) return false;
                    _showPlaneMenu(hitPlane);
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>Esc: cancel a move, or leave picking mode.</summary>
        public bool OnEscape()
        {
            switch (_state)
            {
                case SectionState.Picking:
                    CancelPick();
                    _pickEnded();
                    return true;

                case SectionState.Moving:
                    _movePlane.Point = _preGrabPoint;
                    PushPlanesToManager();
                    UpdateRectTransform(_movePlane);
                    _movePlane = null;
                    _state = SectionState.Idle;
                    SetHovered(null);
                    _setCursor(SectionPlaneCursor.Default);
                    _status("Move cancelled — section plane restored.");
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// True when a hit lies in a half-space removed by an active plane — used
        /// by element selection so clipped-away (invisible) geometry cannot be
        /// picked through the cut.
        /// </summary>
        public bool IsHitClipped(HitTestResult hit)
        {
            return IsWorldPointClipped(hit.PointHit);
        }

        /// <summary>Remove all visuals from the scene (window close).</summary>
        public void Dispose()
        {
            DeleteAll();
            _sceneRoot.Children.Remove(_visualRoot);
            _visualRoot.Children.Clear();
        }

        // ── Picking ───────────────────────────────────────────────────────────

        private void UpdatePickPreview(WpfPoint pos)
        {
            var hit = FindFirstHit(pos, out PlacedPlane hitPlane);
            if (hit == null || hitPlane != null) { HidePreview(); return; }

            var rawN = hit.NormalAtHit;
            if (rawN.LengthSquared() < 1e-6f) { HidePreview(); return; }
            rawN.Normalize();
            var hitPt = hit.PointHit;

            // Size the preview with distance so it stays roughly constant on screen.
            var cam  = _viewport.Camera.Position;
            float dist = (hitPt - new Vector3((float)cam.X, (float)cam.Y, (float)cam.Z)).Length();
            float size = Clamp(dist * 0.12f, 0.2f, 40f);

            BuildBasis(rawN, out var u, out var v);
            // Lift the preview slightly off the face so it doesn't z-fight it.
            var centre = hitPt + rawN * (0.005f + size * 0.01f);

            var xform = MakeTransform(u * size, v * size, rawN, centre);
            _previewFill.Transform   = xform;
            _previewBorder.Transform = xform;
            _previewFill.Visibility   = System.Windows.Visibility.Visible;
            _previewBorder.Visibility = System.Windows.Visibility.Visible;
        }

        private void PlaceFromPick(WpfPoint pos)
        {
            var hit = FindFirstHit(pos, out PlacedPlane hitPlane);
            if (hit == null || hitPlane != null) return; // empty space / existing rect — stay picking

            var rawN = hit.NormalAtHit;
            if (rawN.LengthSquared() < 1e-6f) return;
            rawN.Normalize();

            // BIMcollab-style: cut away the side the face points to (the side
            // the user clicked from), so the model behind the plane survives.
            // Nudge the plane just INSIDE the face: the clicked surface itself
            // is sliced off (instant visual feedback) and the surviving
            // geometry sits clear of the rectangle visual (no z-fight).
            var plane = new PlacedPlane
            {
                Normal = rawN,
                Point  = hit.PointHit - rawN * PlaneEpsilon,
            };
            CreatePlaneVisuals(plane);
            ComputeRectExtents(plane);
            UpdateRectTransform(plane);
            _planes.Add(plane);
            PushPlanesToManager();

            HidePreview();
            _state = SectionState.Idle;
            _setCursor(SectionPlaneCursor.Default);
            _status($"Section plane added ({_planes.Count}/{SectionPlaneManager.MaxPlanes}) — click a rectangle to move it, right-click for options.");
            SessionLogger.Info($"Section plane added ({_planes.Count}) — normal {plane.Normal}, point {plane.Point}.");
            _pickEnded();
        }

        // ── Idle: rect hover ──────────────────────────────────────────────────

        private void UpdateHover(WpfPoint pos)
        {
            FindFirstHit(pos, out PlacedPlane hitPlane);
            SetHovered(hitPlane);
            _setCursor(hitPlane != null ? SectionPlaneCursor.Move : SectionPlaneCursor.Default);
        }

        private void SetHovered(PlacedPlane plane)
        {
            if (ReferenceEquals(_hoverPlane, plane)) return;

            if (_hoverPlane != null)
            {
                _hoverPlane.Fill.Material    = _rectMatIdle;
                _hoverPlane.Border.Color     = BorderIdleColor;
                _hoverPlane.Border.Thickness = 1.8;
            }
            _hoverPlane = plane;
            if (_hoverPlane != null)
            {
                _hoverPlane.Fill.Material    = _rectMatHover;
                _hoverPlane.Border.Color     = BorderHoverColor;
                _hoverPlane.Border.Thickness = 2.6;
            }
        }

        // ── Moving ────────────────────────────────────────────────────────────

        private void BeginMove(PlacedPlane plane, WpfPoint pos)
        {
            _movePlane    = plane;
            _preGrabPoint = plane.Point;
            _grabBase     = plane.Point;
            _grabT        = TryGetAxisT(pos, _grabBase, plane.Normal, out float t) ? t : 0f;
            _state        = SectionState.Moving;
            SetHovered(plane);
            _setCursor(SectionPlaneCursor.Move);
            _status("Moving section plane — click to finish, Esc to cancel, right-click to flip.");
        }

        private void UpdateMove(WpfPoint pos)
        {
            var plane = _movePlane;
            if (plane == null) return;
            if (!TryGetAxisT(pos, _grabBase, plane.Normal, out float t)) return;

            var p = _grabBase + (t - _grabT) * plane.Normal;

            // Clamp travel so the plane cannot be dragged past the model.
            float dn = Vector3.Dot(p - plane.BbCentre, plane.Normal);
            float clamped = Clamp(dn, plane.MinN - TravelPadding, plane.MaxN + TravelPadding);
            if (clamped != dn) p += (clamped - dn) * plane.Normal;

            plane.Point = p;
            PushPlanesToManager();
            UpdateRectTransform(plane);
        }

        private void EndMove()
        {
            _movePlane = null;
            _state = SectionState.Idle;
            _setCursor(SectionPlaneCursor.Default);
            _status("Section plane placed — click a rectangle to move it, right-click for options.");
        }

        /// <summary>
        /// Axis parameter t (plane point = origin + t * axis) whose axis point is
        /// closest to the mouse pick-ray. Standard closest-point-of-two-lines.
        /// </summary>
        private bool TryGetAxisT(WpfPoint pos, Vector3 axisOrigin, Vector3 axis, out float t)
        {
            t = 0f;
            Ray ray;
            try { ray = _viewport.UnProject(pos); }
            catch { return false; }

            var rd = ray.Direction;
            if (rd.LengthSquared() < 1e-9f) return false;
            rd.Normalize();

            float b = Vector3.Dot(rd, axis);
            float denom = 1f - b * b;
            if (Math.Abs(denom) < 1e-6f) return false; // looking straight down the axis

            var w  = ray.Position - axisOrigin;
            float d = Vector3.Dot(rd, w);
            float e = Vector3.Dot(axis, w);
            t = (e - b * d) / denom;
            return true;
        }

        // ── Hit-testing ───────────────────────────────────────────────────────

        /// <summary>
        /// Nearest hit that is either a plane rectangle fill or real scene
        /// geometry. Skips other section visuals, line overlays, and hits lying
        /// in a clipped-away half-space.
        /// </summary>
        private HitTestResult FindFirstHit(WpfPoint pos, out PlacedPlane hitPlane)
        {
            hitPlane = null;
            IList<HitTestResult> hits;
            try { hits = _viewport.FindHits(pos); }
            catch { return null; }
            if (hits == null || hits.Count == 0) return null;

            HitTestResult best = null;
            double bestDist = double.MaxValue;
            PlacedPlane bestPlane = null;

            foreach (var h in hits)
            {
                var model = h.ModelHit;
                PlacedPlane plane = PlaneFromFill(model);
                if (plane == null)
                {
                    if (!(model is MeshGeometryModel3D)) continue; // lines, etc.
                    if (IsOwnVisual(model)) continue;
                    if (IsWorldPointClipped(h.PointHit)) continue;
                }
                if (h.Distance < bestDist)
                {
                    best = h;
                    bestDist = h.Distance;
                    bestPlane = plane;
                }
            }

            hitPlane = bestPlane;
            return best;
        }

        private PlacedPlane PlaneFromFill(object model)
        {
            foreach (var p in _planes)
                if (ReferenceEquals(model, p.Fill)) return p;
            return null;
        }

        private bool IsOwnVisual(object model)
        {
            if (ReferenceEquals(model, _previewFill) || ReferenceEquals(model, _previewBorder))
                return true;
            foreach (var p in _planes)
                if (ReferenceEquals(model, p.Fill) || ReferenceEquals(model, p.Border)) return true;
            return false;
        }

        private bool IsWorldPointClipped(Vector3 wp)
        {
            foreach (var p in _planes)
            {
                if (Vector3.Dot(wp - p.Point, p.Normal) > 1e-3f)
                    return true;
            }
            return false;
        }

        // ── Clipping sync ─────────────────────────────────────────────────────

        private void PushPlanesToManager()
        {
            var defs = new List<SectionPlaneDef>(_planes.Count);
            foreach (var p in _planes)
                defs.Add(new SectionPlaneDef(p.Normal, p.Point));
            _mgr.SetPlanes(defs);
            _mgr.Enabled = _planes.Count > 0;
        }

        // ── Rect geometry ─────────────────────────────────────────────────────

        /// <summary>
        /// Project the scene bounding box onto the plane basis to size the placed
        /// rectangle (and the travel range along the normal).
        /// </summary>
        private void ComputeRectExtents(PlacedPlane plane)
        {
            BuildBasis(plane.Normal, out plane.AxisU, out plane.AxisV);

            var bb = _boundsProvider?.Invoke();
            if (bb == null || bb.Value.Maximum == bb.Value.Minimum)
            {
                // No geometry to measure — fall back to a 20 m rect at the plane point.
                plane.BbCentre = plane.Point;
                plane.MinU = plane.MinV = -10f; plane.MaxU = plane.MaxV = 10f;
                plane.MinN = -10f; plane.MaxN = 10f;
                return;
            }

            var box = bb.Value;
            plane.BbCentre = box.Center;

            plane.MinU = plane.MinV = plane.MinN = float.MaxValue;
            plane.MaxU = plane.MaxV = plane.MaxN = float.MinValue;
            foreach (var corner in box.GetCorners())
            {
                var rel = corner - plane.BbCentre;
                float du = Vector3.Dot(rel, plane.AxisU);
                float dv = Vector3.Dot(rel, plane.AxisV);
                float dn = Vector3.Dot(rel, plane.Normal);
                if (du < plane.MinU) plane.MinU = du; if (du > plane.MaxU) plane.MaxU = du;
                if (dv < plane.MinV) plane.MinV = dv; if (dv > plane.MaxV) plane.MaxV = dv;
                if (dn < plane.MinN) plane.MinN = dn; if (dn > plane.MaxN) plane.MaxN = dn;
            }

            // Small margin so the rect frames the model instead of hugging it.
            float padU = 0.03f * (plane.MaxU - plane.MinU) + 0.25f;
            float padV = 0.03f * (plane.MaxV - plane.MinV) + 0.25f;
            plane.MinU -= padU; plane.MaxU += padU;
            plane.MinV -= padV; plane.MaxV += padV;
        }

        /// <summary>Reposition a plane's rect onto its current plane.</summary>
        private void UpdateRectTransform(PlacedPlane plane)
        {
            // Project the bbox centre onto the plane, then offset to the rect centre.
            var cp = plane.BbCentre
                - Vector3.Dot(plane.BbCentre - plane.Point, plane.Normal) * plane.Normal;
            var centre = cp
                + plane.AxisU * ((plane.MinU + plane.MaxU) * 0.5f)
                + plane.AxisV * ((plane.MinV + plane.MaxV) * 0.5f);

            float w = plane.MaxU - plane.MinU;
            float h = plane.MaxV - plane.MinV;

            var xform = MakeTransform(plane.AxisU * w, plane.AxisV * h, plane.Normal, centre);
            plane.Fill.Transform   = xform;
            plane.Border.Transform = xform;
        }

        // ── Visual construction ───────────────────────────────────────────────

        private void BuildSharedVisuals()
        {
            // Accent blue, matching a BIMcollab-style section rect.
            var accent      = new Color3(0.30f, 0.65f, 1.00f);
            var accentHover = new Color3(0.45f, 0.78f, 1.00f);

            _previewMat   = MakeFlatMaterial(accent, alpha: 0.28f);
            _rectMatIdle  = MakeFlatMaterial(accent, alpha: 0.07f);
            _rectMatHover = MakeFlatMaterial(accentHover, alpha: 0.16f);

            _previewFill = new MeshGeometryModel3D
            {
                Geometry         = BuildUnitQuad(),
                Material         = _previewMat,
                IsTransparent    = true,
                IsHitTestVisible = false,
                CullMode         = SharpDX.Direct3D11.CullMode.None,
                DepthBias        = -60,
                Visibility       = System.Windows.Visibility.Collapsed,
            };
            _previewBorder = new LineGeometryModel3D
            {
                Geometry         = BuildUnitSquareOutline(),
                Color            = System.Windows.Media.Color.FromRgb(0x6C, 0xC2, 0xFF),
                Thickness        = 1.6,
                IsHitTestVisible = false,
                DepthBias        = -120,
                Visibility       = System.Windows.Visibility.Collapsed,
            };

            _visualRoot.Children.Add(_previewFill);
            _visualRoot.Children.Add(_previewBorder);
        }

        private void CreatePlaneVisuals(PlacedPlane plane)
        {
            plane.Fill = new MeshGeometryModel3D
            {
                Geometry         = BuildUnitQuad(),
                Material         = _rectMatIdle,
                IsTransparent    = true,
                IsHitTestVisible = true, // grabbable
                CullMode         = SharpDX.Direct3D11.CullMode.None,
                DepthBias        = -60,
                Visibility       = _visualsVisible
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed,
            };
            plane.Border = new LineGeometryModel3D
            {
                Geometry         = BuildUnitSquareOutline(),
                Color            = BorderIdleColor,
                Thickness        = 1.8,
                IsHitTestVisible = false,
                DepthBias        = -120,
                Visibility       = plane.Fill.Visibility,
            };
            _visualRoot.Children.Add(plane.Fill);
            _visualRoot.Children.Add(plane.Border);
        }

        private void HidePreview()
        {
            _previewFill.Visibility   = System.Windows.Visibility.Collapsed;
            _previewBorder.Visibility = System.Windows.Visibility.Collapsed;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        /// <summary>Flat, lighting-independent translucent material: colour comes
        /// from emissive, opacity from the diffuse alpha.</summary>
        private static PhongMaterial MakeFlatMaterial(Color3 colour, float alpha)
            => new PhongMaterial
            {
                DiffuseColor  = new Color4(0f, 0f, 0f, alpha),
                AmbientColor  = new Color4(0f, 0f, 0f, 1f),
                SpecularColor = new Color4(0f, 0f, 0f, 1f),
                EmissiveColor = new Color4(colour.Red, colour.Green, colour.Blue, 1f),
            };

        /// <summary>
        /// In-plane basis for a plane normal, in the viewer's Y-up world: U runs
        /// horizontally and V vertically for wall faces; for near-horizontal faces
        /// (floors/ceilings) U/V fall back to the world X/Z axes.
        /// </summary>
        private static void BuildBasis(Vector3 n, out Vector3 u, out Vector3 v)
        {
            var up = Vector3.UnitY;
            if (Math.Abs(Vector3.Dot(n, up)) > 0.95f)
            {
                v = Vector3.Normalize(Vector3.Cross(n, Vector3.UnitX));
                u = Vector3.Normalize(Vector3.Cross(v, n));
            }
            else
            {
                u = Vector3.Normalize(Vector3.Cross(up, n));
                v = Vector3.Normalize(Vector3.Cross(n, u));
            }
        }

        /// <summary>World transform mapping the unit quad/outline (XY, centred at
        /// the origin) onto the given basis and centre.</summary>
        private static Media3D.MatrixTransform3D MakeTransform(
            Vector3 xAxis, Vector3 yAxis, Vector3 zAxis, Vector3 origin)
        {
            var m = new Media3D.Matrix3D(
                xAxis.X,  xAxis.Y,  xAxis.Z,  0,
                yAxis.X,  yAxis.Y,  yAxis.Z,  0,
                zAxis.X,  zAxis.Y,  zAxis.Z,  0,
                origin.X, origin.Y, origin.Z, 1);
            return new Media3D.MatrixTransform3D(m);
        }

        private static MeshGeometry3D BuildUnitQuad()
        {
            return new MeshGeometry3D
            {
                Positions = new Vector3Collection
                {
                    new Vector3(-0.5f, -0.5f, 0),
                    new Vector3( 0.5f, -0.5f, 0),
                    new Vector3( 0.5f,  0.5f, 0),
                    new Vector3(-0.5f,  0.5f, 0),
                },
                Normals = new Vector3Collection
                {
                    Vector3.UnitZ, Vector3.UnitZ, Vector3.UnitZ, Vector3.UnitZ,
                },
                Indices = new IntCollection { 0, 1, 2, 0, 2, 3 },
            };
        }

        private static LineGeometry3D BuildUnitSquareOutline()
        {
            var lb = new LineBuilder();
            lb.AddLine(new Vector3(-0.5f, -0.5f, 0), new Vector3( 0.5f, -0.5f, 0));
            lb.AddLine(new Vector3( 0.5f, -0.5f, 0), new Vector3( 0.5f,  0.5f, 0));
            lb.AddLine(new Vector3( 0.5f,  0.5f, 0), new Vector3(-0.5f,  0.5f, 0));
            lb.AddLine(new Vector3(-0.5f,  0.5f, 0), new Vector3(-0.5f, -0.5f, 0));
            return lb.ToLineGeometry3D();
        }

        private static float Clamp(float value, float min, float max)
            => value < min ? min : (value > max ? max : value);
    }
}
