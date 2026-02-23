using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Extracts hard edges from a tessellated mesh for a clean wireframe overlay.
    ///
    /// Problem with <c>RenderWireframe = true</c>: Helix draws every triangle edge,
    /// including the interior diagonals tessellators insert to triangulate flat quads.
    ///
    /// This helper computes a <see cref="LineGeometry3D"/> containing only the
    /// <em>hard</em> edges — edges shared by two faces whose dihedral angle exceeds
    /// <paramref name="thresholdDegrees"/> — plus boundary edges (touching one face).
    ///
    /// <para><b>Key design: snap-grid position keys.</b>
    /// Revit's <c>CustomExporter</c> fires a separate <c>OnPolymesh</c> call per face.
    /// Each call provides its own vertex list, so adjacent faces store independent copies
    /// of shared-edge vertices. After the feet→metres coordinate transform, floating-point
    /// arithmetic can produce slightly different bit patterns for theoretically identical
    /// positions. Exact <c>Vector3</c> equality therefore fails to match the two copies,
    /// making every edge look like a boundary edge (→ all edges emitted).
    /// We fix this by snapping all positions to a 0.1 mm grid and using integer
    /// <c>(x, y, z)</c> tuples as dictionary keys; integer equality is exact.</para>
    /// </summary>
    public static class WireframeHelper
    {
        /// <summary>
        /// Snap grid in metres (0.1 mm). Large enough to merge floating-point duplicates
        /// from separate tessellation calls; small enough to preserve all BIM geometry.
        /// </summary>
        private const float SnapGrid = 0.0001f;

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Extracts hard edges from <paramref name="mesh"/>.
        /// </summary>
        /// <param name="mesh">Source HelixToolkit mesh (Positions + Indices required).</param>
        /// <param name="thresholdDegrees">
        /// Minimum dihedral angle for an edge to be considered hard (default 30°).
        /// Flat tessellation seams (~0°) are suppressed; right-angle corners (90°) and
        /// sharply curved surface facets are retained.
        /// </param>
        /// <returns>
        /// A <see cref="LineGeometry3D"/> with one segment per hard edge, or
        /// <c>null</c> if the mesh has no triangles or produces no hard edges.
        /// </returns>
        public static LineGeometry3D ExtractHardEdges(
            MeshGeometry3D mesh,
            float thresholdDegrees = 30f)
        {
            if (mesh?.Indices == null || mesh.Positions == null || mesh.Indices.Count < 3)
                return null;

            var positions = mesh.Positions;
            var indices   = mesh.Indices;
            int triCount  = indices.Count / 3;

            // Map: canonical (snapped-int-key, snapped-int-key) → edge data.
            // Using snapped integer keys instead of raw Vector3 so that adjacent triangles
            // from different tessellation calls (which may differ by 1–2 ULPs after
            // coordinate transforms) are correctly recognised as sharing an edge.
            var edgeMap = new Dictionary<EdgeKey, EdgeData>(triCount * 3);

            for (int t = 0; t < triCount; t++)
            {
                Vector3 p0 = positions[indices[t * 3]];
                Vector3 p1 = positions[indices[t * 3 + 1]];
                Vector3 p2 = positions[indices[t * 3 + 2]];

                Vector3 n = Vector3.Normalize(Vector3.Cross(p1 - p0, p2 - p0));
                if (float.IsNaN(n.X)) continue; // degenerate triangle

                AddEdge(edgeMap, p0, p1, n);
                AddEdge(edgeMap, p1, p2, n);
                AddEdge(edgeMap, p2, p0, n);
            }

            float cosThreshold = (float)Math.Cos(thresholdDegrees * Math.PI / 180.0);

            var linePts = new Vector3Collection();
            var lineIdx = new IntCollection();

            foreach (var kvp in edgeMap)
            {
                var data    = kvp.Value;
                var normals = data.Normals;
                bool hard   = normals.Count == 1; // boundary edge → always hard

                if (!hard)
                    for (int i = 0; i < normals.Count && !hard; i++)
                        for (int j = i + 1; j < normals.Count && !hard; j++)
                            if (Vector3.Dot(normals[i], normals[j]) < cosThreshold)
                                hard = true;

                if (hard)
                {
                    int idx = linePts.Count;
                    linePts.Add(data.A);
                    linePts.Add(data.B);
                    lineIdx.Add(idx);
                    lineIdx.Add(idx + 1);
                }
            }

            if (linePts.Count == 0) return null;
            return new LineGeometry3D { Positions = linePts, Indices = lineIdx };
        }

        // ── Internals ─────────────────────────────────────────────────────────

        private static void AddEdge(
            Dictionary<EdgeKey, EdgeData> map,
            Vector3 a, Vector3 b, Vector3 faceNormal)
        {
            var key = new EdgeKey(a, b);
            if (!map.TryGetValue(key, out var data))
                map[key] = data = new EdgeData(a, b);
            data.Normals.Add(faceNormal);
        }

        // ── Snap grid helpers ─────────────────────────────────────────────────

        private static (int, int, int) Snap(Vector3 v)
        {
            // Use double arithmetic so the multiply doesn't lose precision for
            // large coordinate values (building-level coords can be ±1000 m).
            const double inv = 1.0 / SnapGrid;
            return ((int)Math.Round(v.X * inv),
                    (int)Math.Round(v.Y * inv),
                    (int)Math.Round(v.Z * inv));
        }

        // ── Helper types ──────────────────────────────────────────────────────

        /// <summary>
        /// Canonical, order-independent edge key using snapped integer coordinates.
        /// Two keys are equal iff both endpoints snap to the same grid cell, regardless
        /// of which endpoint is "first".
        /// </summary>
        private readonly struct EdgeKey : IEquatable<EdgeKey>
        {
            private readonly (int, int, int) _a;
            private readonly (int, int, int) _b;

            public EdgeKey(Vector3 a, Vector3 b)
            {
                var sa = Snap(a);
                var sb = Snap(b);
                // Enforce a canonical (a ≤ b) ordering so (A→B) and (B→A) produce
                // the same key. Compare lexicographically on the three int components.
                if (Compare(sa, sb) <= 0) { _a = sa; _b = sb; }
                else                      { _a = sb; _b = sa; }
            }

            private static int Compare((int x, int y, int z) p, (int x, int y, int z) q)
            {
                int c = p.x.CompareTo(q.x);
                if (c != 0) return c;
                c = p.y.CompareTo(q.y);
                if (c != 0) return c;
                return p.z.CompareTo(q.z);
            }

            public bool Equals(EdgeKey other) => _a == other._a && _b == other._b;
            public override bool Equals(object obj) => obj is EdgeKey k && Equals(k);
            public override int GetHashCode()
            {
                unchecked
                {
                    int h = _a.GetHashCode();
                    h = h * 397 ^ _b.GetHashCode();
                    return h;
                }
            }
        }

        /// <summary>Stores the original (un-snapped) endpoint positions and adjacent face normals.</summary>
        private sealed class EdgeData
        {
            public readonly Vector3 A, B;
            public readonly List<Vector3> Normals = new List<Vector3>(2);
            public EdgeData(Vector3 a, Vector3 b) { A = a; B = b; }
        }
    }
}
