using HelixToolkit.Wpf.SharpDX;
using SharpDX;
using System;
using System.Collections.Generic;

namespace IfcViewer.Viewer
{
    /// <summary>
    /// Extracts hard edges from a tessellated mesh for a clean wireframe overlay.
    ///
    /// Helix's built-in <c>RenderWireframe</c> draws every triangle edge, including
    /// the interior diagonals that tessellators insert to triangulate flat polygons.
    /// This helper computes a <see cref="LineGeometry3D"/> that contains only the
    /// <em>hard</em> edges — edges where two adjacent triangles meet at a dihedral
    /// angle above <paramref name="thresholdDegrees"/> — plus any boundary edges
    /// (edges touching only one triangle). The result looks like a structural outline
    /// rather than a wireframe mesh.
    /// </summary>
    public static class WireframeHelper
    {
        /// <summary>
        /// Extracts hard edges from a mesh.
        /// </summary>
        /// <param name="mesh">Source HelixToolkit mesh (Positions + Indices required).</param>
        /// <param name="thresholdDegrees">
        /// Minimum dihedral angle (degrees) for an edge to be considered hard.
        /// Default 30°: flat tessellation seams (~0°) are suppressed; right-angle
        /// building corners (90°) and curved-surface facets are retained.
        /// </param>
        /// <returns>
        /// A <see cref="LineGeometry3D"/> with one segment per hard edge, or
        /// <c>null</c> if the mesh has no triangles or no hard edges.
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

            // edge key (canonical vertex-pair) → list of face normals for adjacent tris.
            // SharpDX.Vector3 is a value type with exact float equality, which is safe
            // here because positions for a shared edge originate from the same tessellator
            // call and therefore have bit-identical float values.
            var edgeNormals = new Dictionary<(Vector3, Vector3), List<Vector3>>(triCount * 3);

            for (int t = 0; t < triCount; t++)
            {
                Vector3 p0 = positions[indices[t * 3]];
                Vector3 p1 = positions[indices[t * 3 + 1]];
                Vector3 p2 = positions[indices[t * 3 + 2]];

                Vector3 n = Vector3.Normalize(Vector3.Cross(p1 - p0, p2 - p0));
                if (float.IsNaN(n.X)) continue; // degenerate triangle — skip

                AddEdge(edgeNormals, p0, p1, n);
                AddEdge(edgeNormals, p1, p2, n);
                AddEdge(edgeNormals, p2, p0, n);
            }

            float cosThreshold = (float)Math.Cos(thresholdDegrees * Math.PI / 180.0);

            var linePts = new Vector3Collection();
            var lineIdx = new IntCollection();

            foreach (var kvp in edgeNormals)
            {
                var normals = kvp.Value;
                bool hard   = normals.Count == 1; // boundary edge → always hard

                if (!hard)
                {
                    // Check every pair of adjacent face normals.
                    for (int i = 0; i < normals.Count && !hard; i++)
                        for (int j = i + 1; j < normals.Count && !hard; j++)
                            if (Vector3.Dot(normals[i], normals[j]) < cosThreshold)
                                hard = true;
                }

                if (hard)
                {
                    int idx = linePts.Count;
                    linePts.Add(kvp.Key.Item1);
                    linePts.Add(kvp.Key.Item2);
                    lineIdx.Add(idx);
                    lineIdx.Add(idx + 1);
                }
            }

            if (linePts.Count == 0) return null;

            return new LineGeometry3D { Positions = linePts, Indices = lineIdx };
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static void AddEdge(
            Dictionary<(Vector3, Vector3), List<Vector3>> map,
            Vector3 a, Vector3 b, Vector3 faceNormal)
        {
            var key = CanonicalEdge(a, b);
            if (!map.TryGetValue(key, out var list))
                map[key] = list = new List<Vector3>(2);
            list.Add(faceNormal);
        }

        /// <summary>
        /// Returns an edge key with a deterministic ordering so that
        /// (A→B) and (B→A) map to the same dictionary entry.
        /// Ordering: first by X, then Y, then Z.
        /// </summary>
        private static (Vector3, Vector3) CanonicalEdge(Vector3 a, Vector3 b)
        {
            if (a.X != b.X) return a.X < b.X ? (a, b) : (b, a);
            if (a.Y != b.Y) return a.Y < b.Y ? (a, b) : (b, a);
            return a.Z <= b.Z ? (a, b) : (b, a);
        }
    }
}
