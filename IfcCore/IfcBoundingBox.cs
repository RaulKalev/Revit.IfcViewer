namespace IfcCore
{
    /// <summary>
    /// Simple axis-aligned bounding box (framework-agnostic, no SharpDX dependency).
    /// </summary>
    public class IfcBoundingBox
    {
        public float MinX { get; set; }
        public float MinY { get; set; }
        public float MinZ { get; set; }
        public float MaxX { get; set; }
        public float MaxY { get; set; }
        public float MaxZ { get; set; }

        public IfcBoundingBox() { }

        public IfcBoundingBox(float minX, float minY, float minZ,
                              float maxX, float maxY, float maxZ)
        {
            MinX = minX; MinY = minY; MinZ = minZ;
            MaxX = maxX; MaxY = maxY; MaxZ = maxZ;
        }

        /// <summary>Centre point of the bounding box.</summary>
        public void GetCenter(out float cx, out float cy, out float cz)
        {
            cx = (MinX + MaxX) * 0.5f;
            cy = (MinY + MaxY) * 0.5f;
            cz = (MinZ + MaxZ) * 0.5f;
        }
    }
}
