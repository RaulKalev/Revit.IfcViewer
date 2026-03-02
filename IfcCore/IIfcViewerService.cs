using System;
using System.Threading;
using System.Threading.Tasks;

namespace IfcCore
{
    /// <summary>
    /// Service boundary for IFC model loading, querying, and lifecycle management.
    /// All xBIM logic lives behind this interface — UI/Revit code must
    /// communicate only through these methods.
    /// </summary>
    public interface IIfcViewerService : IDisposable
    {
        /// <summary>
        /// Loads and tessellates an IFC file. Returns portable geometry data
        /// and element metadata. The operation supports cancellation.
        /// </summary>
        /// <param name="filePath">Full path to the .ifc file.</param>
        /// <param name="cancellationToken">Token to abort the operation.</param>
        /// <returns>Load result containing geometry + metadata.</returns>
        Task<IfcLoadResult> LoadAsync(string filePath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Looks up an element by its IFC GlobalId.
        /// Returns null if no model is loaded or the GlobalId is not found.
        /// </summary>
        IfcElementData GetElementByGlobalId(string globalId);

        /// <summary>
        /// Returns the axis-aligned bounding box of an element identified by GlobalId.
        /// Returns null if not found.
        /// </summary>
        IfcBoundingBox GetBoundingBox(string globalId);

        /// <summary>
        /// Disposes the currently loaded model and releases all file locks.
        /// Safe to call multiple times.
        /// </summary>
        void DisposeModel();
    }
}
