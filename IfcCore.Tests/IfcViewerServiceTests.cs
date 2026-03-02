using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xbim.Ifc;
using Xbim.Ifc4;
using Xbim.Ifc4.Kernel;
using Xbim.Ifc4.MeasureResource;
using Xbim.Ifc4.SharedBldgElements;
using Xbim.Common.Step21;
using Xbim.IO;

namespace IfcCore.Tests
{
    [TestClass]
    public class IfcViewerServiceTests
    {
        private string _testDir;

        [TestInitialize]
        public void Setup()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "IfcCoreTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDir);
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_testDir))
                    Directory.Delete(_testDir, true);
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Creates a minimal IFC4 file with one wall entity for testing.
        /// </summary>
        private string CreateTestIfc(string fileName = "test.ifc")
        {
            string path = Path.Combine(_testDir, fileName);

            using (var model = IfcStore.Create(XbimSchemaVersion.Ifc4, XbimStoreType.InMemoryModel))
            {
                using (var txn = model.BeginTransaction("Create test entities"))
                {
                    // Project
                    var project = model.Instances.New<IfcProject>();
                    project.Name = "TestProject";
                    project.GlobalId = Xbim.Ifc4.UtilityResource.IfcGloballyUniqueId.ConvertToBase64(Guid.NewGuid());

                    // Wall
                    var wall = model.Instances.New<IfcWall>();
                    wall.Name = "TestWall";
                    wall.GlobalId = Xbim.Ifc4.UtilityResource.IfcGloballyUniqueId.ConvertToBase64(Guid.NewGuid());

                    txn.Commit();
                }

                model.SaveAs(path);
            }

            return path;
        }

        // ── Test: Open IFC and count IfcProduct entities ─────────────────────

        [TestMethod]
        public async Task LoadAsync_ValidIfc_ReturnsSuccess()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            var result = await service.LoadAsync(ifcPath);

            Assert.IsTrue(result.Success, result.ErrorMessage);
            Assert.AreEqual("Ifc4", result.SchemaVersion);
            Assert.IsTrue(result.EntityCount > 0, "Should have entities");
        }

        // ── Test: Retrieve element by GlobalId ───────────────────────────────

        [TestMethod]
        public async Task GetElementByGlobalId_FindsWall()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            var result = await service.LoadAsync(ifcPath);
            Assert.IsTrue(result.Success, result.ErrorMessage);

            // Find the wall's GlobalId from result
            var wallInfo = result.ElementInfo.Values
                .FirstOrDefault(e => e.Type == "IfcWall");
            Assert.IsNotNull(wallInfo, "Expected an IfcWall in the model");

            var lookedUp = service.GetElementByGlobalId(wallInfo.GlobalId);
            Assert.IsNotNull(lookedUp);
            Assert.AreEqual("TestWall", lookedUp.Name);
            Assert.AreEqual("IfcWall", lookedUp.Type);
        }

        // ── Test: Dispose and reopen same file ───────────────────────────────

        [TestMethod]
        public async Task DisposeModel_ReleasesFileLocks_CanReopen()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            // Load once
            var result1 = await service.LoadAsync(ifcPath);
            Assert.IsTrue(result1.Success, result1.ErrorMessage);

            // Dispose model
            service.DisposeModel();

            // Verify lookup returns null after dispose
            var nullResult = service.GetElementByGlobalId("anything");
            Assert.IsNull(nullResult, "After dispose, lookup should return null");

            // Reopen same file — should work without file lock issues
            var result2 = await service.LoadAsync(ifcPath);
            Assert.IsTrue(result2.Success, result2.ErrorMessage);
            Assert.IsTrue(result2.EntityCount > 0);
        }

        // ── Test: Cancellation during load ───────────────────────────────────

        [TestMethod]
        public async Task LoadAsync_Cancelled_ReturnsFailure()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            // Cancel immediately
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            // Pre-cancelled token: Task.Run may throw TaskCanceledException
            // before LoadCore can catch it, or LoadCore catches it itself.
            try
            {
                var result = await service.LoadAsync(ifcPath, cts.Token);
                // If LoadCore catches it, we get a result
                Assert.IsFalse(result.Success);
            }
            catch (TaskCanceledException)
            {
                // Task.Run threw because token was already cancelled — expected
            }
            catch (OperationCanceledException)
            {
                // Also acceptable
            }
        }

        // ── Test: BoundingBox returns null for non-existent GlobalId ──────────

        [TestMethod]
        public async Task GetBoundingBox_NonExistent_ReturnsNull()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            var result = await service.LoadAsync(ifcPath);
            Assert.IsTrue(result.Success, result.ErrorMessage);

            var bb = service.GetBoundingBox("NON_EXISTENT_GUID");
            Assert.IsNull(bb, "Bounding box for non-existent GlobalId should be null");
        }

        // ── Test: LoadResult reports duration ─────────────────────────────────

        [TestMethod]
        public async Task LoadAsync_ReportsDuration()
        {
            string ifcPath = CreateTestIfc();
            using var service = new IfcViewerService();

            var result = await service.LoadAsync(ifcPath);
            Assert.IsTrue(result.Success, result.ErrorMessage);
            Assert.IsTrue(result.Duration.TotalMilliseconds > 0,
                "Duration should be positive");
        }
    }
}
