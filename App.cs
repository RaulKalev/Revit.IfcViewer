using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;

namespace IfcViewer
{
    public class App : IExternalApplication
    {
        private RibbonPanel _ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            // The DependencyLoader module initializer has already extracted the
            // embedded dependency DLLs and hooked assembly resolution before any
            // code in this assembly ran. This call is a no-op safety net.
            DependencyLoader.EnsureInitialized();
            SessionLogger.Info($"IfcViewer starting. Libs: {DependencyLoader.LibDir ?? "(loose files)"}");

            try
            {
                CreateRibbon(application);
            }
            catch (Exception ex)
            {
                SessionLogger.Error("Failed to create IfcViewer ribbon.", ex);
                return Result.Failed;
            }

            SessionLogger.Info("IfcViewer ribbon loaded.");
            return Result.Succeeded;
        }

        // Kept separate from OnStartup so the JIT only needs the ricaun.Revit.UI
        // assembly (resolved from the extracted libs) once this method is called —
        // after DependencyLoader is guaranteed to be initialized.
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private void CreateRibbon(UIControlledApplication application)
        {
            string tabName = "RK Tools";
            try { application.CreateRibbonTab(tabName); } catch { }

            _ribbonPanel = application.CreateOrSelectPanel(tabName, "Tools");

            _ribbonPanel.CreatePushButton<Commands.IfcViewerCommand>()
                .SetLargeImage("pack://application:,,,/IfcViewer;component/Assets/IfcViewer.tiff")
                .SetText("IFC\nViewer")
                .SetToolTip("Launch the IFC Viewer")
                .SetLongDescription("IfcViewer lets you load IFC files and compare them against your live Revit model in a GPU-accelerated 3D viewport.")
                .SetContextualHelp("https://github.com/RaulKalev/Revit.IfcViewer");
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _ribbonPanel?.Remove();
            return Result.Succeeded;
        }
    }
}
