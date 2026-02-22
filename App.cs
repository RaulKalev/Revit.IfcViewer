using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using IfcViewer.Commands;

namespace IfcViewer
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private RibbonPanel _ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "RK Tools";

            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // Tab already exists; continue.
            }

            _ribbonPanel = application.CreateOrSelectPanel(tabName, "Tools");

            _ribbonPanel.CreatePushButton<IfcViewerCommand>()
                .SetLargeImage("pack://application:,,,/IfcViewer;component/Assets/IfcViewer.png")
                .SetText("IFC\nViewer")
                .SetToolTip("Launch the IFC Viewer")
                .SetLongDescription("IfcViewer lets you load IFC files and compare them against your live Revit model in a GPU-accelerated 3D viewport.")
                .SetContextualHelp("https://github.com/RaulKalev/Revit.IfcViewer");

            SessionLogger.Info("IfcViewer ribbon loaded.");
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _ribbonPanel?.Remove();
            return Result.Succeeded;
        }
    }
}
