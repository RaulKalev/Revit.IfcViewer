using Autodesk.Revit.UI;
using ricaun.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using IfcViewer.Commands;

namespace IfcViewer
{
    [AppLoader]
    public class App : IExternalApplication
    {
        private RibbonPanel _ribbonPanel;

        // Folder next to IfcViewer.dll — all dependency DLLs live here.
        private static string _assemblyDir;

        public Result OnStartup(UIControlledApplication application)
        {
            // ── 1. Register assembly resolver FIRST, before anything Helix-related ──
            _assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            SessionLogger.Info($"IfcViewer starting. Assembly dir: {_assemblyDir}");

            // ── 2. Ribbon ──────────────────────────────────────────────────────────
            string tabName = "RK Tools";
            try { application.CreateRibbonTab(tabName); } catch { }

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
            AppDomain.CurrentDomain.AssemblyResolve -= OnAssemblyResolve;
            _ribbonPanel?.Remove();
            return Result.Succeeded;
        }

        /// <summary>
        /// Probes the add-in folder for any DLL that satisfies the requested
        /// assembly name. This handles Helix, SharpDX, and any future deps
        /// that Revit's CLR host doesn't know about.
        /// </summary>
        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                // args.Name is like "HelixToolkit.Wpf.SharpDX, Version=2.25.0.0, ..."
                var simpleName = new AssemblyName(args.Name).Name;

                // 1. Already loaded? Return it to avoid duplicates.
                foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (loaded.GetName().Name.Equals(simpleName, StringComparison.OrdinalIgnoreCase))
                        return loaded;
                }

                // 2. Look for a matching DLL next to IfcViewer.dll
                var candidate = Path.Combine(_assemblyDir, simpleName + ".dll");
                if (File.Exists(candidate))
                {
                    SessionLogger.Info($"AssemblyResolve: loading '{simpleName}' from '{candidate}'");
                    return Assembly.LoadFrom(candidate);
                }

                SessionLogger.Warn($"AssemblyResolve: '{simpleName}' not found in '{_assemblyDir}'");
            }
            catch (Exception ex)
            {
                SessionLogger.Error($"AssemblyResolve exception for '{args.Name}'", ex);
            }

            return null;
        }
    }
}
