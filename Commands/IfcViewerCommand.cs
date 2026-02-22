using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Runtime.InteropServices;

namespace IfcViewer.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class IfcViewerCommand : IExternalCommand
    {
        private static UI.IfcViewerWindow _window;

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Surface existing window if already open.
                if (_window != null && _window.IsLoaded)
                {
                    var hwnd = new System.Windows.Interop.WindowInteropHelper(_window).Handle;
                    if (_window.WindowState == System.Windows.WindowState.Minimized)
                        ShowWindow(hwnd, SW_RESTORE);

                    _window.Activate();
                    _window.Focus();
                    SetForegroundWindow(hwnd);
                    return Result.Succeeded;
                }

                _window = new UI.IfcViewerWindow(commandData.Application);

                // Set Revit process as owner so the window doesn't hide behind Revit.
                var owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new System.Windows.Interop.WindowInteropHelper(_window) { Owner = owner };

                _window.Closed += (s, e) =>
                {
                    SessionLogger.Info("IfcViewer window closed.");
                    _window = null;
                };

                SessionLogger.Info("IfcViewer window opening.");
                _window.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                SessionLogger.Error("IfcViewerCommand.Execute failed.", ex);
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
