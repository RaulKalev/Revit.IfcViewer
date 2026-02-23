using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows.Threading;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Bridges the modeless WPF window to the Revit API thread via
    /// <see cref="IExternalEventHandler"/> / <see cref="ExternalEvent"/>.
    ///
    /// Usage:
    ///   1. Create one instance and keep it alive for the window's lifetime.
    ///   2. Call <see cref="Request"/> from the WPF thread; it raises the event
    ///      and invokes <paramref name="onComplete"/> back on the WPF thread when done.
    /// </summary>
    public sealed class SyncRevitEvent : IExternalEventHandler
    {
        private readonly ExternalEvent _event;
        private readonly Dispatcher    _uiDispatcher;
        private Action<RevitModel>     _onComplete;
        private Action<Exception>      _onError;

        public SyncRevitEvent(Dispatcher uiDispatcher)
        {
            _uiDispatcher = uiDispatcher;
            _event        = ExternalEvent.Create(this);
        }

        /// <summary>
        /// Schedules a Revit-thread export. Callbacks fire on the WPF UI thread.
        /// </summary>
        public void Request(Action<RevitModel> onComplete, Action<Exception> onError)
        {
            _onComplete = onComplete;
            _onError    = onError;
            _event.Raise();
        }

        // ── IExternalEventHandler ─────────────────────────────────────────────

        public string GetName() => "IfcViewer.SyncRevit";

        public void Execute(UIApplication app)
        {
            try
            {
                Document doc  = app.ActiveUIDocument?.Document;
                View3D   view = doc?.ActiveView as View3D;

                if (doc == null)
                {
                    FireError(new InvalidOperationException("No active Revit document."));
                    return;
                }
                if (view == null)
                {
                    FireError(new InvalidOperationException(
                        "The active view must be a 3D view to export Revit geometry."));
                    return;
                }

                RevitModel model = RevitExporter.Export(doc, view, _uiDispatcher);

                _uiDispatcher.BeginInvoke((Action)(() => _onComplete?.Invoke(model)));
            }
            catch (Exception ex)
            {
                SessionLogger.Error("SyncRevitEvent.Execute failed.", ex);
                FireError(ex);
            }
        }

        private void FireError(Exception ex)
            => _uiDispatcher.BeginInvoke((Action)(() => _onError?.Invoke(ex)));

        public void Dispose() => _event.Dispose();
    }
}
