using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Threading;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Bridges the modeless WPF window to the Revit API thread via
    /// <see cref="IExternalEventHandler"/> / <see cref="ExternalEvent"/>.
    ///
    /// Manual sync: call <see cref="Request"/> to trigger a full export.
    /// Auto-sync: call <see cref="StartAutoSync"/> to subscribe to
    /// <c>Application.DocumentChanged</c> and perform incremental updates on every
    /// transaction that modifies geometry (debounced 800 ms).
    /// </summary>
    public sealed class SyncRevitEvent : IExternalEventHandler, IDisposable
    {
        private readonly ExternalEvent _event;
        private readonly Dispatcher    _uiDispatcher;

        // ── Manual sync ───────────────────────────────────────────────────────
        private Action<RevitModel> _onComplete;
        private Action<Exception>  _onError;

        // ── Auto-sync ─────────────────────────────────────────────────────────
        private bool   _autoSyncActive;
        private bool   _documentChangedSubscribed;
        private Autodesk.Revit.ApplicationServices.Application _revitApp;
        private Document _trackedDoc;
        private View3D   _trackedView;
        private Action<RevitModel> _onAutoUpdate;
        private Action<Exception>  _onAutoError;

        // Latest exported model used as the base for incremental patches.
        private RevitModel _liveModel;

        // ── Dirty tracking (DocumentChanged → debounce → Execute) ─────────────
        private readonly object             _dirtyLock  = new object();
        private readonly HashSet<ElementId> _dirtyIds   = new HashSet<ElementId>();
        private readonly HashSet<ElementId> _deletedIds = new HashSet<ElementId>();
        private Timer _debounceTimer;

        // ── Exec mode ─────────────────────────────────────────────────────────
        private enum ExecMode { ManualSync, InitAutoSync, IncrementalSync }
        private ExecMode _pendingMode = ExecMode.ManualSync;

        // ── Construction ──────────────────────────────────────────────────────

        public SyncRevitEvent(Dispatcher uiDispatcher)
        {
            _uiDispatcher = uiDispatcher;
            _event        = ExternalEvent.Create(this);
        }

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Schedules a full manual export. Both callbacks fire on the WPF UI thread.
        /// </summary>
        public void Request(Action<RevitModel> onComplete, Action<Exception> onError)
        {
            _onComplete  = onComplete;
            _onError     = onError;
            _pendingMode = ExecMode.ManualSync;
            _event.Raise();
        }

        /// <summary>
        /// Starts auto-sync: performs an initial full export then subscribes to
        /// <c>DocumentChanged</c> for incremental updates.
        /// </summary>
        public void StartAutoSync(Action<RevitModel> onUpdate, Action<Exception> onError)
        {
            _onAutoUpdate   = onUpdate;
            _onAutoError    = onError;
            _autoSyncActive = true;
            _debounceTimer  = new Timer(OnDebounceElapsed, null, Timeout.Infinite, Timeout.Infinite);
            _pendingMode    = ExecMode.InitAutoSync;
            _event.Raise();
        }

        /// <summary>
        /// Stops auto-sync. Pending DocumentChanged events are discarded.
        /// </summary>
        public void StopAutoSync()
        {
            _autoSyncActive = false;
            _debounceTimer?.Dispose();
            _debounceTimer = null;
        }

        // ── IExternalEventHandler ─────────────────────────────────────────────

        public string GetName() => "IfcViewer.SyncRevit";

        public void Execute(UIApplication app)
        {
            switch (_pendingMode)
            {
                case ExecMode.ManualSync:      ExecuteManualSync(app);      break;
                case ExecMode.InitAutoSync:    ExecuteInitAutoSync(app);    break;
                case ExecMode.IncrementalSync: ExecuteIncrementalSync(app); break;
            }
        }

        // ── Execute implementations ───────────────────────────────────────────

        private void ExecuteManualSync(UIApplication app)
        {
            try
            {
                var doc  = app.ActiveUIDocument?.Document;
                var view = doc?.ActiveView as View3D;

                if (doc == null)
                {
                    FireError(_onError, new InvalidOperationException("No active Revit document."));
                    return;
                }
                if (view == null)
                {
                    FireError(_onError, new InvalidOperationException(
                        "The active view must be a 3D view to export Revit geometry."));
                    return;
                }

                RevitModel model = RevitExporter.Export(doc, view, _uiDispatcher);
                _uiDispatcher.BeginInvoke((Action)(() =>
                {
                    _liveModel = model;
                    _onComplete?.Invoke(model);
                }));
            }
            catch (Exception ex)
            {
                SessionLogger.Error("SyncRevitEvent.ExecuteManualSync failed.", ex);
                FireError(_onError, ex);
            }
        }

        private void ExecuteInitAutoSync(UIApplication app)
        {
            try
            {
                var doc  = app.ActiveUIDocument?.Document;
                var view = doc?.ActiveView as View3D;

                if (doc == null || view == null)
                {
                    FireError(_onAutoError, new InvalidOperationException(
                        "Auto-sync requires an active 3D view."));
                    return;
                }

                _trackedDoc  = doc;
                _trackedView = view;

                // Subscribe to DocumentChanged once for the lifetime of auto-sync.
                if (!_documentChangedSubscribed)
                {
                    _revitApp = app.Application;
                    _revitApp.DocumentChanged += OnDocumentChanged;
                    _documentChangedSubscribed = true;
                    SessionLogger.Info("Auto-sync: DocumentChanged subscribed.");
                }

                RevitModel model = RevitExporter.Export(doc, view, _uiDispatcher);
                _uiDispatcher.BeginInvoke((Action)(() =>
                {
                    _liveModel = model;
                    _onAutoUpdate?.Invoke(model);
                }));
            }
            catch (Exception ex)
            {
                SessionLogger.Error("SyncRevitEvent.ExecuteInitAutoSync failed.", ex);
                FireError(_onAutoError, ex);
            }
        }

        private void ExecuteIncrementalSync(UIApplication app)
        {
            if (!_autoSyncActive) return;

            // Atomically copy and clear dirty sets.
            HashSet<ElementId> dirty, deleted;
            lock (_dirtyLock)
            {
                dirty   = new HashSet<ElementId>(_dirtyIds);
                deleted = new HashSet<ElementId>(_deletedIds);
                _dirtyIds.Clear();
                _deletedIds.Clear();
            }

            if (dirty.Count == 0 && deleted.Count == 0) return;

            try
            {
                var doc  = _trackedDoc;
                var view = _trackedView;
                if (doc == null || view == null || !_autoSyncActive) return;

                RevitModel updated = RevitExporter.ExportIncremental(
                    doc, view, dirty, deleted, _liveModel, _uiDispatcher);

                _uiDispatcher.BeginInvoke((Action)(() =>
                {
                    if (!_autoSyncActive) return;
                    _liveModel = updated;
                    _onAutoUpdate?.Invoke(updated);
                }));
            }
            catch (Exception ex)
            {
                SessionLogger.Error("SyncRevitEvent.ExecuteIncrementalSync failed.", ex);
                FireError(_onAutoError, ex);
            }
        }

        // ── DocumentChanged handler (Revit API thread) ────────────────────────

        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (!_autoSyncActive) return;

            // Filter to the document we're tracking.
            if (_trackedDoc != null && !ReferenceEquals(args.GetDocument(), _trackedDoc)) return;

            lock (_dirtyLock)
            {
                foreach (var id in args.GetAddedElementIds())    _dirtyIds.Add(id);
                foreach (var id in args.GetModifiedElementIds()) _dirtyIds.Add(id);
                foreach (var id in args.GetDeletedElementIds())  _deletedIds.Add(id);
            }

            // Reset the 800 ms debounce window.
            _debounceTimer?.Change(800, Timeout.Infinite);
        }

        // ── Debounce timer callback (thread-pool thread) ──────────────────────

        private void OnDebounceElapsed(object state)
        {
            if (!_autoSyncActive) return;
            _pendingMode = ExecMode.IncrementalSync;
            _event.Raise();
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void FireError(Action<Exception> handler, Exception ex)
            => _uiDispatcher.BeginInvoke((Action)(() => handler?.Invoke(ex)));

        public void Dispose()
        {
            _autoSyncActive = false;
            _debounceTimer?.Dispose();
            _debounceTimer = null;
            _event.Dispose();
            // DocumentChanged subscription is left in place — the _autoSyncActive flag
            // prevents any further processing. Full unsubscription would require the
            // Revit API thread and is not necessary for plugin correctness.
        }
    }
}
