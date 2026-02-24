using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
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
    ///
    /// When the active Revit view is not a 3D view, <see cref="ResolveView3D"/> falls
    /// back to the saved preference (via <see cref="GetSavedViewCallback"/>) or shows
    /// the <see cref="PickView3DCallback"/> dialog so the user can pick a view once.
    /// The selection is saved via <see cref="SaveViewCallback"/> for future sessions.
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
        // volatile: written on thread-pool (debounce timer), read on Revit API thread.
        private volatile ExecMode _pendingMode = ExecMode.ManualSync;

        // ── 3D view picker callbacks (set by IfcViewerWindow after creation) ──
        /// <summary>
        /// Called on the UI thread (via Dispatcher.Invoke) when no 3D view is active
        /// and no saved preference exists.
        /// Arguments: (viewNames, currentlySavedName) — either may be empty/null.
        /// Returns the chosen view name, or <c>null</c> if the user cancels.
        /// </summary>
        public Func<string[], string, string> PickView3DCallback { get; set; }

        /// <summary>Persists (docPath, viewName) to the settings file.</summary>
        public Action<string, string> SaveViewCallback { get; set; }

        /// <summary>Returns the saved view name for a document path, or <c>null</c>.</summary>
        public Func<string, string> GetSavedViewCallback { get; set; }

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
                var doc = app.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    FireError(_onError, new InvalidOperationException("No active Revit document."));
                    return;
                }

                var view = ResolveView3D(doc);
                if (view == null)
                {
                    FireError(_onError, new InvalidOperationException(
                        "No 3D view selected. Please switch to a 3D view or pick one when prompted."));
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
                var doc = app.ActiveUIDocument?.Document;
                if (doc == null)
                {
                    FireError(_onAutoError, new InvalidOperationException("No active Revit document."));
                    return;
                }

                var view = ResolveView3D(doc);
                if (view == null)
                {
                    FireError(_onAutoError, new InvalidOperationException(
                        "No 3D view selected. Please switch to a 3D view or pick one when prompted."));
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

        // ── View resolution ───────────────────────────────────────────────────

        /// <summary>
        /// Returns a <see cref="View3D"/> to use for export, in priority order:
        /// 1. Active view (if already a non-template 3D view)
        /// 2. Saved preference from a previous session (looked up via <see cref="GetSavedViewCallback"/>)
        /// 3. The only available 3D view (auto-selected silently)
        /// 4. User picks via the <see cref="PickView3DCallback"/> dialog (result is saved)
        ///
        /// Runs on the Revit API thread; the picker dialog is marshalled to the UI thread
        /// via a blocking <c>Dispatcher.Invoke</c> so the Revit thread waits for the choice.
        /// </summary>
        private View3D ResolveView3D(Document doc)
        {
            if (doc == null) return null;

            // 1. Active view is already a suitable 3D view — use it directly.
            if (doc.ActiveView is View3D active && !active.IsTemplate)
                return active;

            // 2. Collect all available non-template 3D views.
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .OrderBy(v => v.Name)
                .ToList();

            if (views.Count == 0)
            {
                SessionLogger.Warn("ResolveView3D: document has no non-template 3D views.");
                return null;
            }

            // 3. Check saved preference for this document.
            string docKey    = doc.PathName ?? string.Empty;
            string savedName = string.IsNullOrEmpty(docKey)
                ? null : GetSavedViewCallback?.Invoke(docKey);

            if (!string.IsNullOrEmpty(savedName))
            {
                var savedView = views.FirstOrDefault(v => v.Name == savedName);
                if (savedView != null)
                {
                    SessionLogger.Info($"ResolveView3D: using saved view \"{savedName}\".");
                    return savedView;
                }
                // Saved name no longer exists — fall through to picker.
                SessionLogger.Warn($"ResolveView3D: saved view \"{savedName}\" not found; repicking.");
            }

            // 4. Only one 3D view — use it silently without prompting.
            if (views.Count == 1)
            {
                SessionLogger.Info($"ResolveView3D: single 3D view \"{views[0].Name}\" auto-selected.");
                if (!string.IsNullOrEmpty(docKey))
                    SaveViewCallback?.Invoke(docKey, views[0].Name);
                return views[0];
            }

            // 5. Multiple views and no saved preference — show picker on UI thread.
            if (PickView3DCallback == null)
            {
                SessionLogger.Warn("ResolveView3D: no PickView3DCallback set; cannot show picker.");
                return null;
            }

            var names  = views.Select(v => v.Name).ToArray();
            string chosen = null;
            // Dispatcher.Invoke blocks the Revit API thread while the WPF dialog runs
            // on the UI thread.  The dialog is purely WPF (no Revit API calls inside),
            // so there is no deadlock risk.
            _uiDispatcher.Invoke((Action)(() =>
            {
                chosen = PickView3DCallback(names, savedName);
            }));

            if (string.IsNullOrEmpty(chosen))
            {
                SessionLogger.Info("ResolveView3D: user cancelled view picker.");
                return null;
            }

            var picked = views.FirstOrDefault(v => v.Name == chosen);
            if (picked != null && !string.IsNullOrEmpty(docKey))
            {
                SaveViewCallback?.Invoke(docKey, chosen);
                SessionLogger.Info($"ResolveView3D: user picked \"{chosen}\" — saved for this project.");
            }
            return picked;
        }

        // ── DocumentChanged handler (Revit API thread) ────────────────────────

        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (!_autoSyncActive) return;

            // Filter to the document we're tracking.
            // Compare by PathName rather than ReferenceEquals: Revit COM interop can
            // return a fresh C# wrapper object per call even for the same document,
            // so ReferenceEquals always returns false and would discard every event.
            if (_trackedDoc != null)
            {
                var changedDoc = args.GetDocument();
                if (changedDoc == null) return;
                if (changedDoc.PathName != _trackedDoc.PathName) return;
            }

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
