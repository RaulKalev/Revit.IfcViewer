using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;

namespace IfcViewer.Revit
{
    /// <summary>
    /// Tracks Revit selection changes from a modeless window by polling during Idling.
    /// </summary>
    public sealed class FollowSelectionService : IExternalEventHandler, IDisposable
    {
        private readonly UIApplication _uiApp;
        private readonly Action<UIDocument, ElementId> _onPrimarySelectionChanged;
        private readonly ExternalEvent _stateApplyEvent;

        private bool _isEnabled;
        private bool _isSubscribed;
        private string _lastDocumentKey;
        private long? _lastSeenPrimaryIdValue;
        private ElementId _lastSeenPrimaryId;
        private long? _lastProcessedPrimaryIdValue;
        private DateTime _lastProcessedUtc = DateTime.MinValue;

        /// <summary>
        /// Minimum delay between camera requests.
        /// Effective range is clamped to 150-300 ms.
        /// </summary>
        public int DebounceMilliseconds { get; set; } = 220;

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;
                RequestStateApply();
            }
        }

        public FollowSelectionService(
            UIApplication uiApp,
            Action<UIDocument, ElementId> onPrimarySelectionChanged)
        {
            _uiApp = uiApp ?? throw new ArgumentNullException(nameof(uiApp));
            _onPrimarySelectionChanged = onPrimarySelectionChanged
                ?? throw new ArgumentNullException(nameof(onPrimarySelectionChanged));
            _stateApplyEvent = ExternalEvent.Create(this);
        }

        public void Start()
        {
            IsEnabled = true;
        }

        public void Stop()
        {
            IsEnabled = false;
        }

        public string GetName() => "IfcViewer.FollowSelectionService.StateApply";

        /// <summary>
        /// Runs on Revit API thread via ExternalEvent.
        /// Applies desired enable/disable state by subscribing or unsubscribing Idling.
        /// </summary>
        public void Execute(UIApplication app)
        {
            try
            {
                if (_isEnabled && !_isSubscribed)
                {
                    app.Idling += OnIdling;
                    _isSubscribed = true;
                    ResetTracking();
                    SessionLogger.Info("FollowSelectionService started.");
                }
                else if (!_isEnabled && _isSubscribed)
                {
                    app.Idling -= OnIdling;
                    _isSubscribed = false;
                    ResetTracking();
                    SessionLogger.Info("FollowSelectionService stopped.");
                }
            }
            catch (Exception ex)
            {
                SessionLogger.Error("Failed to apply FollowSelectionService state.", ex);
            }
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (!_isEnabled) return;

            UIDocument uiDoc = _uiApp.ActiveUIDocument;
            if (uiDoc == null)
            {
                ResetTracking();
                return;
            }

            Document doc = uiDoc.Document;
            if (doc == null) return;

            string documentKey = (doc.PathName ?? string.Empty) + "|" + (doc.Title ?? string.Empty);
            if (!string.Equals(_lastDocumentKey, documentKey, StringComparison.Ordinal))
            {
                _lastDocumentKey = documentKey;
                _lastSeenPrimaryIdValue = null;
                _lastSeenPrimaryId = null;
                _lastProcessedPrimaryIdValue = null;
                _lastProcessedUtc = DateTime.MinValue;
            }

            ElementId primaryId = GetPrimarySelection(uiDoc.Selection.GetElementIds());
            long? primaryIdValue = ToElementIdValue(primaryId);

            if (!primaryIdValue.HasValue)
            {
                _lastSeenPrimaryIdValue = null;
                _lastSeenPrimaryId = null;
                _lastProcessedPrimaryIdValue = null;
                return;
            }

            if (_lastSeenPrimaryIdValue != primaryIdValue.Value)
            {
                _lastSeenPrimaryIdValue = primaryIdValue.Value;
                _lastSeenPrimaryId = primaryId;
            }

            if (_lastProcessedPrimaryIdValue == _lastSeenPrimaryIdValue) return;

            int debounceMs = Math.Max(150, Math.Min(300, DebounceMilliseconds));
            if ((DateTime.UtcNow - _lastProcessedUtc).TotalMilliseconds < debounceMs) return;

            _lastProcessedUtc = DateTime.UtcNow;
            _lastProcessedPrimaryIdValue = _lastSeenPrimaryIdValue;

            try
            {
                _onPrimarySelectionChanged(uiDoc, _lastSeenPrimaryId);
            }
            catch (Exception ex)
            {
                SessionLogger.Error("FollowSelectionService callback failed.", ex);
            }
        }

        private static ElementId GetPrimarySelection(ICollection<ElementId> ids)
        {
            if (ids == null || ids.Count == 0) return null;

            foreach (ElementId id in ids)
            {
                if (id != null && id != ElementId.InvalidElementId)
                    return id;
            }

            return null;
        }

        private static long? ToElementIdValue(ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId) return null;
            return id.Value;
        }

        private void ResetTracking()
        {
            _lastDocumentKey = null;
            _lastSeenPrimaryIdValue = null;
            _lastSeenPrimaryId = null;
            _lastProcessedPrimaryIdValue = null;
            _lastProcessedUtc = DateTime.MinValue;
        }

        private void RequestStateApply()
        {
            try
            {
                _stateApplyEvent.Raise();
            }
            catch (Exception ex)
            {
                // Best-effort: if the request is already pending, Revit throws here.
                // The pending request will still apply the latest _isEnabled value.
                SessionLogger.Warn("FollowSelectionService state-apply request deferred: " + ex.Message);
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
