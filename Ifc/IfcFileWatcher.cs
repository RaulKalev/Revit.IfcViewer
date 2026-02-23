using System;
using System.IO;
using System.Threading;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// Watches a single IFC file for changes on disk and fires a debounced callback.
    /// The callback fires on a thread-pool thread — callers must marshal to the UI
    /// thread themselves (e.g. via <c>Dispatcher.BeginInvoke</c>).
    /// Call <see cref="Dispose"/> to stop watching.
    /// </summary>
    internal sealed class IfcFileWatcher : IDisposable
    {
        private readonly FileSystemWatcher _watcher;
        private readonly Action            _onChange;
        private readonly Timer             _debounceTimer;
        private const int DebounceMs = 2000;

        public IfcFileWatcher(string filePath, Action onChange)
        {
            _onChange = onChange ?? throw new ArgumentNullException(nameof(onChange));

            // Timer starts stopped; each file event resets the 2-second window.
            _debounceTimer = new Timer(_ => _onChange(), null, Timeout.Infinite, Timeout.Infinite);

            string dir  = Path.GetDirectoryName(filePath) ?? ".";
            string file = Path.GetFileName(filePath);

            _watcher = new FileSystemWatcher(dir, file)
            {
                NotifyFilter        = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
                EnableRaisingEvents = true,
            };

            _watcher.Changed += OnChanged;
            _watcher.Renamed += OnChanged;
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            // Restart the debounce window on every event.
            _debounceTimer.Change(DebounceMs, Timeout.Infinite);
        }

        public void Dispose()
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Changed -= OnChanged;
            _watcher.Renamed -= OnChanged;
            _watcher.Dispose();
            _debounceTimer.Dispose();
        }
    }
}
