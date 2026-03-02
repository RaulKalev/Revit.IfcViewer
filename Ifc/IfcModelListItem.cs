using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace IfcViewer.Ifc
{
    /// <summary>
    /// UI list entry for an IFC file discovered in the selected folder.
    /// Tracks whether the file is currently loaded into the scene.
    /// </summary>
    public sealed class IfcModelListItem : INotifyPropertyChanged
    {
        private bool _isLoaded;

        public IfcModelListItem(string filePath)
        {
            FilePath = filePath;
        }

        public string FilePath { get; }

        public string DisplayName => Path.GetFileNameWithoutExtension(FilePath);

        public bool IsLoaded
        {
            get { return _isLoaded; }
            set
            {
                if (_isLoaded == value) return;
                _isLoaded = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(name));
        }

        public override string ToString() => DisplayName;
    }
}
