using System.ComponentModel;

namespace ProSchedules.Models
{
    public class RenamePreviewItem : INotifyPropertyChanged
    {
        private string _original;
        private string _new;

        public string Original
        {
            get => _original;
            set
            {
                if (_original != value)
                {
                    _original = value;
                    OnPropertyChanged(nameof(Original));
                }
            }
        }

        public string New
        {
            get => _new;
            set
            {
                if (_new != value)
                {
                    _new = value;
                    OnPropertyChanged(nameof(New));
                }
            }
        }

        public SheetItem Sheet { get; set; }

        public RenamePreviewItem(SheetItem sheet, string original)
        {
            Sheet = sheet;
            Original = original;
            New = original;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
