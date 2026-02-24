using Autodesk.Revit.DB;
using System.ComponentModel;

namespace ProSchedules.Models
{
    public enum SheetItemState
    {
        ExistingInRevit,      // Loaded from Revit, has valid ElementId
        PendingCreation,      // In-memory preview, will be created on Apply
        PendingEdit          // Existing sheet with unsaved name/number changes
    }

    public class SheetItem : INotifyPropertyChanged
    {
        private string _name;
        private string _sheetNumber;
        private bool _isSelected;
        private SheetItemState _state;
        private bool _hasNumberConflict;
        private readonly System.Collections.Generic.Dictionary<string, string> _parameterValues =
            new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                    CheckForPendingEdit();
                }
            }
        }

        public string SheetNumber
        {
            get => _sheetNumber;
            set
            {
                if (_sheetNumber != value)
                {
                    _sheetNumber = value;
                    OnPropertyChanged(nameof(SheetNumber));
                    CheckForPendingEdit();
                }
            }
        }

        public ElementId Id { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public SheetItemState State
        {
            get => _state;
            set
            {
                if (_state != value)
                {
                    _state = value;
                    OnPropertyChanged(nameof(State));
                    OnPropertyChanged(nameof(HasUnsavedChanges));
                }
            }
        }

        public string OriginalSheetNumber { get; set; }
        public string OriginalName { get; set; }

        public bool HasNumberConflict
        {
            get => _hasNumberConflict;
            set
            {
                if (_hasNumberConflict != value)
                {
                    _hasNumberConflict = value;
                    OnPropertyChanged(nameof(HasNumberConflict));
                }
            }
        }

        public string ValidationError { get; set; }
        public ElementId SourceSheetId { get; set; }
        public DuplicateOptions DuplicateOptions { get; set; }
        public System.Collections.Generic.Dictionary<string, string> ParameterValues => _parameterValues;

        public bool HasUnsavedChanges => State == SheetItemState.PendingEdit || State == SheetItemState.PendingCreation;

        // Constructor for existing Revit sheets
        public SheetItem(ViewSheet sheet)
        {
            _name = sheet.Name;
            _sheetNumber = sheet.SheetNumber;
            Id = sheet.Id;
            IsSelected = false;
            State = SheetItemState.ExistingInRevit;
            OriginalSheetNumber = sheet.SheetNumber;
            OriginalName = sheet.Name;
        }

        // Constructor for pending sheets (to be created)
        public SheetItem(string sheetNumber, string name, ElementId sourceSheetId, DuplicateOptions options)
        {
            _sheetNumber = sheetNumber;
            _name = name;
            SourceSheetId = sourceSheetId;
            DuplicateOptions = options;
            Id = ElementId.InvalidElementId;
            State = SheetItemState.PendingCreation;
            OriginalSheetNumber = sheetNumber;
            OriginalName = name;
            IsSelected = false;
        }

        private void CheckForPendingEdit()
        {
            // Only update state if it's currently an existing sheet
            if (State == SheetItemState.ExistingInRevit &&
                (SheetNumber != OriginalSheetNumber || Name != OriginalName))
            {
                State = SheetItemState.PendingEdit;
            }
        }

        public void SetParameterValue(string name, string value)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            _parameterValues[name] = value ?? string.Empty;
            OnPropertyChanged(nameof(ParameterValues));
        }

        public void RemoveParameterValue(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (_parameterValues.Remove(name))
            {
                OnPropertyChanged(nameof(ParameterValues));
            }
        }

        public void ClearParameterValues(System.Collections.Generic.IEnumerable<string> names)
        {
            bool changed = false;
            if (names == null)
            {
                return;
            }

            foreach (var name in names)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (_parameterValues.Remove(name))
                {
                    changed = true;
                }
            }

            if (changed)
            {
                OnPropertyChanged(nameof(ParameterValues));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
