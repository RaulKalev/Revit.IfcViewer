using System.ComponentModel;
using Autodesk.Revit.DB;

namespace ProSchedules.Models
{
    /// <summary>
    /// Represents a pending rename operation for a schedule element's parameter.
    /// </summary>
    public class ScheduleRenameItem : INotifyPropertyChanged
    {
        private string _original;
        private string _new;

        /// <summary>
        /// The ElementId of the Revit element to modify.
        /// </summary>
        public ElementId ElementId { get; set; }

        /// <summary>
        /// The name of the parameter being renamed.
        /// </summary>
        public string ParameterName { get; set; }

        /// <summary>
        /// The original parameter value.
        /// </summary>
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

        /// <summary>
        /// The new (transformed) parameter value.
        /// </summary>
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

        /// <summary>
        /// Whether this parameter is a type parameter (affects all elements of the type).
        /// </summary>
        public bool IsTypeParameter { get; set; }

        /// <summary>
        /// Display name for the element (for preview purposes).
        /// </summary>
        public string ElementName { get; set; }

        public ScheduleRenameItem()
        {
        }

        public ScheduleRenameItem(ElementId elementId, string parameterName, string original, bool isTypeParameter)
        {
            ElementId = elementId;
            ParameterName = parameterName;
            Original = original;
            New = original;
            IsTypeParameter = isTypeParameter;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
