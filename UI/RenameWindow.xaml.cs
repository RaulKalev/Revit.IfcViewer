using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using ProSchedules.Models;

namespace ProSchedules.UI
{
    public partial class RenameWindow : Window
    {
        private readonly DuplicateSheetsWindow _parent;
        private readonly bool _isScheduleMode;
        private readonly List<System.Data.DataRowView> _selectedScheduleRows;
        private readonly List<SheetItem> _selectedSheets;
        private readonly ScheduleData _scheduleData;

        public ObservableCollection<RenameParameterOption> ParameterOptions { get; set; } = new ObservableCollection<RenameParameterOption>();
        public ObservableCollection<ScheduleRenameItem> SchedulePreviewItems { get; set; } = new ObservableCollection<ScheduleRenameItem>();
        public ObservableCollection<RenamePreviewItem> SheetPreviewItems { get; set; } = new ObservableCollection<RenamePreviewItem>();

        public event Action<List<ScheduleRenameItem>> OnScheduleRenameApply;
        public event Action<List<RenamePreviewItem>, string> OnSheetRenameApply;

        public RenameWindow(DuplicateSheetsWindow parent, List<System.Data.DataRowView> scheduleRows, ScheduleData scheduleData)
        {
            _parent = parent;
            _isScheduleMode = true;
            _selectedScheduleRows = scheduleRows;
            _scheduleData = scheduleData;

            InitializeComponent();
            DataContext = this;
            Owner = parent;

            InitializeScheduleMode();
        }

        public RenameWindow(DuplicateSheetsWindow parent, List<SheetItem> sheets)
        {
            _parent = parent;
            _isScheduleMode = false;
            _selectedSheets = sheets;

            InitializeComponent();
            DataContext = this;
            Owner = parent;

            InitializeSheetMode();
        }

        private void InitializeScheduleMode()
        {
            // Populate parameter options from schedule columns
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "Count", "TypeName" };
            foreach (var col in _scheduleData.Columns)
            {
                if (skipColumns.Contains(col)) continue;

                bool isType = _scheduleData.IsTypeParameter.ContainsKey(col) && 
                              _scheduleData.IsTypeParameter[col];

                ParameterOptions.Add(new RenameParameterOption
                {
                    Name = col,
                    IsTypeParameter = isType,
                    IsSheetParameter = false
                });
            }

            RenameParameterCombo.ItemsSource = ParameterOptions;
            if (ParameterOptions.Count > 0)
            {
                RenameParameterCombo.SelectedIndex = 0;
            }

            UpdatePreview();
        }

        private void InitializeSheetMode()
        {
            ParameterOptions.Add(new RenameParameterOption 
            { 
                Name = "Sheet Number", 
                IsSheetParameter = true,
                IsTypeParameter = false
            });
            ParameterOptions.Add(new RenameParameterOption 
            { 
                Name = "Sheet Name", 
                IsSheetParameter = true,
                IsTypeParameter = false
            });

            RenameParameterCombo.ItemsSource = ParameterOptions;
            RenameParameterCombo.SelectedIndex = 0;

            UpdatePreview();
        }

        private void RenameParameter_Changed(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void RenameText_Changed(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            var selectedOption = RenameParameterCombo?.SelectedItem as RenameParameterOption;
            if (selectedOption == null) return;

            string findText = FindTextBox?.Text ?? "";
            string replaceText = ReplaceTextBox?.Text ?? "";
            string prefix = PrefixTextBox?.Text ?? "";
            string suffix = SuffixTextBox?.Text ?? "";

            if (_isScheduleMode)
            {
                SchedulePreviewItems.Clear();
                string paramName = selectedOption.Name;

                foreach (var row in _selectedScheduleRows)
                {
                    string elementIdStr = row["ElementId"]?.ToString() ?? "";
                    if (!long.TryParse(elementIdStr, out long elemIdVal)) continue;

                    ElementId elemId = new ElementId(elemIdVal);
                    string original = row[paramName]?.ToString() ?? "";
                    string newValue = ApplyTransform(original, findText, replaceText, prefix, suffix);

                    var item = new ScheduleRenameItem(elemId, paramName, original, selectedOption.IsTypeParameter)
                    {
                        New = newValue,
                        ElementName = row["TypeName"]?.ToString() ?? $"Element {elemIdVal}"
                    };

                    SchedulePreviewItems.Add(item);
                }

                PreviewDataGrid.ItemsSource = SchedulePreviewItems;
            }
            else
            {
                SheetPreviewItems.Clear();
                bool isSheetNumber = selectedOption.Name == "Sheet Number";

                foreach (var sheet in _selectedSheets)
                {
                    string original = isSheetNumber ? sheet.SheetNumber : sheet.Name;
                    string newValue = ApplyTransform(original, findText, replaceText, prefix, suffix);

                    var item = new RenamePreviewItem(sheet, original) { New = newValue };
                    SheetPreviewItems.Add(item);
                }

                PreviewDataGrid.ItemsSource = SheetPreviewItems;
            }
        }

        private string ApplyTransform(string original, string find, string replace, string prefix, string suffix)
        {
            string result = original ?? "";

            if (!string.IsNullOrEmpty(find))
            {
                result = result.Replace(find, replace);
            }

            if (!string.IsNullOrEmpty(prefix))
            {
                result = prefix + result;
            }

            if (!string.IsNullOrEmpty(suffix))
            {
                result = result + suffix;
            }

            return result;
        }

        private List<ScheduleRenameItem> _pendingRenameItems;

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            var selectedOption = RenameParameterCombo?.SelectedItem as RenameParameterOption;

            if (_isScheduleMode)
            {
                var itemsToApply = SchedulePreviewItems
                    .Where(x => x.Original != x.New)
                    .ToList();

                if (itemsToApply.Count == 0)
                {
                    Close();
                    return;
                }

                // Check for type parameter warning
                bool hasTypeParams = itemsToApply.Any(x => x.IsTypeParameter);
                if (hasTypeParams)
                {
                    // Store items and show custom confirmation popup
                    _pendingRenameItems = itemsToApply;
                    ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Visible;
                    return;
                }

                OnScheduleRenameApply?.Invoke(itemsToApply);
            }
            else
            {
                var itemsToApply = SheetPreviewItems
                    .Where(x => x.Original != x.New)
                    .ToList();

                OnSheetRenameApply?.Invoke(itemsToApply, selectedOption?.Name);
            }

            Close();
        }

        private void ConfirmContinue_Click(object sender, RoutedEventArgs e)
        {
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            
            if (_pendingRenameItems != null)
            {
                OnScheduleRenameApply?.Invoke(_pendingRenameItems);
                _pendingRenameItems = null;
                Close();
            }
        }

        private void ConfirmCancel_Click(object sender, RoutedEventArgs e)
        {
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            _pendingRenameItems = null;
        }


        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
