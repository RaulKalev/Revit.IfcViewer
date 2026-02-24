using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using ProSchedules.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Data;
using ProSchedules.ExternalEvents;

namespace ProSchedules.UI
{
    public class SortItem : System.ComponentModel.INotifyPropertyChanged
    {
        private string _selectedColumn;
        public string SelectedColumn
        {
            get => _selectedColumn;
            set { _selectedColumn = value; OnPropertyChanged(nameof(SelectedColumn)); }
        }

        private bool _isAscending = true;
        public bool IsAscending
        {
            get => _isAscending;
            set { _isAscending = value; OnPropertyChanged(nameof(IsAscending)); }
        }
        
        // Visual placeholders matching screenshot
        public bool ShowHeader { get; set; }
        public bool ShowFooter { get; set; }
        public bool ShowBlankLine { get; set; }

        public SortItem Clone()
        {
            return new SortItem
            {
                SelectedColumn = this.SelectedColumn,
                IsAscending = this.IsAscending,
                ShowHeader = this.ShowHeader,
                ShowFooter = this.ShowFooter,
                ShowBlankLine = this.ShowBlankLine
            };
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
    }

    public class InverseBooleanConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool?) && targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return !(bool)value;
        }
    }

    /// <summary>
    /// Represents an option in the Rename Parameter dropdown.
    /// </summary>
    public class RenameParameterOption
    {
        public string Name { get; set; }
        public bool IsTypeParameter { get; set; }
        public bool IsSheetParameter { get; set; } // True for Sheet Number/Sheet Name
    }

    public partial class DuplicateSheetsWindow : Window
    {
        #region Constants / PInvoke

        private const string ConfigFilePath = @"C:\ProgramData\RK Tools\ProSchedules\config.json";
        private const string WindowLeftKey = "DuplicateSheetsWindow.Left";
        private const string WindowTopKey = "DuplicateSheetsWindow.Top";
        private const string WindowWidthKey = "DuplicateSheetsWindow.Width";
        private const string WindowHeightKey = "DuplicateSheetsWindow.Height";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Revit state / UI state

        public static readonly DependencyProperty IsCopyModeProperty =
            DependencyProperty.Register("IsCopyMode", typeof(bool), typeof(DuplicateSheetsWindow), 
            new PropertyMetadata(false, (d, e) => ((DuplicateSheetsWindow)d).UpdateSelectionAdorner()));

        public bool IsCopyMode
        {
            get { return (bool)GetValue(IsCopyModeProperty); }
            set { SetValue(IsCopyModeProperty, value); }
        }

        private List<SheetItem> _allSheets;
        private Action _onPopupClose;
        private Action _onConfirmAction;
        private Action _onCancelAction;
        private ExternalEvent _externalEvent;
        private ExternalEvents.SheetDuplicationHandler _handler;
        private ExternalEvent _editExternalEvent;
        private ExternalEvents.SheetEditHandler _editHandler;


        private readonly WindowResizer _windowResizer;
        private bool _isDarkMode = true;
        private bool _isDataLoaded;
        private UIApplication _uiApplication;
        private RevitService _revitService;
        private ExternalEvent _parameterRenameExternalEvent;
        private ExternalEvents.ParameterRenameHandler _parameterRenameHandler;
        private ExternalEvent _scheduleFieldsExternalEvent;
        private ExternalEvents.ScheduleFieldsHandler _scheduleFieldsHandler;
        private ExternalEvent _parameterLoadExternalEvent;
        private ExternalEvents.ParameterDataLoadHandler _parameterLoadHandler;
        private ExternalEvent _highlightInModelExternalEvent;
        private ExternalEvents.HighlightInModelHandler _highlightInModelHandler;

        private ExternalEvent _parameterValueUpdateExternalEvent;
        private ExternalEvents.ParameterValueUpdateHandler _parameterValueUpdateHandler;

        public ObservableCollection<SheetItem> Sheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<SheetItem> FilteredSheets { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<RenamePreviewItem> RenamePreviewItems { get; set; } = new ObservableCollection<RenamePreviewItem>();
        public ObservableCollection<SortItem> SortCriteria { get; set; } = new ObservableCollection<SortItem>();
        public ObservableCollection<string> AvailableSortColumns { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<ScheduleRenameItem> ScheduleRenamePreviewItems { get; set; } = new ObservableCollection<ScheduleRenameItem>();
        public ObservableCollection<RenameParameterOption> RenameParameterOptions { get; set; } = new ObservableCollection<RenameParameterOption>();

        private ProSchedules.Models.ScheduleData _currentScheduleData;
        private Dictionary<ElementId, bool> _scheduleItemizeSettings = new Dictionary<ElementId, bool>();
        private System.Data.DataTable _rawScheduleData;
        private Dictionary<ElementId, ObservableCollection<SortItem>> _scheduleSortSettings = new Dictionary<ElementId, ObservableCollection<SortItem>>();
        private string _lastSelectedScheduleName;



        #endregion

        #region Ctor / Init

        public DuplicateSheetsWindow(UIApplication app)
        {
            _uiApplication = app;
            InitializeComponent();
            DataContext = this;

            // Create duplication handler
            _handler = new ExternalEvents.SheetDuplicationHandler();
            _handler.OnDuplicationFinished += OnDuplicationFinished;
            _externalEvent = ExternalEvent.Create(_handler);

            // Create edit handler
            _editHandler = new ExternalEvents.SheetEditHandler();
            _editHandler.OnEditFinished += OnEditFinished;
            _editExternalEvent = ExternalEvent.Create(_editHandler);

            _parameterRenameHandler = new ExternalEvents.ParameterRenameHandler();
            _parameterRenameHandler.OnRenameFinished += OnParameterRenameFinished;
            _parameterRenameExternalEvent = ExternalEvent.Create(_parameterRenameHandler);

            // Create schedule fields handler
            _scheduleFieldsHandler = new ExternalEvents.ScheduleFieldsHandler();
            _scheduleFieldsHandler.OnUpdateFinished += OnScheduleFieldsUpdateFinished;
            _scheduleFieldsExternalEvent = ExternalEvent.Create(_scheduleFieldsHandler);

            // Create parameter load handler
            _parameterLoadHandler = new ExternalEvents.ParameterDataLoadHandler();
            _parameterLoadHandler.OnDataLoaded += OnParameterDataLoaded;
            _parameterLoadExternalEvent = ExternalEvent.Create(_parameterLoadHandler);

            // Create parameter value update handler
            _parameterValueUpdateHandler = new ExternalEvents.ParameterValueUpdateHandler();
            _parameterValueUpdateHandler.OnUpdateFinished += OnParameterValueUpdateFinished;
            _parameterValueUpdateExternalEvent = ExternalEvent.Create(_parameterValueUpdateHandler);

            // Create highlight in model handler
            _highlightInModelHandler = new ExternalEvents.HighlightInModelHandler();
            _highlightInModelExternalEvent = ExternalEvent.Create(_highlightInModelHandler);


            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            DeferWindowShow();

            // Window infrastructure
            _windowResizer = new WindowResizer(this);
            Closed += MainWindow_Closed;

            // Enforce Cell selection (Excel-like)
            SheetsDataGrid.SelectionUnit = DataGridSelectionUnit.Cell;
            SheetsDataGrid.SelectionMode = DataGridSelectionMode.Extended;
            
            // Selection Adorner Events
            SheetsDataGrid.SelectedCellsChanged += (s, e) => UpdateSelectionAdorner();
            SheetsDataGrid.LayoutUpdated += (s, e) => UpdateSelectionAdorner();
            SheetsDataGrid.SizeChanged += (s, e) => UpdateSelectionAdorner();
            SheetsDataGrid.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, e) => UpdateSelectionAdorner()));

            // Debug: Show selection count in title
            SheetsDataGrid.SelectedCellsChanged += (s, e) => {
                // this.Title = $"Debug: Selected Cells = {SheetsDataGrid.SelectedCells.Count}";
            };

            MouseLeftButtonUp += Window_MouseLeftButtonUp;

            // Theme
            LoadThemeState();
            LoadWindowState();
            DataContext = this;

            // Load persistent settings (must be before LoadData so they're available when schedule is restored)
            LoadSortSettings();
            
            LoadData(app.ActiveUIDocument.Document);

            // Check for updates after window loads
            Loaded += (s, e) => 
            {
                Dispatcher.BeginInvoke(new Action(() => 
                {
                    Services.UpdateLogService.CheckAndShow(this);
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            };
        }


        private void MainWindow_Closed(object sender, EventArgs e)
        {
            SaveSortSettings();
            SaveWindowState();
        }

        private void LoadData(Document doc)
        {
            _revitService = new RevitService(doc);
            
            var collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));
            _allSheets = new List<SheetItem>();

            foreach (ViewSheet sheet in collector)
            {
                if (sheet.IsPlaceholder) continue;
                var sheetItem = new SheetItem(sheet);
                sheetItem.State = SheetItemState.ExistingInRevit;
                sheetItem.OriginalSheetNumber = sheet.SheetNumber;
                sheetItem.OriginalName = sheet.Name;
                sheetItem.PropertyChanged += OnSheetPropertyChanged;
                _allSheets.Add(sheetItem);
            }

            // Initially show all
            Sheets.Clear();
            FilteredSheets.Clear();
            foreach (var s in _allSheets.OrderBy(s => s.SheetNumber).ThenBy(s => s.Name))
            {
                Sheets.Add(s);
                FilteredSheets.Add(s);
            }

            // Load Schedules
            var schedules = _revitService.GetSchedules();
            var comboItems = new List<ScheduleOption>();
            comboItems.Add(new ScheduleOption { Name = "No Schedules Selected", Id = ElementId.InvalidElementId, Schedule = null });
            foreach(var s in schedules)
            {
                comboItems.Add(new ScheduleOption { Name = s.Name, Id = s.Id, Schedule = s });
            }
            SchedulesComboBox.ItemsSource = comboItems;
            
            // Try to restore last selected schedule from config
            string savedScheduleName = GetSavedScheduleName();
            if (!string.IsNullOrEmpty(savedScheduleName))
            {
                var savedItem = comboItems.FirstOrDefault(x => x.Name == savedScheduleName);
                if (savedItem != null)
                {
                    SchedulesComboBox.SelectedItem = savedItem;
                }
                else
                {
                    SchedulesComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                SchedulesComboBox.SelectedIndex = 0;
            }

            UpdateButtonStates();
            _isDataLoaded = true;
            TryShowWindow();
        }

        private void SchedulesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            
            // Save selected schedule name for persistence
            if (selectedItem != null)
            {
                SaveScheduleName(selectedItem.Name);
            }
            
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
                {
                    // Save previous sorting (Deep Copy)
                    var list = new ObservableCollection<SortItem>();
                    foreach(var item in SortCriteria) list.Add(item.Clone());
                    _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
                }

                LoadScheduleData(selectedItem.Schedule);
                
                // Load sorting if exists (Deep Copy Restore)
                SortCriteria.Clear();
                if (_scheduleSortSettings.ContainsKey(selectedItem.Id))
                {
                    foreach(var item in _scheduleSortSettings[selectedItem.Id]) SortCriteria.Add(item.Clone());
                }

                // Restore Itemize Setting
                bool itemize = true;
                if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                {
                    itemize = _scheduleItemizeSettings[selectedItem.Id];
                }
                else
                {
                     _scheduleItemizeSettings[selectedItem.Id] = true; // Default
                }
                
                // Set CheckBox (this triggers checked/unchecked events which call RefreshScheduleView)
                // We need to avoid double refresh if possible, but simplest is to just set it.
                // However, the event handler calls ApplyCurrentSortLogic? No, we added that.
                
                // Temporarily detach events if we want to manually control flow, or just let it trigger.
                // Let's just set IsChecked. The event handler calls RefreshScheduleView(itemize) AND ApplyCurrentSortLogic().
                // This is exactly what we want.
                
                if (ItemizeCheckBox != null)
                {
                    ItemizeCheckBox.IsChecked = itemize;
                }
                
                // If the check state didn't change (already matched), the event won't fire.
                // In that case we must ensure data is loaded.
                // LoadScheduleData does NOT refresh the view passed the initial setup.
                
                // Force apply if event didn't trigger?
                // Actually, LoadScheduleData populates _rawScheduleData.
                // We need to call RefreshScheduleView at least once.
                
                // Let's force it manually if we suspect it might not trigger.
                // But changing source calls LoadScheduleData first. 
                
                // Better approach: 
                // 1. Set check box (might fire event)
                // 2. Ensure ApplyCurrentSortLogic works.
                
                // If I set IsChecked = itemize, and it was already itemize, no event fires.
                // We need to force refresh then.
                
                if (ItemizeCheckBox != null && ItemizeCheckBox.IsChecked == itemize)
                {
                   // Event won't fire, manually refresh
                   RefreshScheduleView(itemize);
                   ApplyCurrentSortLogic();
                }
            }
            else
            {
                // No schedule selected - show empty DataGrid
                SheetsDataGrid.Columns.Clear();
                SheetsDataGrid.ItemsSource = null;
                _currentScheduleData = null;
                AvailableSortColumns.Clear();
            }


        }

        private void RestoreSheetView()
        {
            SheetsDataGrid.ItemsSource = null;
            SheetsDataGrid.Columns.Clear();
            
            SheetsDataGrid.ItemsSource = FilteredSheets;
            SheetsDataGrid.AutoGenerateColumns = false;
            
            InitializeSheetColumns();
        }

        private void InitializeSheetColumns()
        {
            SheetsDataGrid.Columns.Clear();
            
            var checkBoxColumn = CreateCheckBoxColumn();
            SheetsDataGrid.Columns.Add(checkBoxColumn);
            
            var numberCol = new DataGridTextColumn
            {
                Header = "Sheet Number",
                Binding = new System.Windows.Data.Binding("SheetNumber"),
                Width = new DataGridLength(150)
            };
            SheetsDataGrid.Columns.Add(numberCol);
            
            var nameCol = new DataGridTextColumn
            {
                Header = "Sheet Name",
                Binding = new System.Windows.Data.Binding("Name"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            };
            SheetsDataGrid.Columns.Add(nameCol);
        }

        private DataGridTemplateColumn CreateCheckBoxColumn()
        {
            var headerCheckBox = new CheckBox
            {
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Style = (Style)FindResource("CustomCheckBoxStyle")
            };
            headerCheckBox.Checked += HeaderCheckBox_Checked;
            headerCheckBox.Unchecked += HeaderCheckBox_Unchecked;

            var baseStyle = FindResource("ExcelLikeCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center));
            cellStyle.Setters.Add(new Setter(System.Windows.Controls.Control.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center));

            // Create template for checkbox cells
            var cellTemplate = new DataTemplate();
            var checkBoxFactory = new FrameworkElementFactory(typeof(CheckBox));
            checkBoxFactory.SetValue(CheckBox.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            checkBoxFactory.SetValue(CheckBox.StyleProperty, FindResource("CustomCheckBoxStyle"));
            checkBoxFactory.SetBinding(CheckBox.IsCheckedProperty, new System.Windows.Data.Binding("IsSelected") 
            { 
                Mode = System.Windows.Data.BindingMode.TwoWay, 
                UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged 
            });
            checkBoxFactory.AddHandler(CheckBox.ClickEvent, new RoutedEventHandler(RowCheckBox_Click));
            cellTemplate.VisualTree = checkBoxFactory;


            var col = new DataGridTemplateColumn
            {
                Header = headerCheckBox,
                CellTemplate = cellTemplate,
                Width = new DataGridLength(40),
                CellStyle = cellStyle
            };
            return col;
        }

        private void RowCheckBox_Click(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as CheckBox;
            if (checkBox == null) return;
            bool isChecked = checkBox.IsChecked == true;

            // 1. Handle "SelectedItems" (Full Row Selection mode - though we are in Cell mode, this might still be populated if user selects full rows)
            if (SheetsDataGrid.SelectedItems.Count > 1)
            {
               foreach (var selectedItem in SheetsDataGrid.SelectedItems)
                {
                    SetRowSelection(selectedItem, isChecked);
                }
            }

            // 2. Handle "SelectedCells" (Cell Selection mode)
            // If the user selected a range of cells in the first column (Checkbox column), we want to toggle all of them.
            // Checkbox column is usually index 0.
            if (SheetsDataGrid.SelectedCells.Count > 1)
            {
                foreach(var cellInfo in SheetsDataGrid.SelectedCells)
                {
                    // Check if this cell is in the Checkbox column
                    // We can check Column.DisplayIndex or checking the content. 
                    // Our Checkbox column is a DataGridTemplateColumn created in code.
                    // Let's assume it's the one with index 0.
                    if (cellInfo.Column.DisplayIndex == 0)
                    {
                        SetRowSelection(cellInfo.Item, isChecked);
                    }
                }
            }
        }

        private void SetRowSelection(object item, bool isSelected)
        {
            if (item is System.Data.DataRowView rowView)
            {
                rowView["IsSelected"] = isSelected;
            }
            else if (item is SheetItem sheet)
            {
                sheet.IsSelected = isSelected;
            }
        }

        private void HeaderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var view = SheetsDataGrid.ItemsSource;
            if (view is System.Data.DataView dataView)
            {
                foreach (System.Data.DataRowView row in dataView)
                {
                    row["IsSelected"] = true;
                }
            }
            else if (view is ObservableCollection<SheetItem> sheets)
            {
                foreach (var sheet in sheets)
                {
                    sheet.IsSelected = true;
                }
            }
        }

        private void HeaderCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            var view = SheetsDataGrid.ItemsSource;
            if (view is System.Data.DataView dataView)
            {
                foreach (System.Data.DataRowView row in dataView)
                {
                    row["IsSelected"] = false;
                }
            }
            else if (view is ObservableCollection<SheetItem> sheets)
            {
                foreach (var sheet in sheets)
                {
                    sheet.IsSelected = false;
                }
            }
        }



        private void LoadScheduleData(ViewSchedule schedule)
        {
            try
            {
                var data = _revitService.GetScheduleData(schedule);
                _currentScheduleData = data;
                
                if (data == null) throw new Exception("Failed to retrieve schedule data from Revit.");

                if (!_scheduleItemizeSettings.ContainsKey(schedule.Id))
                {
                    _scheduleItemizeSettings[schedule.Id] = true;
                }
                bool isItemized = _scheduleItemizeSettings[schedule.Id];
                
                // Removed UI checkbox update from here to prevent premature event firing.
                // It is now handled in SelectionChanged after SortCriteria is restored.
                
                var dt = new System.Data.DataTable();
                dt.Columns.Add("IsSelected", typeof(bool)).DefaultValue = false;
                dt.Columns.Add("RowState", typeof(string)).DefaultValue = "Unchanged";
                dt.Columns.Add("Count", typeof(int));

                var newSortColumns = new List<string> { "(none)" };

                // Detect column types
                for(int i = 0; i < data.Columns.Count; i++)
                {
                    string safeName = data.Columns[i];
                    int dupIdx = 1;
                    while(dt.Columns.Contains(safeName))
                    {
                        safeName = $"{data.Columns[i]} ({dupIdx++})";
                    }

                    // Filter out internal columns from sorting options
                    if (safeName != "ElementId" && safeName != "TypeName" && !safeName.StartsWith("Count"))
                    {
                        newSortColumns.Add(safeName);
                    }


                    // Check if column is numeric
                    bool isNumeric = true;
                    bool hasValue = false;

                    if (safeName == "ElementId")
                    {
                        isNumeric = false;
                        hasValue = true; // Force string creation
                    }
                    else
                    {
                        foreach(var r in data.Rows)
                        {
                            string val = r[i];
                            if (string.IsNullOrWhiteSpace(val)) continue;
                            hasValue = true;
                            if (!double.TryParse(val, out _))
                            {
                                isNumeric = false;
                                break;
                            }
                        }
                    }

                    if (isNumeric && hasValue)
                    {
                        dt.Columns.Add(safeName, typeof(double));
                    }
                    else
                    {
                        dt.Columns.Add(safeName, typeof(string));
                    }
                }

                // Smart Update AvailableSortColumns to preserve bindings
                bool updateNeeded = AvailableSortColumns.Count != newSortColumns.Count;
                if (!updateNeeded)
                {
                    for (int i = 0; i < newSortColumns.Count; i++)
                    {
                        if (AvailableSortColumns[i] != newSortColumns[i])
                        {
                            updateNeeded = true;
                            break;
                        }
                    }
                }

                if (updateNeeded)
                {
                    AvailableSortColumns.Clear();
                    foreach (var col in newSortColumns)
                    {
                        AvailableSortColumns.Add(col);
                    }
                }
                
                foreach(var row in data.Rows)
                {
                    var newRow = dt.NewRow();
                    newRow["IsSelected"] = false;
                    newRow["RowState"] = "Unchanged";
                    newRow["Count"] = 1;
                    
                    for(int i = 0; i < data.Columns.Count; i++)
                    {
                        string val = row[i];
                        // If column is numeric, parse it
                        if (dt.Columns[i + 3].DataType == typeof(double))
                        {
                            if (double.TryParse(val, out double dVal))
                            {
                                newRow[i + 3] = dVal;
                            }
                            else
                            {
                                newRow[i + 3] = DBNull.Value;
                            }
                        }
                        else
                        {
                            newRow[i + 3] = val;
                        }
                    }
                    
                    dt.Rows.Add(newRow);
                }
                
                _rawScheduleData = dt;
                // RefreshScheduleView(isItemized); <-- Removed to prevent refresh with stale SortCriteria
            }
            catch (Exception ex)
            {
                ShowPopup("Error Loading Schedule", ex.Message);
            }
        }

        private void RefreshScheduleView(bool itemize)
        {
            System.Data.DataTable viewTable = _rawScheduleData;
            
            if (!itemize && viewTable != null)
            {
                viewTable = viewTable.Clone();
                // Group by sorting rules
                var validSorts = SortCriteria.Where(s => s.SelectedColumn != "(none)" && !string.IsNullOrEmpty(s.SelectedColumn))
                                             .Select(s => s.SelectedColumn)
                                             .ToList();

                var grouped = _rawScheduleData.AsEnumerable()
                    .GroupBy(r => 
                    {
                        if (validSorts.Count == 0) return ""; // Group all if no sort
                        
                        // Create composite key
                        return string.Join("||", validSorts.Select(col => 
                            r.Table.Columns.Contains(col) ? r[col]?.ToString() ?? "" : ""));
                    });
                
                foreach(var grp in grouped)
                {
                    var firstRow = grp.First();
                    var newRow = viewTable.NewRow();
                    newRow.ItemArray = firstRow.ItemArray;
                    newRow["Count"] = grp.Count();

                    // Aggregate ElementIds if present
                    if (viewTable.Columns.Contains("ElementId"))
                    {
                        var ids = grp.Select(r => r["ElementId"]?.ToString())
                                     .Where(s => !string.IsNullOrEmpty(s));
                        newRow["ElementId"] = string.Join(",", ids);
                    }

                    viewTable.Rows.Add(newRow);
                }
            }
            
            SheetsDataGrid.ItemsSource = null;
            SheetsDataGrid.Columns.Clear();
            SheetsDataGrid.AutoGenerateColumns = false;
            
            if (viewTable == null) return;
            
            var baseStyle = FindResource("ExcelLikeCellStyle") as Style;
            var cellStyle = new Style(typeof(DataGridCell), baseStyle);
            var cellTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("RowState"),
                Value = "Pending"
            };
            cellTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(Colors.Yellow) { Opacity = 0.5 }));
            cellStyle.Triggers.Add(cellTrigger);
            
            // First add checkbox column
            var checkCol = CreateCheckBoxColumn();
            SheetsDataGrid.Columns.Add(checkCol);
            
            // Then add schedule data columns (skip RowState, ElementId, Count, IsSelected)
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "Count", "TypeName" };
            foreach(System.Data.DataColumn col in viewTable.Columns)
            {
                if (skipColumns.Contains(col.ColumnName)) continue;
                
                // Skip columns that end with " (1)", " (2)", etc. - these are duplicates
                if (System.Text.RegularExpressions.Regex.IsMatch(col.ColumnName, @"\s\(\d+\)$")) continue;
                
                var textCol = new DataGridTextColumn
                {
                    Header = col.ColumnName,
                    Binding = new System.Windows.Data.Binding(col.ColumnName),
                    CellStyle = cellStyle,
                    EditingElementStyle = (Style)FindResource("DataGridEditingStyle"),
                    IsReadOnly = false
                };
                SheetsDataGrid.Columns.Add(textCol);
            }
            
            // Finally add Count column at the end
            if (viewTable.Columns.Contains("Count"))
            {
                var countCol = new DataGridTextColumn
                {
                    Header = "Count",
                    Binding = new System.Windows.Data.Binding("Count"),
                    CellStyle = cellStyle,
                    IsReadOnly = true
                };
                SheetsDataGrid.Columns.Add(countCol);
            }
            
            SheetsDataGrid.ItemsSource = new System.Data.DataView(viewTable);
            
            // Subscribe to cell editing event (unsubscribe first to avoid duplicates)
            SheetsDataGrid.CellEditEnding -= SheetsDataGrid_CellEditEnding;
            SheetsDataGrid.CellEditEnding += SheetsDataGrid_CellEditEnding;
        }

        private void Itemize_Checked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = true;
                RefreshScheduleView(true);
                ApplyCurrentSortLogic();
            }
        }

        private void Itemize_Unchecked(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem != null && selectedItem.Schedule != null)
            {
                _scheduleItemizeSettings[selectedItem.Id] = false;
                // Calling ApplyCurrentSortLogic will see the 'false' setting and trigger RefreshScheduleView(false) internally
                // Then it will proceed to apply SortDescriptions.
                ApplyCurrentSortLogic();
            }
        }

        private async void SheetsDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            
            // Get the edited cell data
            var row = e.Row.Item as System.Data.DataRowView;
            if (row == null || _currentScheduleData == null) return;

            string columnName = e.Column.Header?.ToString();
            if (string.IsNullOrEmpty(columnName)) return;

            // Skip non-parameter columns
            var skipColumns = new[] { "IsSelected", "RowState", "ElementId", "TypeName", "Count" };
            if (skipColumns.Contains(columnName)) return;

            // Get the new value from the editing element
            var editingElement = e.EditingElement as System.Windows.Controls.TextBox;
            if (editingElement == null) return;

            string newValue = editingElement.Text;
            string oldValue = row[columnName]?.ToString() ?? "";

            // If value hasn't changed, skip
            if (newValue == oldValue) return;

            // Get element ID from the row
            if (!row.Row.Table.Columns.Contains("ElementId")) return;
            string elementIdStr = row["ElementId"]?.ToString();
            if (string.IsNullOrEmpty(elementIdStr)) return;

            // Find the parameter ID for this column
            int columnIndex = _currentScheduleData.Columns.IndexOf(columnName);
            if (columnIndex < 0 || columnIndex >= _currentScheduleData.ParameterIds.Count) return;

            ElementId parameterId = _currentScheduleData.ParameterIds[columnIndex];
            
            // Check if it's a type parameter
            bool isTypeParameter = _currentScheduleData.IsTypeParameter.ContainsKey(columnName) && 
                                   _currentScheduleData.IsTypeParameter[columnName];

            // Show confirmation for type parameters
            if (isTypeParameter)
            {
                // Store values for the confirmation callback
                var tempElementIdStr = elementIdStr;
                var tempParameterId = parameterId;
                var tempNewValue = newValue;
                var tempOldValue = oldValue;
                var tempRow = row;
                var tempColumnName = columnName;

                ShowConfirmationPopup(
                    "Type Parameter Warning",
                    "This is a TYPE parameter. Changing it will affect ALL elements of this type.\n\nDo you want to proceed?",
                    () => PerformParameterUpdate(tempElementIdStr, tempParameterId, tempNewValue, tempOldValue, tempRow, tempColumnName),
                    () => 
                    {
                        // Cancelled - revert the value in the UI
                        // Since CellEditEnding happens after commit, the DataTable has the new value.
                        // We need to set it back to the old value.
                        tempRow.Row[tempColumnName] = tempOldValue;
                    });
                return;
            }

            // For instance parameters, update immediately
            PerformParameterUpdate(elementIdStr, parameterId, newValue, oldValue, row, columnName);
        }

        private void PerformParameterUpdate(string elementIdStr, ElementId parameterId, string newValue, 
                                                  string oldValue, System.Data.DataRowView row, string columnName)
        {
            // Update the parameter value via external event
            _parameterValueUpdateHandler.IsBatchMode = false;
            _parameterValueUpdateHandler.ElementIdStr = elementIdStr;
            _parameterValueUpdateHandler.ParameterIdStr = parameterId.Value.ToString();
            _parameterValueUpdateHandler.NewValue = newValue;

            _parameterValueUpdateExternalEvent.Raise();
            
            // UI Update is now handled by OnParameterValueUpdateFinished
        }

        private void OnDuplicationFinished(int success, int fail, string errorMsg, List<ElementId> newSheetIds)
        {
            Dispatcher.Invoke(() =>
            {
                // Update pending creation sheets with new IDs
                if (newSheetIds != null && newSheetIds.Count > 0)
                {
                    var pendingCreations = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                    for (int i = 0; i < Math.Min(pendingCreations.Count, newSheetIds.Count); i++)
                    {
                        pendingCreations[i].Id = newSheetIds[i];
                        pendingCreations[i].State = SheetItemState.ExistingInRevit;
                        pendingCreations[i].OriginalSheetNumber = pendingCreations[i].SheetNumber;
                        pendingCreations[i].OriginalName = pendingCreations[i].Name;
                    }
                }

                // Reload data to show all sheets
                if (_uiApplication != null)
                {
                    LoadData(_uiApplication.ActiveUIDocument.Document);
                }

                UpdateButtonStates();

                // Show Result Popup
                if (fail > 0)
                {
                    ShowPopup("Duplication Report", $"Success: {success}\nFailures: {fail}\nLast Error: {errorMsg}");
                }
                else
                {
                    ShowPopup("Success", $"Successfully created {success} sheet(s).");
                }
            });
        }

        #endregion

        #region Actions

        private void AddDuplicates_Click(object sender, RoutedEventArgs e)
        {
            // Validate selection
            var selectedCount = Sheets.Count(x => x.IsSelected);
            if (selectedCount == 0)
            {
                ShowPopup("No Sheets Selected", "Please select at least one sheet to duplicate.");
                return;
            }

            OptCopies.Text = "1";
            OptKeepViews.IsChecked = false;
            OptKeepLegends.IsChecked = false;
            OptKeepSchedules.IsChecked = false;
            OptCopyRevisions.IsChecked = true;
            OptCopyParams.IsChecked = true;

            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void Popup_MouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true; // Prevent click from bubbling to background
        }

        private void OptionsPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Close on background click? User choice, but safe to allow cancelling
            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void DuplicateCancel_Click(object sender, RoutedEventArgs e)
        {
            OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void DuplicateConfirm_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Validate and parse options
                if (!int.TryParse(OptCopies.Text, out int copies) || copies < 1)
                {
                    ShowPopup("Invalid Input", "Please enter a valid number of copies (1 or more).");
                    return;
                }

                // 2. Build DuplicateOptions object
                var options = new DuplicateOptions
                {
                    NumberOfCopies = copies,
                    DuplicateMode = OptKeepViews.IsChecked == true
                        ? ExternalEvents.SheetDuplicateMode.WithViews
                        : ExternalEvents.SheetDuplicateMode.WithSheetDetailing,
                    KeepLegends = OptKeepLegends.IsChecked == true,
                    KeepSchedules = OptKeepSchedules.IsChecked == true,
                    CopyRevisions = OptCopyRevisions.IsChecked == true,
                    CopyParameters = OptCopyParams.IsChecked == true
                };

                // 3. Create pending SheetItems
                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                foreach (var sourceSheet in selectedSheets)
                {
                    for (int i = 0; i < copies; i++)
                    {
                        var pendingSheet = new SheetItem(
                            sheetNumber: $"{sourceSheet.SheetNumber} - Copy {i + 1}",
                            name: $"{sourceSheet.Name} - Copy {i + 1}",
                            sourceSheetId: sourceSheet.Id,
                            options: options
                        );

                        // Subscribe to property changes
                        pendingSheet.PropertyChanged += OnSheetPropertyChanged;

                        // Add to collections
                        Sheets.Add(pendingSheet);

                        // Check if it matches current search filter
                        var searchText = SheetSearchBox?.Text?.ToLowerInvariant() ?? "";
                        if (string.IsNullOrEmpty(searchText) ||
                            pendingSheet.Name.ToLowerInvariant().Contains(searchText) ||
                            pendingSheet.SheetNumber.ToLowerInvariant().Contains(searchText))
                        {
                            FilteredSheets.Add(pendingSheet);
                        }
                    }
                }

                // 4. Update button states
                UpdateButtonStates();

                // 5. Close Popup
                OptionsPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private SortingWindow _sortingWindow;

        private void Sort_Click(object sender, RoutedEventArgs e)
        {


            if (_sortingWindow == null || !_sortingWindow.IsLoaded)
            {
                _sortingWindow = new SortingWindow(this);
                _sortingWindow.Owner = this;
                _sortingWindow.Show();
            }
            else
            {
                _sortingWindow.Activate();
                if (_sortingWindow.WindowState == WindowState.Minimized)
                    _sortingWindow.WindowState = WindowState.Normal;
            }
        }



        internal void ApplyCurrentSortLogicInternal()
        {
            ApplyCurrentSortLogic();
        }

        private void ApplyCurrentSortLogic()
        {
            if (SheetsDataGrid.ItemsSource == null) return;

            // Check if we are in non-itemized mode (grouped)
            // If so, we must REFRESH the view to regroup based on new sort criteria
            if (_currentScheduleData != null && 
                _scheduleItemizeSettings.ContainsKey(_currentScheduleData.ScheduleId) && 
                _scheduleItemizeSettings[_currentScheduleData.ScheduleId] == false)
            {
                RefreshScheduleView(false);
            }
            
            System.ComponentModel.ICollectionView view = System.Windows.Data.CollectionViewSource.GetDefaultView(SheetsDataGrid.ItemsSource);
            view.SortDescriptions.Clear();

            foreach (var sortItem in SortCriteria)
            {
                if (string.IsNullOrEmpty(sortItem.SelectedColumn) || sortItem.SelectedColumn == "(none)") continue;
                
                // Map display name to binding path if necessary
                string propertyName = sortItem.SelectedColumn;
                
                // If using DataTable, property name is column name
                // If using SheetItem, property maps: 
                if (propertyName == "Sheet Number") propertyName = "SheetNumber"; // SheetItem property
                else if (propertyName == "Sheet Name") propertyName = "Name"; // SheetItem property
                
                view.SortDescriptions.Add(new System.ComponentModel.SortDescription(
                    propertyName, 
                    sortItem.IsAscending ? System.ComponentModel.ListSortDirection.Ascending : System.ComponentModel.ListSortDirection.Descending
                ));
            }
        }

/*
        private void SortCancel_Click(object sender, RoutedEventArgs e)
        {
            // Revert changes (if any were made live without backup? No, we bound directly)
            // Wait, if we bound directly, changes are LIVE.
            // We need to restore backup.
            
            // But we didn't store backup in a field? We did in Sort_Click but it was local.
            // Logic was flawed or reliant on local variable capture? No, WPF ensures modal? No, it was a Popup.
            // Actually the original Sort_Click logic was weird, it cleared then added.
            
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }
*/

        #region Persistence

        private void SaveSortSettings()
        {
            try
            {
                // 1. Commit current schedule settings to dictionary before saving
                if (_currentScheduleData != null && _currentScheduleData.ScheduleId != ElementId.InvalidElementId)
                {
                    var list = new ObservableCollection<SortItem>();
                    foreach(var item in SortCriteria) list.Add(item.Clone());
                    _scheduleSortSettings[_currentScheduleData.ScheduleId] = list;
                }

                var dtos = new List<SavedScheduleSort>();
                foreach(var kvp in _scheduleSortSettings)
                {
                    // Use robust ID extraction (Value or IntegerValue)
                    long idVal = GetIdValue(kvp.Key);
                    
                    bool itemize = true;
                    if (_scheduleItemizeSettings.ContainsKey(kvp.Key)) itemize = _scheduleItemizeSettings[kvp.Key];
                    
                    dtos.Add(new SavedScheduleSort 
                    { 
                        ScheduleId = idVal, 
                        Items = kvp.Value.ToList(),
                        ItemizeEveryInstance = itemize
                    });
                }

                string folder = GetProjectSettingsFolder();
                // if (!System.IO.Directory.Exists(folder)) System.IO.Directory.CreateDirectory(folder); // Helper does it

                string file = System.IO.Path.Combine(folder, "sort_settings.json");
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(dtos, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(file, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private void LoadSortSettings()
        {
            try
            {
                string folder = GetProjectSettingsFolder();
                string file = System.IO.Path.Combine(folder, "sort_settings.json");
                if (System.IO.File.Exists(file))
                {
                    string json = System.IO.File.ReadAllText(file);
                    var dtos = Newtonsoft.Json.JsonConvert.DeserializeObject<List<SavedScheduleSort>>(json);
                    
                    if (dtos != null)
                    {
                        _scheduleSortSettings.Clear();
                        
                        foreach(var dto in dtos)
                        {
                            // Reconstruct ElementId
                            // Using the constructor available in older/newer API via helper or #if logic?
                            // Just use reflection or try generic constructor if possible.
                            // Actually, ElementId(long) exists in 2024+. ElementId(int) in older.
                            // Since we target multiple frameworks, let's try to map safely?
                            // Or just use the constructor that accepts long?
                            // Warnings earlier said ElementId(int) is deprecated, use ElementId(long).
                            
                            ElementId eid = new ElementId((long)dto.ScheduleId);
                            
                            _scheduleSortSettings[eid] = new ObservableCollection<SortItem>(dto.Items);
                            
                            // Load itemize setting (default true)
                            _scheduleItemizeSettings[eid] = dto.ItemizeEveryInstance;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }
        }

        private long GetIdValue(ElementId id)
        {
            return id.Value;
        }
        
        public class SavedScheduleSort
        {
            public long ScheduleId { get; set; }
            public List<SortItem> Items { get; set; }
            public bool ItemizeEveryInstance { get; set; } = true;
        }
        
        #endregion

        private void AddSortLevel_Click(object sender, RoutedEventArgs e)
        {
            SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
        }


/*
        private bool _isSortDragging = false;
        private System.Windows.Point _sortDragStart;

        private void SortHeader_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                var element = sender as IInputElement;
                _isSortDragging = true;
                _sortDragStart = e.GetPosition(this);
                element.CaptureMouse();
                e.Handled = true;
            }
        }
        
        private void SortPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            SortPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void SortHeader_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isSortDragging && sender is IInputElement element)
            {
                var current = e.GetPosition(this);
                var diff = current - _sortDragStart;
                
                if (SortPopupTransform != null)
                {
                    SortPopupTransform.X += diff.X;
                    SortPopupTransform.Y += diff.Y;
                }
                
                _sortDragStart = current;
            }
        }

        private void SortHeader_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSortDragging)
            {
                _isSortDragging = false;
                (sender as IInputElement)?.ReleaseMouseCapture();
            }
        }
*/



        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Allow clicks on interactive controls or DataGrid rows to function normally
            DependencyObject obj = e.OriginalSource as DependencyObject;
            while (obj != null)
            {
                // If we clicked a control that should handle its own events, or the DataGridRow itself, don't deselect
                if (obj is System.Windows.Controls.Button || 
                    obj is System.Windows.Controls.TextBox || 
                    obj is System.Windows.Controls.CheckBox || 
                    obj is System.Windows.Controls.ComboBox || 
                    obj is System.Windows.Controls.Primitives.ScrollBar || 
                    obj is System.Windows.Controls.Primitives.DataGridColumnHeader || 
                    obj is System.Windows.Controls.DataGridRow || 
                    obj is System.Windows.Controls.Primitives.Thumb)
                {
                    return;
                }

                if (obj is System.Windows.ContentElement contentElement)
                {
                    obj = System.Windows.LogicalTreeHelper.GetParent(contentElement);
                    continue; // Skip rest of loop for this iteration
                }

                if (obj is System.Windows.Media.Visual || obj is System.Windows.Media.Media3D.Visual3D)
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }
                else
                {
                    obj = null; // Stop traversal if not a visual
                }
            }

            // If we are here, we clicked empty space (background, borders, etc.) -> Deselect All
            if (SheetsDataGrid != null)
            {
                SheetsDataGrid.UnselectAll();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Validate no sheet number conflicts
                if (!ValidateAllSheets())
                {
                    return; // Error message already shown
                }

                // 2. Separate pending creations from pending edits
                var sheetsToCreate = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                var sheetsToEdit = Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList();

                bool hasWork = sheetsToCreate.Count > 0 || sheetsToEdit.Count > 0;
                if (!hasWork)
                {
                    return;
                }

                // 3. Trigger creation handler if needed
                if (sheetsToCreate.Count > 0)
                {
                    _handler.PendingSheetData = sheetsToCreate;
                    _externalEvent.Raise();
                }

                // 4. Trigger edit handler if needed
                if (sheetsToEdit.Count > 0)
                {
                    _editHandler.SheetsToEdit = sheetsToEdit;
                    _editExternalEvent.Raise();
                }

                // Buttons removed - changes are applied immediately
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void DiscardPending_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var pendingCount = Sheets.Count(s => s.HasUnsavedChanges);
                if (pendingCount == 0) return;

                ShowConfirmPopup("Confirm Discard", $"Discard {pendingCount} pending change(s)?", () =>
                {
                    // Remove pending creations
                    var toRemove = Sheets.Where(s => s.State == SheetItemState.PendingCreation).ToList();
                    foreach (var sheet in toRemove)
                    {
                        sheet.PropertyChanged -= OnSheetPropertyChanged; // Unsubscribe
                        Sheets.Remove(sheet);
                        FilteredSheets.Remove(sheet);
                    }

                    // Revert pending edits
                    foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit).ToList())
                    {
                        // Reset properties to original? 
                        // Simplified: Just clear dirty flag if that was the only change
                        // Real revert needs original values which we might not have stored easily here
                        // For now, just reset the state
                        sheet.State = SheetItemState.ExistingInRevit;
                    }

                    UpdateButtonStates();
                    ShowPopup("Success", "Discarded all pending changes.");
                });
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void HighlightInModel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<ElementId> elementIds = new List<ElementId>();

                // Check referencing ItemsSource to determine mode
                if (SheetsDataGrid.ItemsSource is System.Data.DataView dataView)
                {
                    // Schedule Mode (DataTable)
                    foreach (System.Data.DataRowView row in dataView)
                    {
                        // Check IsSelected column
                        if (row.Row.Table.Columns.Contains("IsSelected") && 
                            row["IsSelected"] != DBNull.Value && 
                            Convert.ToBoolean(row["IsSelected"]))
                        {
                            if (row.Row.Table.Columns.Contains("ElementId"))
                            {
                                string idStr = row["ElementId"]?.ToString();
                                
                                // Handle grouped rows (comma-separated IDs)
                                if (!string.IsNullOrEmpty(idStr))
                                {
                                    var parts = idStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach(var part in parts)
                                    {
                                        if (long.TryParse(part.Trim(), out long idLong))
                                        {
#if NET8_0_OR_GREATER
                                            elementIds.Add(new ElementId(idLong));
#else
                                            elementIds.Add(new ElementId(idLong));
#endif
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Sheet Mode (SheetItem collection)
                    // Use ticked sheets (Checkbox = IsSelected)
                    elementIds = Sheets
                        .Where(s => s.IsSelected)
                        .Select(s => s.Id)
                        .ToList();
                }

                if (elementIds.Count == 0)
                {
                    ShowPopup("Selection Required", "Please tick at least one item to highlight.");
                    return;
                }

                // Filter valid IDs
                var validIds = elementIds
                    .Where(id => id != null && id != ElementId.InvalidElementId)
                    .ToList();

                if (validIds.Count == 0)
                {
                    ShowPopup("Error", "Selected items do not have valid Revit Element IDs.");
                    return;
                }

                // Raise External Event
                _highlightInModelHandler.ElementIds = validIds;
                _highlightInModelExternalEvent.Raise();
            }
            catch (Exception ex)
            {
                ShowPopup("Error", $"Failed to highlight elements: {ex.Message}");
            }
        }

        private void OnSheetPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (sender is SheetItem sheet && (e.PropertyName == "Name" || e.PropertyName == "SheetNumber"))
            {
                // Mark as edited if it's an existing sheet
                if (sheet.State == SheetItemState.ExistingInRevit &&
                    (sheet.SheetNumber != sheet.OriginalSheetNumber || sheet.Name != sheet.OriginalName))
                {
                    sheet.State = SheetItemState.PendingEdit;
                }

                ValidateSheetNumber(sheet);
                UpdateButtonStates();
            }
        }

        private void OnEditFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Update states to ExistingInRevit
                foreach (var sheet in Sheets.Where(s => s.State == SheetItemState.PendingEdit))
                {
                    sheet.State = SheetItemState.ExistingInRevit;
                    sheet.OriginalSheetNumber = sheet.SheetNumber;
                    sheet.OriginalName = sheet.Name;
                }

                UpdateButtonStates();

                // Show results
                if (fail > 0)
                {
                    ShowPopup("Edit Report", $"Success: {success}\nFailures: {fail}\nError: {errorMsg}");
                }
                else
                {
                    ShowPopup("Success", $"Successfully updated {success} sheet(s).");
                }
            });
        }

        private void ValidateSheetNumber(SheetItem sheet)
        {
            var duplicates = Sheets.Where(s =>
                s != sheet &&
                s.SheetNumber == sheet.SheetNumber &&
                (s.State == SheetItemState.ExistingInRevit || s.State == SheetItemState.PendingEdit)
            ).ToList();

            sheet.HasNumberConflict = duplicates.Any();
            sheet.ValidationError = duplicates.Any()
                ? $"Sheet number '{sheet.SheetNumber}' already exists"
                : null;
        }

        private bool ValidateAllSheets()
        {
            bool hasErrors = false;

            foreach (var sheet in Sheets.Where(s => s.HasUnsavedChanges))
            {
                ValidateSheetNumber(sheet);
                if (sheet.HasNumberConflict)
                {
                    hasErrors = true;
                }
            }

            if (hasErrors)
            {
                ShowPopup("Validation Error", "Please fix duplicate sheet numbers before applying.");
                return false;
            }

            return true;
        }

        private void UpdateButtonStates()
        {
            // Buttons removed - this method is kept in case it's referenced elsewhere
        }

        private RenameWindow _renameWindow;

        private void Rename_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Determine if we're working with schedule data or sheets
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                bool isScheduleMode = selectedItem != null && selectedItem.Schedule != null && _currentScheduleData != null;

                if (isScheduleMode)
                {
                    // Get selected rows from schedule DataView
                    var view = SheetsDataGrid.ItemsSource as System.Data.DataView;
                    if (view == null)
                    {
                        ShowPopup("Error", "No schedule data available.");
                        return;
                    }

                    var selectedRows = new List<System.Data.DataRowView>();
                    foreach (System.Data.DataRowView row in view)
                    {
                        var isSelectedValue = row["IsSelected"];
                        bool isSelected = false;
                        
                        if (isSelectedValue is bool b)
                        {
                            isSelected = b;
                        }
                        else if (isSelectedValue != null && isSelectedValue != DBNull.Value)
                        {
                            isSelected = Convert.ToBoolean(isSelectedValue);
                        }
                        
                        if (isSelected)
                        {
                            selectedRows.Add(row);
                        }
                    }

                    if (selectedRows.Count == 0)
                    {
                        ShowPopup("No Rows Selected", "Please tick at least one row to rename.");
                        return;
                    }

                    // Open RenameWindow in schedule mode
                    _renameWindow = new RenameWindow(this, selectedRows, _currentScheduleData);
                    _renameWindow.OnScheduleRenameApply += OnScheduleRenameApply;
                    _renameWindow.ShowDialog();
                }
                else
                {
                    // Sheet mode
                    var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                    if (selectedSheets.Count == 0)
                    {
                        ShowPopup("No Sheets Selected", "Please select at least one sheet to rename.");
                        return;
                    }

                    // Open RenameWindow in sheet mode
                    _renameWindow = new RenameWindow(this, selectedSheets);
                    _renameWindow.OnSheetRenameApply += OnSheetRenameApply;
                    _renameWindow.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }

        private void OnScheduleRenameApply(List<ScheduleRenameItem> items)
        {
            _parameterRenameHandler.RenameItems = items;
            _parameterRenameExternalEvent.Raise();
        }

        private void OnSheetRenameApply(List<RenamePreviewItem> items, string parameterName)
        {
            bool isSheetNumber = parameterName == "Sheet Number";

            foreach (var item in items)
            {
                if (isSheetNumber)
                {
                    item.Sheet.SheetNumber = item.New;
                }
                else
                {
                    item.Sheet.Name = item.New;
                }
            }
        }



        private void RenameParameter_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateRenamePreview();
        }

        private void RenameText_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateRenamePreview();
        }

        private void UpdateRenamePreview()
        {
            try
            {
                RenamePreviewItems.Clear();

                string findText = RenameFindText?.Text ?? "";
                string replaceText = RenameReplaceText?.Text ?? "";
                string prefix = RenamePrefixText?.Text ?? "";
                string suffix = RenameSuffixText?.Text ?? "";

                var selectedOption = RenameParameter?.SelectedItem as RenameParameterOption;
                if (selectedOption == null) return;

                // This popup is only used for sheet mode now (schedule mode uses RenameWindow)
                var selectedSheets = Sheets.Where(x => x.IsSelected).ToList();
                if (selectedSheets.Count == 0) return;

                bool isSheetNumber = selectedOption.Name == "Sheet Number";

                foreach (var sheet in selectedSheets)
                {
                    string original = isSheetNumber ? sheet.SheetNumber : sheet.Name;
                    string newValue = ApplyRenameTransform(original, findText, replaceText, prefix, suffix);

                    var previewItem = new RenamePreviewItem(sheet, original)
                    {
                        New = newValue
                    };

                    RenamePreviewItems.Add(previewItem);
                }

                RenamePreviewDataGrid.ItemsSource = RenamePreviewItems;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating rename preview: {ex.Message}");
            }
        }


        /// <summary>
        /// Applies find/replace, prefix, and suffix transformations to a value.
        /// </summary>
        private string ApplyRenameTransform(string original, string find, string replace, string prefix, string suffix)
        {
            string newValue = original ?? "";

            // Apply find/replace
            if (!string.IsNullOrEmpty(find))
            {
                newValue = newValue.Replace(find, replace);
            }

            // Apply prefix
            if (!string.IsNullOrEmpty(prefix))
            {
                newValue = prefix + newValue;
            }

            // Apply suffix
            if (!string.IsNullOrEmpty(suffix))
            {
                newValue = newValue + suffix;
            }

            return newValue;
        }


        private void RenameApply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedOption = RenameParameter?.SelectedItem as RenameParameterOption;

                // This popup is only used for sheet mode now (schedule mode uses RenameWindow)
                bool isSheetNumber = selectedOption?.Name == "Sheet Number";

                foreach (var previewItem in RenamePreviewItems)
                {
                    if (previewItem.Original != previewItem.New)
                    {
                        if (isSheetNumber)
                        {
                            previewItem.Sheet.SheetNumber = previewItem.New;
                        }
                        else
                        {
                            previewItem.Sheet.Name = previewItem.New;
                        }
                    }
                }

                // Close popup
                RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ShowPopup("Error", ex.Message);
            }
        }


        /// <summary>
        /// Executes the schedule parameter rename via ExternalEvent.
        /// </summary>
        private void ExecuteScheduleRename(List<ScheduleRenameItem> items)
        {
            _parameterRenameHandler.RenameItems = items;
            _parameterRenameExternalEvent.Raise();
            RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        /// <summary>
        /// Callback when the parameter rename operation completes.
        /// </summary>
        private void OnParameterRenameFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Reload schedule data to show updated values
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                if (selectedItem?.Schedule != null)
                {
                    LoadScheduleData(selectedItem.Schedule);

                    // Restore itemize setting and refresh view
                    bool itemize = true;
                    if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                    {
                        itemize = _scheduleItemizeSettings[selectedItem.Id];
                    }
                    RefreshScheduleView(itemize);
                    ApplyCurrentSortLogic();
                }

                // Show result
                if (fail > 0)
                {
                    ShowPopup("Rename Report", $"Success: {success}\nFailures: {fail}\nError: {errorMsg}");
                }
                else if (success > 0)
                {
                    ShowPopup("Success", $"Successfully renamed {success} parameter value(s).");
                }
            });
        }

        private void OnParameterValueUpdateFinished(int success, int fail, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                // Reload schedule data to show updated values from Revit
                var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                if (selectedItem?.Schedule != null)
                {
                    LoadScheduleData(selectedItem.Schedule);

                    // Restore itemize setting and refresh view
                    bool itemize = true;
                    if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                    {
                        itemize = _scheduleItemizeSettings[selectedItem.Id];
                    }
                    RefreshScheduleView(itemize);
                    ApplyCurrentSortLogic();
                }

                // Show error if failed
                if (fail > 0 || !string.IsNullOrEmpty(errorMsg))
                {
                    string msg = string.IsNullOrEmpty(errorMsg) ? "Operation failed." : errorMsg;
                    if (fail > 0) msg += $"\nFailures: {fail}";
                    if (success > 0) msg += $"\nSuccess: {success}";
                    
                    ShowPopup("Update Report", msg);
                }
                // No popup for pure success to avoid annoying the user on every cell edit
            });
        }


        private void RenameCancel_Click(object sender, RoutedEventArgs e)
        {
            RenamePopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void Parameters_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem == null || selectedItem.Schedule == null)
            {
                ShowPopup("No Schedule Selected", "Please select a schedule first.");
                return;
            }

            _parameterLoadHandler.ScheduleId = selectedItem.Id;
            _parameterLoadExternalEvent.Raise();
        }

        private void OnParameterDataLoaded(List<ParameterItem> available, List<ParameterItem> scheduled, string categoryName)
        {
            Dispatcher.Invoke(() =>
            {
                var win = new ParametersWindow(available, scheduled, categoryName);
                win.Owner = this;
                win.OnApply += OnParametersApply;
                win.ShowDialog();
            });
        }

        private void OnParametersApply(List<ParameterItem> newFields)
        {
            var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
            if (selectedItem == null) return;

            _scheduleFieldsHandler.ScheduleId = selectedItem.Id;
            _scheduleFieldsHandler.NewFields = newFields;
            _scheduleFieldsExternalEvent.Raise();
        }

        private void OnScheduleFieldsUpdateFinished(int count, string errorMsg)
        {
            Dispatcher.Invoke(() =>
            {
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    ShowPopup("Error Updating Fields", errorMsg);
                }
                else
                {
                    // Refresh data
                    var selectedItem = SchedulesComboBox.SelectedItem as ScheduleOption;
                    if (selectedItem?.Schedule != null)
                    {
                        LoadScheduleData(selectedItem.Schedule);
                        
                        // Restore itemize setting and refresh view
                        bool itemize = true;
                        if (_scheduleItemizeSettings.ContainsKey(selectedItem.Id))
                        {
                            itemize = _scheduleItemizeSettings[selectedItem.Id];
                        }
                        RefreshScheduleView(itemize);
                        ApplyCurrentSortLogic();
                    }
                    ShowPopup("Success", "Schedule fields updated successfully.");
                }
            });
        }

        private void ParametersClose_Click(object sender, RoutedEventArgs e)
        {
            ParametersPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void ParametersPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        private void RenamePopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
            // Optionally close on background click
            // RenamePopupOverlay.Visibility = Visibility.Collapsed;
        }

        private void SheetsDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // Clear cell selection
                if (SheetsDataGrid != null && SheetsDataGrid.SelectedCells.Count > 0)
                {
                    SheetsDataGrid.SelectedCells.Clear();
                    SheetsDataGrid.CurrentCell = new DataGridCellInfo();
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Space)
            {
                var dataGrid = sender as DataGrid;
                if (dataGrid?.SelectedItems != null && dataGrid.SelectedItems.Count > 0)
                {
                    // Check if items are SheetItem (Duplicate Sheets mode)
                    if (dataGrid.SelectedItems[0] is SheetItem)
                    {
                        e.Handled = true;
                        
                        var selectedSheets = dataGrid.SelectedItems.Cast<SheetItem>().ToList();
                        bool newState = !selectedSheets.First().IsSelected;
                        
                        foreach (var sheet in selectedSheets)
                        {
                            sheet.IsSelected = newState;
                        }
                        
                        dataGrid.Items.Refresh();
                    }
                }
            }
        }

        private void FillHandle_DragDelta(object sender, DragDeltaEventArgs e)
        {
            var thumb = sender as Thumb;
            if (thumb != null)
            {
                // Find the target cell under the mouse
                var point = Mouse.GetPosition(SheetsDataGrid);
                var hitResult = VisualTreeHelper.HitTest(SheetsDataGrid, point);
                if (hitResult == null) return;
                
                DataGridCell targetCell = FindVisualParent<DataGridCell>(hitResult.VisualHit);
                if (targetCell == null) return;

                // Stop if we are over the same cell as the start
                // Or checking if selection actually needs change could be optimization
                
                // Get Anchor (Current Cell)
                var anchorInfo = SheetsDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                var anchorItem = anchorInfo.Item;
                var anchorCol = anchorInfo.Column;
                
                // Resolve Indices
                int anchorRowIdx = SheetsDataGrid.Items.IndexOf(anchorItem);
                int anchorColIdx = anchorCol.DisplayIndex;
                
                int targetRowIdx = SheetsDataGrid.Items.IndexOf(targetCell.DataContext);
                int targetColIdx = targetCell.Column.DisplayIndex;
                
                if (anchorRowIdx < 0 || targetRowIdx < 0) return;
                
                // Determine Range
                int minRow = Math.Min(anchorRowIdx, targetRowIdx);
                int maxRow = Math.Max(anchorRowIdx, targetRowIdx);
                int minCol = Math.Min(anchorColIdx, targetColIdx);
                int maxCol = Math.Max(anchorColIdx, targetColIdx);
                
                SheetsDataGrid.SelectedCells.Clear();
                
                // Select Range
                for (int r = minRow; r <= maxRow; r++)
                {
                    var item = SheetsDataGrid.Items[r];
                    for (int c = minCol; c <= maxCol; c++)
                    {
                        var col = SheetsDataGrid.Columns[c];
                        SheetsDataGrid.SelectedCells.Add(new DataGridCellInfo(item, col));
                    }
                }
                
                UpdateSelectionAdorner(); // Force update during drag
            }
        }

        private enum SmartFillMode { Copy, Series }

        private void FillHandle_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                PerformAutoFill(SmartFillMode.Copy);
            }
            else
            {
                PerformAutoFill(SmartFillMode.Series);
            }
        }

        private void PerformAutoFill(SmartFillMode mode)
        {
            try
            {
                if (SheetsDataGrid.SelectedCells.Count < 2) return;

                // 1. Get Anchor Value (CurrentCell)
                var anchorInfo = SheetsDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                var anchorRow = anchorInfo.Item as System.Data.DataRowView;
                if (anchorRow == null || _currentScheduleData == null) return;

                // Safely get column name or identify if it's the CheckBox column
                string colName = "";
                if (anchorInfo.Column.Header is CheckBox)
                {
                    colName = "IsSelected";
                }
                else
                {
                    colName = anchorInfo.Column.Header?.ToString();
                }

                if (string.IsNullOrEmpty(colName)) return;

                // Special handling for IsSelected (Checkbox)
                if (colName == "IsSelected")
                {
                    PerformCheckboxAutoFill();
                    return;
                }

                // Check if Anchor passes validation (e.g. not ReadOnly)
                // Actually, Anchor value is valid. We need to check Targets.
                
                string sourceValue = anchorRow[colName]?.ToString() ?? "";
                
                int anchorRowIndex = SheetsDataGrid.Items.IndexOf(anchorRow);

                // 2. Prepare Updates
                var updates = new List<ParameterUpdateInfo>();
                bool hasTypeParameters = false;
                var affectedTypeParams = new HashSet<string>();

                foreach (var cellInfo in SheetsDataGrid.SelectedCells)
                {
                    // Skip if it's the anchor itself (Reference equality check on Item and Column)
                    if (cellInfo.Item == anchorInfo.Item && cellInfo.Column == anchorInfo.Column) continue;

                    var targetRow = cellInfo.Item as System.Data.DataRowView;
                    if (targetRow == null) continue;

                    var targetCol = cellInfo.Column;
                    string targetColName = targetCol.Header?.ToString();
                    
                    // We only auto-fill within the SAME column usually?
                    // Excel allows filling across columns if dragging corner?
                    // Typically dragging corner fills the selected range.
                    // If I select a 2x2 range, and drag...
                    // But here we are dragging the Grip of the Anchor.
                    // The selection expands.
                    // Typically we fill with the Anchor's value into ALL selected cells.
                    // But usually you only fill into same-column cells unless standard Copy/Paste.
                    // Let's assume SAME COLUMN as Anchor for safety?
                    // Or follow Selection?
                    // If I drag right, I copy to right.
                    // If I drag down, I copy down.
                    // Let's allow all selected cells.
                    
                    if (string.IsNullOrEmpty(targetColName)) continue;
                    
                    // Check if Column is ReadOnly
                    if (targetCol.IsReadOnly) continue;

                    // Get IDs
                    if (!targetRow.Row.Table.Columns.Contains("ElementId")) continue;
                    string elementIdStr = targetRow["ElementId"]?.ToString();
                    if (string.IsNullOrEmpty(elementIdStr)) continue;

                    // Find Parameter ID
                    int colIndex = _currentScheduleData.Columns.IndexOf(targetColName);
                    if (colIndex < 0 || colIndex >= _currentScheduleData.ParameterIds.Count) continue;
                    ElementId paramId = _currentScheduleData.ParameterIds[colIndex];

                    // Check Type Parameter
                    bool isType = _currentScheduleData.IsTypeParameter.ContainsKey(targetColName) && 
                                  _currentScheduleData.IsTypeParameter[targetColName];

                    if (isType) 
                    {
                        hasTypeParameters = true;
                        affectedTypeParams.Add(targetColName);
                    }

                    // Determine New Value
                    string newValue;
                    if (mode == SmartFillMode.Series)
                    {
                        int targetRowIndex = SheetsDataGrid.Items.IndexOf(targetRow);
                        int offset = targetRowIndex - anchorRowIndex;
                        newValue = GetSequentialValue(sourceValue, offset);
                    }
                    else
                    {
                        newValue = sourceValue;
                    }

                    // Check if value actually changes (Optimization)
                    string currentVal = targetRow[targetColName]?.ToString() ?? "";
                    if (currentVal == newValue) continue;

                    updates.Add(new ParameterUpdateInfo
                    {
                        ElementIdStr = elementIdStr,
                        ParameterId = paramId,
                        NewValue = newValue,
                        ColumnName = targetColName,
                        IsTypeParameter = isType,
                        Row = targetRow
                    });
                }

                if (updates.Count == 0) return;

                // 3. Confirm Type Parameters
                if (hasTypeParameters)
                {
                    string paramNames = string.Join(", ", affectedTypeParams);
                    ShowConfirmationPopup(
                        "Batch Update Type Parameters",
                        $"You are about to update Type Parameters ({paramNames}).\nThis will affect ALL elements of the corresponding types.\n\nProceed with Auto-Fill?",
                        () => ExecuteBatchUpdates(updates)
                    );
                }
                if (updates.Count > 0)
                {
                    ExecuteBatchUpdates(updates);
                }
            }
            catch (Exception ex)
            {
                ShowPopup("AutoFill Error", ex.Message);
            }
        }

        private void PerformCheckboxAutoFill()
        {
            try
            {
                var anchorInfo = SheetsDataGrid.CurrentCell;
                if (!anchorInfo.IsValid) return;

                bool anchorValue = false;

                // determining anchor value
                if (anchorInfo.Item is SheetItem sheet)
                {
                    anchorValue = sheet.IsSelected;
                }
                else if (anchorInfo.Item is System.Data.DataRowView row)
                {
                    if (row.Row.Table.Columns.Contains("IsSelected") && row["IsSelected"] != DBNull.Value)
                    {
                        anchorValue = Convert.ToBoolean(row["IsSelected"]);
                    }
                }

                int updatedCount = 0;

                foreach (var cellInfo in SheetsDataGrid.SelectedCells)
                {
                    // Skip if it's the anchor itself
                    if (cellInfo.Item == anchorInfo.Item && cellInfo.Column == anchorInfo.Column) continue;

                    // Apply value
                    if (cellInfo.Item is SheetItem targetSheet)
                    {
                        if (targetSheet.IsSelected != anchorValue)
                        {
                            targetSheet.IsSelected = anchorValue;
                            updatedCount++;
                        }
                    }
                    else if (cellInfo.Item is System.Data.DataRowView targetRow)
                    {
                        if (targetRow.Row.Table.Columns.Contains("IsSelected"))
                        {
                            bool currentValue = targetRow["IsSelected"] != DBNull.Value && Convert.ToBoolean(targetRow["IsSelected"]);
                            if (currentValue != anchorValue)
                            {
                                targetRow["IsSelected"] = anchorValue;
                                updatedCount++;
                            }
                        }
                    }
                }

                // If items are in DataTable mode, we might need to notify UI or just rely on Binding.
                // Normally DataRowView changes reflect in UI if bound properly.
            }
            catch (Exception ex)
            {
                ShowPopup("Selection Fill Error", ex.Message);
            }
        }
        
        private void UpdateSelectionAdorner()
        {
            if (SelectionCanvas == null) return;

            if (SheetsDataGrid.SelectedCells.Count == 0)
            {
                SelectionBox.Visibility = System.Windows.Visibility.Collapsed;
                FillHandle.Visibility = System.Windows.Visibility.Collapsed;
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
                return;
            }

            // Calculate Union of Visible Cells
            System.Windows.Rect unionRect = System.Windows.Rect.Empty;
            bool hasVisibleCells = false;

            foreach (var cellInfo in SheetsDataGrid.SelectedCells)
            {
                var row = SheetsDataGrid.ItemContainerGenerator.ContainerFromItem(cellInfo.Item) as DataGridRow;
                if (row == null) continue; // Row not loaded/visible

                var col = cellInfo.Column;
                var cellContent = col.GetCellContent(row);
                if (cellContent == null) continue; // Cell not loaded

                var cell = cellContent.Parent as FrameworkElement;
                if (cell == null) continue;

                // Create Rect for this cell
                System.Windows.Point p = cell.TranslatePoint(new System.Windows.Point(0, 0), SelectionCanvas);
                System.Windows.Rect cellRect = new System.Windows.Rect(p, new System.Windows.Size(cell.ActualWidth, cell.ActualHeight));

                if (unionRect == System.Windows.Rect.Empty)
                    unionRect = cellRect;
                else
                    unionRect.Union(cellRect);

                hasVisibleCells = true;
            }

            if (!hasVisibleCells)
            {
                SelectionBox.Visibility = System.Windows.Visibility.Collapsed;
                FillHandle.Visibility = System.Windows.Visibility.Collapsed;
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
                return;
            }

            // Update UI
            SelectionBox.Width = unionRect.Width;
            SelectionBox.Height = unionRect.Height;
            Canvas.SetLeft(SelectionBox, unionRect.Left);
            Canvas.SetTop(SelectionBox, unionRect.Top);
            SelectionBox.Visibility = System.Windows.Visibility.Visible;

            // Fill Handle (Bottom Right)
            Canvas.SetLeft(FillHandle, unionRect.Right - 6); 
            Canvas.SetTop(FillHandle, unionRect.Bottom - 6);
            FillHandle.Visibility = System.Windows.Visibility.Visible;
            
            // Copy Indicator
            if (IsCopyMode)
            {
                FillIndicator.Visibility = System.Windows.Visibility.Visible;
                // Position at Top Right of the square (FillHandle)
                // FillHandle is 6x6, placed at (Right-6, Bottom-6).
                // So FillHandle Top-Right is (Right, Bottom-6).
                // We place the "+" so its bottom-left is roughly there?
                // Or center it?
                // User said: "top right of the square in the corner" and "2x its size"
                
                // Let's place it slightly offset to look "top right"
                Canvas.SetLeft(FillIndicator, unionRect.Right);
                Canvas.SetTop(FillIndicator, unionRect.Bottom - 20); // Moved up to sit above/corner
            }
            else
            {
                FillIndicator.Visibility = System.Windows.Visibility.Collapsed;
            }
        }



        private string GetSequentialValue(string original, int offset)
        {
            if (string.IsNullOrEmpty(original)) return original;

            // Find the last sequence of digits
            var matches = System.Text.RegularExpressions.Regex.Matches(original, @"\d+");
            if (matches.Count > 0)
            {
                var lastMatch = matches[matches.Count - 1];
                long number;
                if (long.TryParse(lastMatch.Value, out number))
                {
                    long newNumber = number + offset;
                    
                    // Format preservation: if original was "001", try to keep "002" length
                    // If "1", keep "2". 
                    // Use zero padding if original had it.
                    string format = new string('0', lastMatch.Length);
                    // Check if actually zero-padded
                    bool isZeroPadded = lastMatch.Value.StartsWith("0") && lastMatch.Value.Length > 1;
                    
                    string newNumStr = isZeroPadded ? newNumber.ToString(format) : newNumber.ToString();
                    
                    string prefix = original.Substring(0, lastMatch.Index);
                    string suffix = original.Substring(lastMatch.Index + lastMatch.Length);
                    
                    return prefix + newNumStr + suffix;
                }
            }
            
            return original;
        }

        private void ExecuteBatchUpdates(List<ParameterUpdateInfo> updates)
        {
            if (updates == null || updates.Count == 0) return;

            var batchData = new List<ParameterBatchData>();
            foreach (var update in updates)
            {
                batchData.Add(new ParameterBatchData
                {
                    ElementIdStr = update.ElementIdStr,
                    ParameterId = update.ParameterId,
                    Value = update.NewValue
                });
            }

            _parameterValueUpdateHandler.IsBatchMode = true;
            _parameterValueUpdateHandler.BatchData = batchData;
            _parameterValueUpdateExternalEvent.Raise();
        }

        private class ParameterUpdateInfo
        {
            public string ElementIdStr { get; set; }
            public ElementId ParameterId { get; set; }
            public string NewValue { get; set; }
            public string ColumnName { get; set; }
            public bool IsTypeParameter { get; set; }
            public System.Data.DataRowView Row { get; set; }
        }

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            var parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent) return parent;
            return FindVisualParent<T>(parentObject);
        }

        private void DataGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var dep = e.OriginalSource as DependencyObject;
            while (dep != null && !(dep is DataGridCell) && !(dep is DataGridColumnHeader))
            {
                if (dep is CheckBox)
                {
                    e.Handled = true;
                    var checkbox = dep as CheckBox;
                    checkbox.IsChecked = !checkbox.IsChecked;
                    return;
                }
                dep = VisualTreeHelper.GetParent(dep);
            }
        }



        private void SheetsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectionAdorner();
        }

        private void SheetsDataGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var scrollViewer = GetScrollViewer(SheetsDataGrid);
            if (scrollViewer == null) return;

            // Handle horizontal scrolling with Shift + MouseWheel OR horizontal wheel
            if (Keyboard.Modifiers == ModifierKeys.Shift || e.Delta == 0)
            {
                e.Handled = true;
                
                // For horizontal wheel, Delta is 0 but we need to check the actual MouseDevice
                // In WPF, horizontal scrolling is typically handled via MouseWheel with Shift modifier
                // However, if we detect Shift is pressed, we definitely want horizontal scroll
                if (Keyboard.Modifiers == ModifierKeys.Shift)
                {
                    if (e.Delta > 0)
                        scrollViewer.LineLeft();
                    else
                        scrollViewer.LineRight();
                }
            }
        }


        private static ScrollViewer GetScrollViewer(DependencyObject depObj)
        {
            if (depObj is ScrollViewer) return depObj as ScrollViewer;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);
                var result = GetScrollViewer(child);
                if (result != null) return result;
            }
            return null;
        }

        private void SheetSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = (sender as System.Windows.Controls.TextBox)?.Text?.ToLowerInvariant() ?? "";
            
            FilteredSheets.Clear();
            foreach (var sheet in Sheets.Where(s => 
                string.IsNullOrEmpty(searchText) || 
                s.Name.ToLowerInvariant().Contains(searchText) ||
                s.SheetNumber.ToLowerInvariant().Contains(searchText)))
            {
                FilteredSheets.Add(sheet);
            }
        }

        private void ClearSheetSearch_Click(object sender, RoutedEventArgs e)
        {
            SheetSearchBox.Clear();
        }

        private void SheetCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is SheetItem clickedItem)
            {
                e.Handled = true;
                bool newState = !(checkBox.IsChecked ?? false);
                checkBox.IsChecked = newState;

                if (SheetsDataGrid.SelectedItems.Contains(clickedItem))
                {
                    foreach (SheetItem item in SheetsDataGrid.SelectedItems)
                    {
                        if (item != clickedItem)
                        {
                            item.IsSelected = newState;
                        }
                    }
                }
            }
        }

        private void SelectAllSheets_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in FilteredSheets)
            {
                sheet.IsSelected = true;
            }
        }

        private void SelectAllSheets_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var sheet in FilteredSheets)
            {
                sheet.IsSelected = false;
            }
        }

        #endregion

        #region Theme

        private ResourceDictionary _currentThemeDictionary;

        private void LoadTheme()
        {
            try
            {
                var themeUri = new Uri(_isDarkMode 
                    ? "pack://application:,,,/ProSchedules;component/UI/Themes/DarkTheme.xaml" 
                    : "pack://application:,,,/ProSchedules;component/UI/Themes/LightTheme.xaml", UriKind.Absolute);
                
                var newDict = new ResourceDictionary { Source = themeUri };
                
                if (_currentThemeDictionary != null)
                {
                    Resources.MergedDictionaries.Remove(_currentThemeDictionary);
                }
                
                Resources.MergedDictionaries.Add(newDict);
                _currentThemeDictionary = newDict;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading theme: {ex.Message}");
            }
        }

        private void ToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = ThemeToggleButton.IsChecked == true;
            LoadTheme();
            SaveThemeState();

            var icon = ThemeToggleButton?.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                       as MaterialDesignThemes.Wpf.PackIcon;
            if (icon != null)
            {
                icon.Kind = _isDarkMode
                    ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                    : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
            }
        }

        private void LoadThemeState()
        {
            try
            {
                var config = LoadConfig();
                if (TryGetBool(config, "IsDarkMode", out var isDark))
                {
                    _isDarkMode = isDark;
                }
            }
            catch (Exception)
            {
            }

            if (ThemeToggleButton != null)
            {
                ThemeToggleButton.IsChecked = _isDarkMode;
                var icon = ThemeToggleButton.Template?.FindName("ThemeToggleIcon", ThemeToggleButton)
                           as MaterialDesignThemes.Wpf.PackIcon;
                if (icon != null)
                {
                    icon.Kind = _isDarkMode
                        ? MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOffOutline
                        : MaterialDesignThemes.Wpf.PackIconKind.ToggleSwitchOutline;
                }
            }
            
            LoadTheme();
        }

        private void SaveThemeState()
        {
            try
            {
                var config = LoadConfig();
                config["IsDarkMode"] = _isDarkMode;
                SaveConfig(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Custom Popup



        public void ShowPopup(string title, string message, Action onCloseAction = null)
        {
            PopupTitle.Text = title;
            PopupMessage.Text = message;
            _onPopupClose = onCloseAction;
            MainContentGrid.IsEnabled = false;
            PopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ClosePopup()
        {
            MainContentGrid.IsEnabled = true;
            PopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onPopupClose != null)
            {
                var action = _onPopupClose;
                _onPopupClose = null;
                action.Invoke();
            }
        }

        private void PopupClose_Click(object sender, RoutedEventArgs e)
        {
            ClosePopup();
        }

        private void PopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }



        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                IsCopyMode = true;
            }
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
            {
                IsCopyMode = false;
            }
        }
        
        public void ShowConfirmationPopup(string title, string message, Action onConfirmAction, Action onCancelAction = null)
        {
            ConfirmationTitle.Text = title;
            ConfirmationMessage.Text = message;
            _onConfirmAction = onConfirmAction;
            _onCancelAction = onCancelAction;
            MainContentGrid.IsEnabled = false;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }

        private void ConfirmationOK_Click(object sender, RoutedEventArgs e)
        {
            MainContentGrid.IsEnabled = true;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onConfirmAction != null)
            {
                var action = _onConfirmAction;
                _onConfirmAction = null;
                _onCancelAction = null;
                action.Invoke();
            }
        }

        private void ConfirmationCancel_Click(object sender, RoutedEventArgs e)
        {
            MainContentGrid.IsEnabled = true;
            ConfirmationPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            if (_onCancelAction != null)
            {
                var action = _onCancelAction;
                _onCancelAction = null;
                _onConfirmAction = null;
                action.Invoke();
            }
            else
            {
                _onConfirmAction = null;
            }
        }

        private void ShowConfirmPopup(string title, string message, Action onConfirmAction, string confirmButtonText = "Discard")
        {
            ConfirmPopupTitle.Text = title;
            ConfirmPopupMessage.Text = message;
            _onConfirmAction = onConfirmAction;
            ConfirmActionButton.Content = confirmButtonText;
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Visible;
        }


        private void CloseConfirmPopup()
        {
            ConfirmPopupOverlay.Visibility = System.Windows.Visibility.Collapsed;
            _onConfirmAction = null;
        }

        private void ConfirmDiscard_Click(object sender, RoutedEventArgs e)
        {
            var action = _onConfirmAction;
            CloseConfirmPopup();
            action?.Invoke();
        }

        private void ConfirmCancel_Click(object sender, RoutedEventArgs e)
        {
            CloseConfirmPopup();
        }

        private void ConfirmPopupBackground_Click(object sender, MouseButtonEventArgs e)
        {
        }

        #endregion

        #region Window chrome / resize handlers

        private void TitleBar_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void LeftEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void RightEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeWE;
        private void BottomEdge_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNS;
        private void Edge_MouseLeave(object sender, MouseEventArgs e) => Cursor = Cursors.Arrow;
        private void BottomLeftCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNESW;
        private void BottomRightCorner_MouseEnter(object sender, MouseEventArgs e) => Cursor = Cursors.SizeNWSE;

        private void Window_MouseMove(object sender, MouseEventArgs e) => _windowResizer.ResizeWindow(e);
        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) => _windowResizer.StopResizing();
        private void LeftEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Left);
        private void RightEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Right);
        private void BottomEdge_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.Bottom);
        private void BottomLeftCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomLeft);
        private void BottomRightCorner_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _windowResizer.StartResizing(e, ResizeDirection.BottomRight);

        #endregion


        #region Window Startup

        private void DeferWindowShow()
        {
            Opacity = 0;
            Loaded += DuplicateSheetsWindow_Loaded;
        }

        private void DuplicateSheetsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            TryShowWindow();
        }

        private void TryShowWindow()
        {
            if (!_isDataLoaded)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                Opacity = 1;
            }), DispatcherPriority.Render);
        }

        #endregion

        #region Window State

        private void LoadWindowState()
        {
            try
            {
                var config = LoadConfig();
                bool hasLeft = TryGetDouble(config, WindowLeftKey, out var left);
                bool hasTop = TryGetDouble(config, WindowTopKey, out var top);
                bool hasWidth = TryGetDouble(config, WindowWidthKey, out var width);
                bool hasHeight = TryGetDouble(config, WindowHeightKey, out var height);

                bool hasSize = hasWidth && hasHeight && width > 0 && height > 0;
                bool hasPos = hasLeft && hasTop && !double.IsNaN(left) && !double.IsNaN(top);

                if (!hasSize && !hasPos)
                {
                    return;
                }

                WindowStartupLocation = WindowStartupLocation.Manual;

                if (hasSize)
                {
                    Width = Math.Max(MinWidth, width);
                    Height = Math.Max(MinHeight, height);
                }

                if (hasPos)
                {
                    Left = left;
                    Top = top;
                }
            }
            catch (Exception)
            {
            }
        }

        private void SaveWindowState()
        {
            try
            {
                var config = LoadConfig();
                var bounds = WindowState == WindowState.Normal
                    ? new Rect(Left, Top, Width, Height)
                    : RestoreBounds;

                config[WindowLeftKey] = bounds.Left;
                config[WindowTopKey] = bounds.Top;
                config[WindowWidthKey] = bounds.Width;
                config[WindowHeightKey] = bounds.Height;

                SaveConfig(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save window state: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Config Helpers

        private Dictionary<string, object> LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (config != null)
                    {
                        return config;
                    }
                }
            }
            catch (Exception)
            {
            }

            return new Dictionary<string, object>();
        }

        private void SaveConfig(Dictionary<string, object> config)
        {
            var dir = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllText(ConfigFilePath, JsonConvert.SerializeObject(config, Formatting.Indented));
        }

        private static bool TryGetBool(Dictionary<string, object> config, string key, out bool value)
        {
            value = false;
            if (!config.TryGetValue(key, out var raw) || raw == null)
            {
                return false;
            }

            if (raw is bool boolValue)
            {
                value = boolValue;
                return true;
            }

            if (raw is JToken token && token.Type == JTokenType.Boolean)
            {
                value = token.Value<bool>();
                return true;
            }

            if (raw is string text && bool.TryParse(text, out var parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        private static bool TryGetDouble(Dictionary<string, object> config, string key, out double value)
        {
            value = 0;
            if (!config.TryGetValue(key, out var raw) || raw == null)
            {
                return false;
            }

            switch (raw)
            {
                case double doubleValue:
                    value = doubleValue;
                    return true;
                case float floatValue:
                    value = floatValue;
                    return true;
                case decimal decimalValue:
                    value = (double)decimalValue;
                    return true;
                case long longValue:
                    value = longValue;
                    return true;
                case int intValue:
                    value = intValue;
                    return true;
                case JToken token when token.Type == JTokenType.Float || token.Type == JTokenType.Integer:
                    value = token.Value<double>();
                    return true;
                case string text when double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed):
                    value = parsed;
                    return true;
            }

            return false;
        }

        private void SaveScheduleName(string scheduleName)
        {
            try
            {
                string folder = GetProjectSettingsFolder();
                string file = System.IO.Path.Combine(folder, "last_schedule.txt");
                System.IO.File.WriteAllText(file, scheduleName ?? "");
                _lastSelectedScheduleName = scheduleName;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving schedule name: {ex.Message}");
            }
        }

        private string GetSavedScheduleName()
        {
            try
            {
                string folder = GetProjectSettingsFolder();
                string file = System.IO.Path.Combine(folder, "last_schedule.txt");
                if (System.IO.File.Exists(file))
                {
                    string scheduleName = System.IO.File.ReadAllText(file);
                    _lastSelectedScheduleName = scheduleName;
                    return scheduleName;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading schedule name: {ex.Message}");
            }
            
            return null;
        }


        #endregion

        #region Helpers

        private string GetProjectSettingsFolder()
        {
            try
            {
                string docTitle = "Default";
                if (_uiApplication?.ActiveUIDocument?.Document != null)
                {
                    docTitle = _uiApplication.ActiveUIDocument.Document.Title;
                    // Remove extension if present
                    if (docTitle.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                    {
                        docTitle = docTitle.Substring(0, docTitle.Length - 4);
                    }
                }

                // Use Roaming AppData for user settings: %APPDATA%\RK Tools\ProSchedules\{ProjectName}
                string folder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "RK Tools", "ProSchedules", docTitle);

                if (!System.IO.Directory.Exists(folder))
                {
                    System.IO.Directory.CreateDirectory(folder);
                }
                
                return folder;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting settings folder: {ex.Message}");
                // Fallback
                return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RK Tools", "ProSchedules", "Default");
            }
        }

        #endregion

    }
}
