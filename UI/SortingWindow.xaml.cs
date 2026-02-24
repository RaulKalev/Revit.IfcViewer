using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ProSchedules.UI;

namespace ProSchedules.UI
{
    public partial class SortingWindow : Window
    {
        private DuplicateSheetsWindow _parent;
        private List<SortItem> _checkpoint;

        public ObservableCollection<SortItem> SortCriteria { get; private set; }
        public ObservableCollection<string> AvailableSortColumns { get; private set; }

        public SortingWindow(DuplicateSheetsWindow parent)
        {
            InitializeComponent();
            _parent = parent;
            
            // Link to parent collections
            SortCriteria = parent.SortCriteria;
            AvailableSortColumns = parent.AvailableSortColumns;
            
            DataContext = this;



            // Create initial checkpoint
            CreateCheckpoint();
            
            // Handle window move via TitleBar drag (handled by TitleBar control usually, but checking DuplicateSheetsWindow logic)
            // TitleBar control usually has logic. If not, we can add standard drag logic.
            // DuplicateSheetsWindow is WindowStyle=None, so it likely handles dragging.
        }

        private void CreateCheckpoint()
        {
            _checkpoint = new List<SortItem>();
            foreach (var item in SortCriteria) _checkpoint.Add(item.Clone());
        }

        private void RestoreCheckpoint()
        {
            SortCriteria.Clear();
            foreach (var item in _checkpoint) SortCriteria.Add(item.Clone());
        }

        private void AddSortLevel_Click(object sender, RoutedEventArgs e)
        {
            SortCriteria.Add(new SortItem { SelectedColumn = "(none)", IsAscending = true });
        }

        private void RemoveSortItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is SortItem item)
            {
                SortCriteria.Remove(item);
            }
        }

        private void SortApply_Click(object sender, RoutedEventArgs e)
        {
            // Apply logic on parent
            _parent.ApplyCurrentSortLogicInternal();
            
            // Update checkpoint since we are committing
            CreateCheckpoint();
            
            // Do NOT close window (requested behavior)
        }

        private void SortCancel_Click(object sender, RoutedEventArgs e)
        {
            RestoreCheckpoint();
            Close();
        }
    }
}
