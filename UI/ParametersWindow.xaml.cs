using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using ProSchedules.Models;

namespace ProSchedules.UI
{


    public partial class ParametersWindow : Window
    {
        public ObservableCollection<ParameterItem> AvailableParams { get; set; } = new ObservableCollection<ParameterItem>();
        public ObservableCollection<ParameterItem> ScheduledParams { get; set; } = new ObservableCollection<ParameterItem>();

        public event Action<List<ParameterItem>> OnApply;

        public ParametersWindow(List<ParameterItem> available, List<ParameterItem> scheduled, string categoryName)
        {
            InitializeComponent();
            DataContext = this;
            Title = $"Schedule Parameters - {categoryName}";

            foreach (var p in available) AvailableParams.Add(p);
            foreach (var p in scheduled) ScheduledParams.Add(p);

            AvailableParamsList.ItemsSource = AvailableParams;
            ScheduledParamsList.ItemsSource = ScheduledParams;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            var selected = AvailableParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            foreach (var item in selected)
            {
                AvailableParams.Remove(item);
                ScheduledParams.Add(item);
            }
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            foreach (var item in selected)
            {
                ScheduledParams.Remove(item);
                AvailableParams.Add(item);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            if (selected.Count == 0) return;

            // Sort by index to move them properly
            var sortedSelected = selected.OrderBy(x => ScheduledParams.IndexOf(x)).ToList();

            foreach (var item in sortedSelected)
            {
                int index = ScheduledParams.IndexOf(item);
                if (index > 0)
                {
                    ScheduledParams.Move(index, index - 1);
                }
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            var selected = ScheduledParamsList.SelectedItems.Cast<ParameterItem>().ToList();
            if (selected.Count == 0) return;

            // Sort reverse to move from bottom up
            var sortedSelected = selected.OrderByDescending(x => ScheduledParams.IndexOf(x)).ToList();

            foreach (var item in sortedSelected)
            {
                int index = ScheduledParams.IndexOf(item);
                if (index < ScheduledParams.Count - 1)
                {
                    ScheduledParams.Move(index, index + 1);
                }
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            OnApply?.Invoke(ScheduledParams.ToList());
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #region Drag and Drop
        
        private System.Windows.Point _startPoint;

        private void ScheduledParamsList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);
        }

        private void ScheduledParamsList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                System.Windows.Point position = e.GetPosition(null);
                if (Math.Abs(position.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    ListBox listBox = sender as ListBox;
                    ListBoxItem listBoxItem = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);

                    if (listBoxItem == null) return;

                    ParameterItem item = (ParameterItem)listBox.ItemContainerGenerator.ItemFromContainer(listBoxItem);

                    if (item != null)
                    {
                        DataObject data = new DataObject("ParameterItem", item);
                        DragDrop.DoDragDrop(listBoxItem, data, DragDropEffects.Move);
                    }
                }
            }
        }

        private void ScheduledParamsList_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("ParameterItem"))
            {
                e.Effects = DragDropEffects.Move;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void ScheduledParamsList_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("ParameterItem"))
            {
                ParameterItem droppedData = e.Data.GetData("ParameterItem") as ParameterItem;
                
                // Find the target item under the mouse
                ListBoxItem targetItem = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
                
                if (targetItem != null)
                {
                    ParameterItem target = targetItem.DataContext as ParameterItem;

                    if (droppedData != null && target != null && droppedData != target)
                    {
                        int oldIndex = ScheduledParams.IndexOf(droppedData);
                        int newIndex = ScheduledParams.IndexOf(target);

                        if (oldIndex != -1 && newIndex != -1)
                        {
                            ScheduledParams.Move(oldIndex, newIndex);
                        }
                    }
                }
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            do
            {
                if (current is T)
                {
                    return (T)current;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            while (current != null);
            return null;
        }

        #endregion
    }
}
