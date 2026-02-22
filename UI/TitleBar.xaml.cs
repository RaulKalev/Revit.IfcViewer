using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace IfcViewer.UI
{
    public partial class TitleBar : UserControl
    {
        public TitleBar()
        {
            InitializeComponent();
        }

        public static readonly DependencyProperty IsMinimizeVisibleProperty =
            DependencyProperty.Register("IsMinimizeVisible", typeof(bool), typeof(TitleBar), new PropertyMetadata(true));

        public bool IsMinimizeVisible
        {
            get { return (bool)GetValue(IsMinimizeVisibleProperty); }
            set { SetValue(IsMinimizeVisibleProperty, value); }
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Window.GetWindow(this)?.DragMove();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.GetWindow(this);
            if (window != null)
            {
                window.WindowState = WindowState.Minimized;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Window.GetWindow(this)?.Close();
        }
        private void ToggleFeatureButton_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ToggleFeatureButton_Unchecked(object sender, RoutedEventArgs e)
        {

        }

    }
}

