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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SMO.Pages
{
    public partial class Step1Page : Page
    {
        private MainWindow _mainWindow;

        public Step1Page(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;

            RbMulti.IsChecked = _mainWindow.SharedData.IsMultiChannel;
            RbSingle.IsChecked = !_mainWindow.SharedData.IsMultiChannel;
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            _mainWindow.SharedData.IsMultiChannel = RbMulti.IsChecked == true;
            _mainWindow.MainFrame.Navigate(new Step2Page(_mainWindow));
        }
    }
}
