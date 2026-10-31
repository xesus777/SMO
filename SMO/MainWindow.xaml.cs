using SMO.Pages;
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

namespace SMO
{
    public partial class MainWindow : Window
    {
        public SmoData SharedData { get; } = new SmoData();

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(new Step1Page(this));
        }
    }
}
