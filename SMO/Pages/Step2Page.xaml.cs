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
    public partial class Step2Page : Page
    {
        private MainWindow _mainWindow;
        private bool _isUpdatingFields = false;

        public Step2Page(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;

            var data = _mainWindow.SharedData;
            CbModelType.SelectedIndex = data.ModelTypeIndex;
            TxtLambda.Text = data.LambdaRaw.ToString();
            TxtMu.Text = data.MuRaw.ToString();
            TxtChannels.Text = data.Channels.ToString();
            TxtQueue.Text = data.QueueLength.ToString();
            CbLambdaUnit.SelectedIndex = data.LambdaUnitIndex;
            CbMuUnit.SelectedIndex = data.MuUnitIndex;
            ChbAutoChannels.IsChecked = data.AutoCalculateChannels;
            TxtTargetP.Text = data.TargetProbability.ToString();

            ConfigureUI();
        }

        private void ConfigureUI()
        {
            if (PanelChannelsContainer == null || PanelQueue == null) return;

            bool isMultiChannel = _mainWindow.SharedData.IsMultiChannel;
            int modelType = CbModelType.SelectedIndex;

            // Показываем весь контейнер каналов только для многоканальной
            if (isMultiChannel)
            {
                PanelChannelsContainer.Visibility = Visibility.Visible;
                // Блокируем ввод ручного 'n', если стоит галочка авторасчета
                TxtChannels.IsEnabled = ChbAutoChannels.IsChecked == false;
                TxtTargetP.IsEnabled = ChbAutoChannels.IsChecked == true;
            }
            else
            {
                PanelChannelsContainer.Visibility = Visibility.Collapsed;
                TxtChannels.Text = "1";
            }

            if (modelType == 0)
            {
                PanelQueue.Visibility = Visibility.Collapsed;
                TxtQueue.Text = "0";
            }
            else
            {
                PanelQueue.Visibility = Visibility.Visible;
                LblQueueTitle.Text = modelType == 1 ? "Длина очереди, m:" : "Число заявок в очереди, m:";
            }
        }

        private void ModelConfiguration_Changed(object sender, SelectionChangedEventArgs e)
        {
            ConfigureUI();
        }

        private void ChbAutoChannels_Changed(object sender, RoutedEventArgs e)
        {
            ConfigureUI();
        }

        private void CbMuUnit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LblTobsUnit == null || CbMuUnit == null) return;
            var selectedItem = (ComboBoxItem)CbMuUnit.SelectedItem;
            LblTobsUnit.Text = selectedItem.Content.ToString();

            if (!_isUpdatingFields)
            {
                TxtMu.Text = "0";
                TxtTobs.Text = "0";
            }
        }

        private void TxtMu_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingFields) return;
            _isUpdatingFields = true;
            if (double.TryParse(TxtMu.Text, out double mu) && mu > 0)
                TxtTobs.Text = Math.Round(1.0 / mu, 4).ToString();
            else
                TxtTobs.Text = "0";
            _isUpdatingFields = false;
        }

        private void TxtTobs_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingFields) return;
            _isUpdatingFields = true;
            if (double.TryParse(TxtTobs.Text, out double tobs) && tobs > 0)
                TxtMu.Text = Math.Round(1.0 / tobs, 4).ToString();
            else
                TxtMu.Text = "0";
            _isUpdatingFields = false;
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _mainWindow.MainFrame.GoBack();
        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _mainWindow.SharedData;
                data.ModelTypeIndex = CbModelType.SelectedIndex;
                data.LambdaRaw = Convert.ToDouble(TxtLambda.Text);
                data.MuRaw = Convert.ToDouble(TxtMu.Text);
                data.LambdaUnitIndex = CbLambdaUnit.SelectedIndex;
                data.MuUnitIndex = CbMuUnit.SelectedIndex;
                data.QueueLength = Convert.ToInt32(TxtQueue.Text);

                data.AutoCalculateChannels = ChbAutoChannels.IsChecked == true && data.IsMultiChannel;
                data.TargetProbability = Convert.ToDouble(TxtTargetP.Text);

                if (!data.AutoCalculateChannels)
                {
                    data.Channels = Convert.ToInt32(TxtChannels.Text);
                }

                if (data.Lambda <= 0 || data.Mu <= 0 || (!data.AutoCalculateChannels && data.Channels <= 0))
                {
                    MessageBox.Show("Значения параметров λ, μ и n должны быть больше нуля!", "Ошибка валидации", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (data.AutoCalculateChannels && (data.TargetProbability <= 0 || data.TargetProbability >= 1))
                {
                    MessageBox.Show("Требуемая вероятность P должна быть в диапазоне от 0 до 1 (не включая границы)!", "Ошибка валидации", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _mainWindow.MainFrame.Navigate(new Step3Page(_mainWindow));
            }
            catch
            {
                MessageBox.Show("Пожалуйста, проверьте корректность ввода числовых параметров.", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
