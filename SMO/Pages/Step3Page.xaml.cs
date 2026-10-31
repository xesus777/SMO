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
using Excel = Microsoft.Office.Interop.Excel;

namespace SMO.Pages
{

    public partial class Step3Page : Page
    {
        private MainWindow _mainWindow;

        public class ResultRow
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        private class TempMetrics
        {
            public double P0 { get; set; }
            public double Potk { get; set; }
            public double Q { get; set; }
            public double A { get; set; }
            public double L_och { get; set; }
            public double W_och { get; set; }
            public double L_obs { get; set; }
            public double L_sist { get; set; }
            public double W_sist { get; set; }
            public List<double> P_channels { get; set; } = new List<double>();
        }

        public Step3Page(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
            CalculateAndShow();
        }

        private void CalculateAndShow()
        {
            var data = _mainWindow.SharedData;
            double lambda = data.Lambda;
            double mu = data.Mu;
            int m = data.QueueLength;
            int modelType = data.ModelTypeIndex;

            string[] units = { "мин", "час", "сут" };
            string tUnit = units[data.MuUnitIndex];
            string lambdaUnit = $"заяв/{units[data.LambdaUnitIndex]}";
            string muUnit = $"заяв/{units[data.MuUnitIndex]}";

            var results = new List<ResultRow>();

            double rho = lambda / mu;
            double t_obs_converted = data.MuRaw > 0 ? (1.0 / data.MuRaw) : 0;

            results.Add(new ResultRow { Name = "1. Интенсивность нагрузки (ρ)", Value = Math.Round(rho, 4).ToString() });
            results.Add(new ResultRow { Name = $"2. Время обслуживания (t_обс), {tUnit}", Value = Math.Round(t_obs_converted, 4).ToString() });

            double p0 = 0, potk = 0, q = 0, a = 0;
            double L_och = 0, W_och = 0, L_obs = 0, L_sist = 0, W_sist = 0;
            int n = data.Channels;
            List<double> p_channels = new List<double>();

            if (data.AutoCalculateChannels)
            {
                n = 1;
                bool found = false;
                while (n <= 100)
                {
                    var tempMetrics = CalculateMetricsForN(n, lambda, mu, m, modelType, rho);
                    if (tempMetrics.Q >= data.TargetProbability)
                    {
                        p0 = tempMetrics.P0; potk = tempMetrics.Potk; q = tempMetrics.Q; a = tempMetrics.A;
                        L_och = tempMetrics.L_och; W_och = tempMetrics.W_och; L_obs = tempMetrics.L_obs;
                        L_sist = tempMetrics.L_sist; W_sist = tempMetrics.W_sist;
                        p_channels = tempMetrics.P_channels;
                        found = true;
                        break;
                    }
                    n++;
                }

                if (!found)
                {
                    MessageBox.Show("Не удалось подобрать число каналов для P ≥ " + data.TargetProbability + " (требуется > 100 каналов). Взят расчет для n=100.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    n = 100;
                    var tempMetrics = CalculateMetricsForN(n, lambda, mu, m, modelType, rho);
                    p0 = tempMetrics.P0; potk = tempMetrics.Potk; q = tempMetrics.Q; a = tempMetrics.A;
                    L_och = tempMetrics.L_och; W_och = tempMetrics.W_och; L_obs = tempMetrics.L_obs;
                    L_sist = tempMetrics.L_sist; W_sist = tempMetrics.W_sist;
                    p_channels = tempMetrics.P_channels;
                }
                results.Add(new ResultRow { Name = $"[!] Необходимое число каналов для P ≥ {data.TargetProbability}", Value = n.ToString() });
            }
            else
            {
                var metrics = CalculateMetricsForN(n, lambda, mu, m, modelType, rho);
                p0 = metrics.P0; potk = metrics.Potk; q = metrics.Q; a = metrics.A;
                L_och = metrics.L_och; W_och = metrics.W_och; L_obs = metrics.L_obs;
                L_sist = metrics.L_sist; W_sist = metrics.W_sist;
                p_channels = metrics.P_channels;
            }

            a = data.LambdaRaw * q;
            double nominalProd_converted = n * data.MuRaw;
            double factProdPercent = nominalProd_converted > 0 ? (a / nominalProd_converted) * 100.0 : 0;

            double unitFactor = data.LambdaUnitIndex == 1 ? 60.0 : (data.LambdaUnitIndex == 2 ? 1440.0 : 1.0);
            double W_och_converted = W_och / unitFactor;
            double W_sist_converted = W_sist / unitFactor;

            double idleChannels = n - L_obs;
            double busyFactor = n > 0 ? L_obs / n : 0;
            double totalRejectedPerHour = data.LambdaUnitIndex == 1 ? (data.LambdaRaw * potk) : (data.Lambda * potk * 60.0);

            double p_queue_formation = 0;
            if (modelType == 1)
            {
                double sum2 = 0; for (int i = 1; i <= m; i++) sum2 += Math.Pow(rho, n + i) / (Factorial(n) * Math.Pow(n, i));
                p_queue_formation = sum2 * p0;
            }
            else if (modelType == 2)
            {
                double chi = rho / n; if (chi < 1.0) p_queue_formation = (Math.Pow(rho, n) / (Factorial(n) * (1.0 - chi))) * p0;
            }

            results.Add(new ResultRow { Name = "3. Вероятность, что канал свободен (доля времени простоя, P₀)", Value = Math.Round(p0, 4).ToString() });

            if (p_channels.Count > 0)
            {
                results.Add(new ResultRow { Name = "--- Вероятности занятости каналов обслуживанием ---", Value = "" });
                for (int i = 0; i < p_channels.Count; i++)
                {
                    string label = (i == p_channels.Count - 1) ? $"   • Заняты все n ({n}) каналов" : $"   • Занят ровно {i + 1} канал";
                    results.Add(new ResultRow { Name = label, Value = Math.Round(p_channels[i], 4).ToString() });
                }
            }

            results.Add(new ResultRow { Name = "4. Вероятность отказа (доля заявок, получивших отказ, Pотк)", Value = Math.Round(potk, 4).ToString() });
            results.Add(new ResultRow { Name = "5. Вероятность обслуживания поступающих заявок (Q = pобс)", Value = Math.Round(q, 4).ToString() });
            results.Add(new ResultRow { Name = "6. Среднее число каналов, занятых обслуживанием", Value = Math.Round(L_obs, 4).ToString() });
            results.Add(new ResultRow { Name = "   • Среднее число простаивающих каналов", Value = Math.Round(idleChannels, 4).ToString() });
            results.Add(new ResultRow { Name = "7. Коэффициент занятости каналов обслуживанием", Value = Math.Round(busyFactor, 4).ToString() });
            results.Add(new ResultRow { Name = $"8. Абсолютная пропускная способность (A), {lambdaUnit}", Value = Math.Round(a, 4).ToString() });

            if (modelType == 0)
            {
                results.Add(new ResultRow { Name = "9. Среднее время простоя СМО (Доля времени бездействия)", Value = Math.Round(p0, 4).ToString() });
                results.Add(new ResultRow { Name = $"10. Среднее время простоя отдельного канала, {units[data.LambdaUnitIndex]}", Value = Math.Round(p0 / mu, 4).ToString() });
            }

            if (modelType > 0)
            {
                results.Add(new ResultRow { Name = "--- Очередь и Система ---", Value = "" });
                results.Add(new ResultRow { Name = "Вероятность образования очереди", Value = Math.Round(p_queue_formation, 4).ToString() });
                results.Add(new ResultRow { Name = "Среднее число заявок, находящихся в очереди (L_оч), шт", Value = Math.Round(L_och, 4).ToString() });
                results.Add(new ResultRow { Name = $"11. Среднее время ожидания обслуживания заявки в очереди (W_оч), {units[data.LambdaUnitIndex]}", Value = Math.Round(W_och_converted, 4).ToString() });
                results.Add(new ResultRow { Name = "12. Среднее число обслуживаемых заявок, шт", Value = Math.Round(L_obs, 4).ToString() });
                results.Add(new ResultRow { Name = "13. Среднее число заявок в системе (L_сист), шт", Value = Math.Round(L_sist, 4).ToString() });
                results.Add(new ResultRow { Name = $"14. Среднее время пребывания заявки в СМО (W_сист), {units[data.LambdaUnitIndex]}", Value = Math.Round(W_sist_converted, 4).ToString() });
            }

            if (modelType == 1 && m > 0)
            {
                results.Add(new ResultRow { Name = "--- Предельные вероятности состояний очереди ---", Value = "" });
                for (int i = 1; i <= m; i++)
                {
                    double p_queue = data.IsMultiChannel || data.AutoCalculateChannels ? (Math.Pow(rho, n + i) / (Factorial(n) * Math.Pow(n, i))) * p0 : Math.Pow(rho, i + 1) * p0;
                    results.Add(new ResultRow { Name = $"   • Вероятность того, что в очереди ровно {i} заявка(и)", Value = Math.Round(p_queue, 4).ToString() });
                }
            }

            results.Add(new ResultRow { Name = "--- Производительность СМО ---", Value = "" });
            results.Add(new ResultRow { Name = "Число заявок, получивших отказ в течение часа, шт", Value = Math.Round(totalRejectedPerHour, 1).ToString() });
            results.Add(new ResultRow { Name = $"Номинальная производительность СМО ({n} каналов), {muUnit}", Value = Math.Round(nominalProd_converted, 4).ToString() });
            results.Add(new ResultRow { Name = "Фактическая производительность СМО, % от номинальной", Value = $"{Math.Round(factProdPercent, 2)}%" });

            DgResults.ItemsSource = results;
        }


        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = _mainWindow.SharedData;
                string[] modelNames = { "С отказами (потерями)", "С ограниченной длиной очереди", "С неограниченной очередью" };
                string[] units = { "минута", "час", "сутки" };

                var application = new Excel.Application();
                application.SheetsInNewWorkbook = 1; 
                Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
                worksheet.Name = "Отчет СМО";

                int startRowIndex = 1;

                Excel.Range headerInputsRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[2][startRowIndex]];
                headerInputsRange.Merge();
                headerInputsRange.Value = "ВХОДНЫЕ ПАРАМЕТРЫ МОДЕЛИ";
                headerInputsRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerInputsRange.Font.Bold = true;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Тип системы:";
                worksheet.Cells[2][startRowIndex] = data.IsMultiChannel ? "Многоканальная" : "Одноканальная";
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Модель очереди:";
                worksheet.Cells[2][startRowIndex] = modelNames[data.ModelTypeIndex];
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Поток заявок (λ):";
                worksheet.Cells[2][startRowIndex] = $"{data.LambdaRaw} заяв. в {units[data.LambdaUnitIndex]}";
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Интенсивность обслуживания (μ):";
                worksheet.Cells[2][startRowIndex] = $"{data.MuRaw} заяв. в {units[data.MuUnitIndex]}";
                startRowIndex++;

                if (data.IsMultiChannel)
                {
                    worksheet.Cells[1][startRowIndex] = "Заданное число каналов (n):";
                    worksheet.Cells[2][startRowIndex] = data.AutoCalculateChannels ? $"Автоподбор под P ≥ {data.TargetProbability}" : data.Channels.ToString();
                    startRowIndex++;
                }
                if (data.ModelTypeIndex == 1)
                {
                    worksheet.Cells[1][startRowIndex] = "Макс. длина очереди (m):";
                    worksheet.Cells[2][startRowIndex] = data.QueueLength;
                    startRowIndex++;
                }

                Excel.Range bordersInputs = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                bordersInputs.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                startRowIndex += 2; 

                Excel.Range headerResultsRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[2][startRowIndex]];
                headerResultsRange.Merge();
                headerResultsRange.Value = "РЕЗУЛЬТАТЫ И МЕТРИКИ ЭФФЕКТИВНОСТИ";
                headerResultsRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerResultsRange.Font.Bold = true;
                int resultsTableStartRow = startRowIndex;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Параметр эффективности СМО";
                worksheet.Cells[2][startRowIndex] = "Значение метрики";
                Excel.Range colHeaderRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[2][startRowIndex]];
                colHeaderRange.Font.Bold = true;
                colHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                startRowIndex++;

                var rows = DgResults.ItemsSource as List<ResultRow>;
                if (rows != null)
                {
                    foreach (ResultRow row in rows)
                    {
                        worksheet.Cells[1][startRowIndex] = row.Name;
                        worksheet.Cells[2][startRowIndex] = row.Value;

                        if (row.Name.StartsWith("---"))
                        {
                            Excel.Range sectionRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[2][startRowIndex]];
                            sectionRange.Merge();
                            sectionRange.Font.Italic = true;
                            sectionRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        else
                        {
                            (worksheet.Cells[2][startRowIndex] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        }
                        startRowIndex++;
                    }
                }

                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][resultsTableStartRow], worksheet.Cells[2][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                worksheet.Columns.AutoFit();

                application.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте в Excel: " + ex.Message, "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private TempMetrics CalculateMetricsForN(int n, double lambda, double mu, int m, int modelType, double rho)
        {
            var mtr = new TempMetrics();
            if (n == 1)
            {
                if (modelType == 0)
                {
                    mtr.P0 = 1.0 / (1.0 + rho); mtr.Potk = rho / (1.0 + rho); mtr.Q = 1.0 / (1.0 + rho); mtr.L_obs = rho / (1.0 + rho); mtr.P_channels.Add(mtr.Potk);
                }
                else if (modelType == 1)
                {
                    if (Math.Abs(rho - 1.0) < 0.0001)
                    {
                        mtr.P0 = 1.0 / (m + 2.0); mtr.Potk = mtr.P0; mtr.L_och = (m * (m + 1.0)) / (2.0 * (m + 2.0));
                    }
                    else
                    {
                        mtr.P0 = (1.0 - rho) / (1.0 - Math.Pow(rho, m + 2.0)); mtr.Potk = Math.Pow(rho, m + 1) * mtr.P0;
                        mtr.L_och = (Math.Pow(rho, 2) * (1.0 - Math.Pow(rho, m) * (m * (1.0 - rho) + 1.0))) / ((1.0 - rho) * (1.0 - Math.Pow(rho, m + 2.0)));
                    }
                    mtr.Q = 1.0 - mtr.Potk; mtr.L_obs = mtr.Q * rho; mtr.L_sist = mtr.L_och + mtr.L_obs; mtr.W_och = mtr.L_och / lambda; mtr.W_sist = mtr.L_sist / lambda;
                    mtr.P_channels.Add(rho * mtr.P0);
                }
                else
                {
                    if (rho >= 1.0) { mtr.P0 = 0; mtr.Potk = 1; mtr.Q = 0; mtr.L_sist = double.PositiveInfinity; }
                    else
                    {
                        mtr.P0 = 1.0 - rho; mtr.Potk = 0; mtr.Q = 1.0; mtr.L_och = Math.Pow(rho, 2) / (1.0 - rho);
                        mtr.L_sist = rho / (1.0 - rho); mtr.L_obs = rho; mtr.W_och = mtr.L_och / lambda; mtr.W_sist = mtr.L_sist / lambda;
                    }
                    mtr.P_channels.Add(rho * mtr.P0);
                }
            }
            else
            {
                if (modelType == 0)
                {
                    double sum = 0; for (int i = 0; i <= n; i++) sum += Math.Pow(rho, i) / Factorial(i);
                    mtr.P0 = 1.0 / sum; mtr.Potk = (Math.Pow(rho, n) / Factorial(n)) * mtr.P0; mtr.Q = 1.0 - mtr.Potk; mtr.L_obs = rho * mtr.Q;
                    for (int i = 1; i <= n; i++) mtr.P_channels.Add((Math.Pow(rho, i) / Factorial(i)) * mtr.P0);
                }
                else if (modelType == 1)
                {
                    double sum1 = 0; for (int i = 0; i <= n; i++) sum1 += Math.Pow(rho, i) / Factorial(i);
                    double sum2 = 0; for (int i = 1; i <= m; i++) sum2 += Math.Pow(rho, n + i) / (Factorial(n) * Math.Pow(n, i));
                    mtr.P0 = 1.0 / (sum1 + sum2); mtr.Potk = (Math.Pow(rho, n + m) / (Math.Pow(n, m) * Factorial(n))) * mtr.P0;
                    mtr.Q = 1.0 - mtr.Potk; mtr.L_obs = rho * mtr.Q;
                    for (int i = 1; i <= n; i++) mtr.P_channels.Add((Math.Pow(rho, i) / Factorial(i)) * mtr.P0);
                    double chi = rho / n;
                    if (Math.Abs(chi - 1.0) < 0.0001) mtr.L_och = (Math.Pow(rho, n) * m * (m + 1.0)) / (2.0 * Factorial(n)) * mtr.P0;
                    else mtr.L_och = (Math.Pow(rho, n + 1) * (1.0 - Math.Pow(chi, m) * (m * (1.0 - chi) + 1.0))) / (Factorial(n) * n * Math.Pow(1.0 - chi, 2)) * mtr.P0;
                    mtr.L_sist = mtr.L_och + mtr.L_obs; mtr.W_och = mtr.L_och / lambda; mtr.W_sist = mtr.L_sist / lambda;
                }
                else
                {
                    double chi = rho / n;
                    if (chi >= 1.0) { mtr.P0 = 0; mtr.Potk = 1; mtr.Q = 0; mtr.L_sist = double.PositiveInfinity; }
                    else
                    {
                        double sum = 0; for (int i = 0; i <= n - 1; i++) sum += Math.Pow(rho, i) / Factorial(i);
                        double lastTerm = Math.Pow(rho, n) / (Factorial(n) * (1.0 - chi));
                        mtr.P0 = 1.0 / (sum + lastTerm); mtr.Potk = 0; mtr.Q = 1.0; mtr.L_obs = rho;
                        for (int i = 1; i <= n; i++) mtr.P_channels.Add((Math.Pow(rho, i) / Factorial(i)) * mtr.P0);
                        mtr.L_och = (Math.Pow(rho, n + 1) / (n * Factorial(n) * Math.Pow(1.0 - chi, 2))) * mtr.P0;
                        mtr.L_sist = mtr.L_och + mtr.L_obs; mtr.W_och = mtr.L_och / lambda; mtr.W_sist = mtr.L_sist / lambda;
                    }
                }
            }
            mtr.A = lambda * mtr.Q;
            return mtr;
        }
        private double Factorial(int f)
        {
            if (f <= 1) return 1.0;
            double result = 1.0;
            for (int i = 2; i <= f; i++) result *= i;
            return result;
        }
        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            _mainWindow.MainFrame.Navigate(new Step1Page(_mainWindow));
        }
    }
}
