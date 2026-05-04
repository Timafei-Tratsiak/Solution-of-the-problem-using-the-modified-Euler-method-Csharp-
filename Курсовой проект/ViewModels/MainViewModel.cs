using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Курсовой_проект.Models;
using Курсовой_проект.Services;
using AboutLibra;

namespace Курсовой_проект.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly OdeSolver _solver;
        private readonly MatlabIntegrationService _matlabService;
        private readonly SettingsManager _settingsManager;

        // Поля
        private string _function = "y - x*x + 1";
        private double _x0 = 0;
        private double _y0 = 0.5;
        private double _x = 2;
        private double _h = 0.1;
        private List<ResultPoint> _results;
        private string _status = "Готов";
        private PlotModel _plotModel;

        // Строковые поля для привязки TextBox (работа с точкой/запятой)
        private string _x0Text = "0";
        private string _y0Text = "0.5";
        private string _xText = "2";
        private string _hText = "0.1";

        // Параметры для проверки изменений
        private string _lastFunction;
        private double _lastX0, _lastY0, _lastX, _lastH;

        public MainViewModel()
        {
            _solver = new OdeSolver();
            _matlabService = new MatlabIntegrationService();
            _settingsManager = new SettingsManager();

            // Инициализация команд
            CalculateCommand = new RelayCommand(_ => ExecuteCalculate());
            ExportToExcelCommand = new RelayCommand(_ => ExecuteExportToExcel());
            CreateWordReportCommand = new RelayCommand(_ => ExecuteCreateWordReport());
            ExportToMatlabCommand = new RelayCommand(_ => ExecuteIntegrateWithMatlab());
            IntegrateWithMathcadCommand = new RelayCommand(_ => ExecuteIntegrateWithMathcad());
            SaveGraphAsPngCommand = new RelayCommand(_ => ExecuteSaveGraphAsPng());
            ShowAboutCommand = new RelayCommand(_ => ExecuteShowAbout());
            ExitCommand = new RelayCommand(_ => ExecuteExit());

            // Загружаем сохранённые параметры
            LoadSavedSettings();

            ExecuteCalculate();
        }

        // Строковые свойства для привязки (работают с точкой и запятой)
        public string X0Text
        {
            get => _x0Text;
            set
            {
                _x0Text = value;
                OnPropertyChanged();
                if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double result))
                {
                    X0 = result;
                }
            }
        }

        public string Y0Text
        {
            get => _y0Text;
            set
            {
                _y0Text = value;
                OnPropertyChanged();
                if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double result))
                {
                    Y0 = result;
                }
            }
        }

        public string XText
        {
            get => _xText;
            set
            {
                _xText = value;
                OnPropertyChanged();
                if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double result))
                {
                    X = result;
                }
            }
        }

        public string HText
        {
            get => _hText;
            set
            {
                _hText = value;
                OnPropertyChanged();
                if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double result))
                {
                    H = result;
                }
            }
        }

        // Загрузка параметров из INI файла
        private void LoadSavedSettings()
        {
            try
            {
                var (function, x0, y0, X, h) = _settingsManager.LoadSettings();
                _function = function;
                _x0 = x0;
                _y0 = y0;
                _x = X;
                _h = h;

                // Обновляем строковые свойства для отображения
                X0Text = x0.ToString(System.Globalization.CultureInfo.InvariantCulture);
                Y0Text = y0.ToString(System.Globalization.CultureInfo.InvariantCulture);
                XText = X.ToString(System.Globalization.CultureInfo.InvariantCulture);
                HText = h.ToString(System.Globalization.CultureInfo.InvariantCulture);

                OnPropertyChanged(nameof(Function));

                Status = "Параметры загружены из настроек";
            }
            catch (Exception ex)
            {
                Status = $"Ошибка загрузки настроек: {ex.Message}";
            }
        }

        // Сохранение параметров в INI файл
        private void SaveCurrentSettings()
        {
            try
            {
                _settingsManager.SaveSettings(Function, X0, Y0, X, H);
            }
            catch (Exception ex)
            {
                Status = $"Ошибка сохранения настроек: {ex.Message}";
            }
        }

        // Свойства для привязки
        public string Function { get => _function; set { _function = value; OnPropertyChanged(); } }
        public double X0 { get => _x0; set { _x0 = value; OnPropertyChanged(); } }
        public double Y0 { get => _y0; set { _y0 = value; OnPropertyChanged(); } }
        public double X { get => _x; set { _x = value; OnPropertyChanged(); } }
        public double H { get => _h; set { _h = value; OnPropertyChanged(); } }
        public List<ResultPoint> Results { get => _results; set { _results = value; OnPropertyChanged(); } }
        public string Status { get => _status; set { _status = value; OnPropertyChanged(); } }
        public PlotModel PlotModel { get => _plotModel; set { _plotModel = value; OnPropertyChanged(); } }

        // Команды
        public ICommand CalculateCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand CreateWordReportCommand { get; }
        public ICommand ExportToMatlabCommand { get; }
        public ICommand IntegrateWithMathcadCommand { get; }
        public ICommand SaveGraphAsPngCommand { get; }
        public ICommand ShowAboutCommand { get; }
        public ICommand ExitCommand { get; }

        // Логика расчёта
        private void ExecuteCalculate()
        {
            try
            {
                var parser = new ExpressionParser(Function);
                if (!parser.ValidateExpression(out string error))
                {
                    Status = $"Ошибка формулы: {error}";
                    MessageBox.Show($"Ошибка в формуле: {error}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var f = parser.ParseFunction();
                Results = _solver.SolveModifiedEuler(f, X0, Y0, X, H);

                _lastFunction = Function;
                _lastX0 = X0;
                _lastY0 = Y0;
                _lastX = X;
                _lastH = H;

                SaveCurrentSettings();

                Status = $"Расчёт завершён. Получено {Results.Count} точек. x от {X0} до {X}, шаг {H}";
                BuildGraph();
            }
            catch (Exception ex)
            {
                Status = $"Ошибка расчёта: {ex.Message}";
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Построение графика
        private void BuildGraph()
        {
            if (Results == null) return;

            var model = new PlotModel
            {
                Title = "Решение дифференциального уравнения",
                Subtitle = $"Модифицированный метод Эйлера, h = {H}",
                Background = OxyColors.White
            };

            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "x" });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "y(x)" });

            var series = new LineSeries
            {
                Title = "y(x) - численное решение",
                Color = OxyColors.Blue,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3
            };
            foreach (var p in Results) series.Points.Add(new DataPoint(p.X, p.Y));
            model.Series.Add(series);

            PlotModel = model;
        }

        // Проверка и выполнение расчёта если нужно
        private void EnsureCalculated()
        {
            bool changed = _lastFunction != Function ||
                           Math.Abs(_lastX0 - X0) > 1e-10 ||
                           Math.Abs(_lastY0 - Y0) > 1e-10 ||
                           Math.Abs(_lastX - X) > 1e-10 ||
                           Math.Abs(_lastH - H) > 1e-10;
            if (Results == null || Results.Count == 0 || changed) ExecuteCalculate();
        }

        // Экспорт в Excel
        private void ExecuteExportToExcel()
        {
            try
            {
                EnsureCalculated();

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    FileName = $"Результаты_расчёта_{DateTime.Now:dd_MM_yyyy_HH/mm}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var exporter = new ExcelExporter();
                    exporter.ExportToExcel(Results, saveDialog.FileName);

                    MessageBox.Show($"Файл Excel создан:\n{saveDialog.FileName}\n\n" +
                        "Содержимое:\n• Таблица с результатами (x, y, f(x,y))\n• График решения y(x)",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    Status = $"Экспорт в Excel завершён: {System.IO.Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Отчёт в Word
        private void ExecuteCreateWordReport()
        {
            try
            {
                EnsureCalculated();

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = $"Отчёт_{DateTime.Now:dd_MM_yyyy_HH/mm}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    string imagePath = ExportGraphToPng();
                    var reporter = new WordReporter();
                    reporter.CreateReport(Results, X0, Y0, X, H, Function, saveDialog.FileName, imagePath);
                    if (File.Exists(imagePath)) File.Delete(imagePath);

                    MessageBox.Show($"Отчёт Word создан:\n{saveDialog.FileName}", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    Status = $"Отчёт в Word создан: {System.IO.Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании отчёта: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Интеграция с MATLAB (ode45)
        private void ExecuteIntegrateWithMatlab()
        {
            try
            {
                EnsureCalculated();
                Status = "Запуск MATLAB с методом ode45...";
                _matlabService.SolveWithOde45(X0, Y0, X, H, Function);
                Status = "MATLAB ode45 успешно выполнен. График построен в MATLAB.";
            }
            catch (Exception ex)
            {
                Status = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Ошибка при вызове MATLAB: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Интеграция с Mathcad
        private void ExecuteIntegrateWithMathcad()
        {
            try
            {
                EnsureCalculated();

                var mathcadService = new MathcadIntegrationService();

                if (!mathcadService.StartMathcad())
                {
                    MessageBox.Show("Не удалось запустить Mathcad Prime 10", "Ошибка");
                    return;
                }

                Status = "Передача данных в Mathcad...";

                mathcadService.SetRealValue("x0_input", X0);
                mathcadService.SetRealValue("y0_input", Y0);
                mathcadService.SetRealValue("h_input", H);
                mathcadService.SetRealValue("X_input", X);

                string functionWithQuotes = $"\"{Function}\"";
                mathcadService.SetStringValue("func_def", functionWithQuotes);

                Status = $"Данные переданы в Mathcad Prime 10";

                MessageBox.Show(
                    $"Mathcad Prime 10 запущен!\n\nПереданы параметры:\n" +
                    $"• x0 = {X0}\n• y0 = {Y0}\n• h = {H}\n• X = {X}\n• Функция: {functionWithQuotes}\n\n" +
                    $"В Mathcad: найдите переменную func_def и перепишите её без кавычек в f(x,y):=",
                    "Интеграция с Mathcad", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        // Сохранение графика в PNG
        private void ExecuteSaveGraphAsPng()
        {
            try
            {
                if (PlotModel == null)
                {
                    MessageBox.Show("График не построен. Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "PNG files (*.png)|*.png",
                    DefaultExt = ".png",
                    FileName = $"graph_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    using (var stream = File.Create(saveDialog.FileName))
                    {
                        var exporter = new OxyPlot.WindowsForms.PngExporter { Width = 1200, Height = 800 };
                        exporter.Export(PlotModel, stream);
                    }
                    MessageBox.Show($"График сохранён:\n{saveDialog.FileName}", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    Status = $"График сохранён: {Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения графика: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void SaveSettingsOnExit()
        {
            SaveCurrentSettings();
        }

        // О программе (из DLL)
        private void ExecuteShowAbout()
        {
            var aboutWindow = new AboutLibra.AboutWindow();
            aboutWindow.Owner = Application.Current.MainWindow;
            aboutWindow.ShowDialog();
        }

        // Выход из приложения
        private void ExecuteExit()
        {
            SaveCurrentSettings();
            Application.Current.Shutdown();
        }

        // Экспорт графика во временный PNG
        private string ExportGraphToPng()
        {
            if (PlotModel == null) return null;
            string tempFile = Path.Combine(Path.GetTempPath(), $"graph_{Guid.NewGuid()}.png");
            var exporter = new OxyPlot.WindowsForms.PngExporter { Width = 800, Height = 600 };
            using (var stream = File.Create(tempFile)) exporter.Export(PlotModel, stream);
            return tempFile;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}