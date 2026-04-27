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

namespace Курсовой_проект.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly OdeSolver _solver;
        private readonly MatlabIntegrationService _matlabService;

        // Поля
        private string _function = "y - x*x + 1";
        private double _x0 = 0;
        private double _y0 = 0.5;
        private double _x = 2;
        private double _h = 0.1;
        private List<ResultPoint> _results;
        private string _status = "Готов";
        private PlotModel _plotModel;

        // Параметры для проверки изменений
        private string _lastFunction;
        private double _lastX0, _lastY0, _lastX, _lastH;

        public MainViewModel()
        {
            _solver = new OdeSolver();
            _matlabService = new MatlabIntegrationService();

            // Инициализация команд
            CalculateCommand = new RelayCommand(_ => ExecuteCalculate());
            ExportToExcelCommand = new RelayCommand(_ => ExecuteExportToExcel());
            CreateWordReportCommand = new RelayCommand(_ => ExecuteCreateWordReport());
            ExportToMatlabCommand = new RelayCommand(_ => ExecuteIntegrateWithMatlab());
            IntegrateWithMathcadCommand = new RelayCommand(_ => ExecuteIntegrateWithMathcad());

            ExecuteCalculate();
        }

        // Свойства для привязки (Binding)
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

        private void ExecuteCalculate()
        {
            try
            {
                var parser = new ExpressionParser(Function);
                if (!parser.ValidateExpression(out string error))
                {
                    Status = $"Ошибка формулы: {error}";
                    return;
                }

                var f = parser.ParseFunction();
                Results = _solver.SolveModifiedEuler(f, X0, Y0, X, H);

                // Сохраняем состояние последнего успешного расчета
                _lastFunction = Function; _lastX0 = X0; _lastY0 = Y0; _lastX = X; _lastH = H;

                Status = $"Расчёт завершён. Получено {Results.Count} точек. x от {X0} до {X}, шаг {H}";
                BuildGraph();
            }
            catch (Exception ex) { Status = "Ошибка: " + ex.Message; }
        }

        private void BuildGraph()
        {
            if (Results == null) return;
            var model = new PlotModel { Title = "Решение дифференциального уравнения", Subtitle = $"Модифицированный метод Эйлера, h = {H}" };

            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "x" });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "y(x)" });

            var series = new LineSeries { Title = "y(x) - численное решение", Color = OxyColors.Blue, MarkerType = MarkerType.Circle, MarkerSize = 3 };
            foreach (var p in Results) series.Points.Add(new DataPoint(p.X, p.Y));
            model.Series.Add(series);

            PlotModel = model;
        }

        // --- МЕТОДЫ ЭКСПОРТА И ИНТЕГРАЦИИ ---

        private void ExecuteExportToExcel()
        {
            try
            {
                EnsureCalculated();

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    FileName = $"Результаты_расчёта_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var exporter = new ExcelExporter();
                    exporter.ExportToExcel(Results, saveDialog.FileName);

                    MessageBox.Show($"Файл Excel создан:\n{saveDialog.FileName}\n\n" +
                        "Содержимое:\n" +
                        "• Таблица с результатами (x, y, f(x,y))\n" +
                        "• График решения y(x)",
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

        private void ExecuteCreateWordReport()
        {
            try
            {
                EnsureCalculated();
                var saveDialog = new SaveFileDialog { Filter = "Word documents (*.docx)|*.docx", FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmmss}" };
                if (saveDialog.ShowDialog() == true)
                {
                    string imgPath = ExportGraphToPng();
                    var reporter = new WordReporter();
                    reporter.CreateReport(Results, X0, Y0, X, H, Function, saveDialog.FileName, imgPath);
                    if (File.Exists(imgPath)) File.Delete(imgPath);
                    MessageBox.Show("Отчёт Word создан!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Word: " + ex.Message); }
        }

        private void ExecuteIntegrateWithMatlab()
        {
            try
            {
                EnsureCalculated();
                Status = "Запуск MATLAB...";
                _matlabService.SolveWithOde45(X0, Y0, X, H, Function);
                Status = "MATLAB: расчет завершен";
            }
            catch (Exception ex) { MessageBox.Show("Ошибка MATLAB: " + ex.Message); }
        }

        private void ExecuteIntegrateWithMathcad()
        {
            try
            {
                EnsureCalculated();
                var mathcadService = new MathcadIntegrationService();
                if (!mathcadService.StartMathcad())
                {
                    MessageBox.Show("Mathcad Prime не найден.");
                    return;
                }
                mathcadService.SetRealValue("x0_input", X0);
                mathcadService.SetRealValue("y0_input", Y0);
                mathcadService.SetRealValue("h_input", H);
                mathcadService.SetRealValue("X_input", X);
                mathcadService.SetStringValue("func_def", $"\"{Function}\"");
                MessageBox.Show("Параметры переданы в Mathcad Prime.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка Mathcad: " + ex.Message); }
        }

        // --- ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ---

        private string ExportGraphToPng()
        {
            if (PlotModel == null) return null;
            string tempFile = Path.Combine(Path.GetTempPath(), $"graph_{Guid.NewGuid()}.png");
            var exporter = new OxyPlot.WindowsForms.PngExporter { Width = 800, Height = 600};
            using (var stream = File.Create(tempFile)) { exporter.Export(PlotModel, stream); }
            return tempFile;
        }

        private void EnsureCalculated()
        {
            bool changed = _lastFunction != Function ||
                           Math.Abs(_lastX0 - X0) > 1e-10 ||
                           Math.Abs(_lastY0 - Y0) > 1e-10 ||
                           Math.Abs(_lastX - X) > 1e-10 ||
                           Math.Abs(_lastH - H) > 1e-10;
            if (Results == null || Results.Count == 0 || changed) ExecuteCalculate();
        }

        // Реализация INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}