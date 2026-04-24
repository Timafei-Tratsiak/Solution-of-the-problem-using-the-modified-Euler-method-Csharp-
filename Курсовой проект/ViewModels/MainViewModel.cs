using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Курсовой_проект.Models;
using Курсовой_проект.Services;
using System.Diagnostics;
using System.IO;

namespace Курсовой_проект.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private OdeSolver _solver;

        private string _function = "y - x*x + 1";
        private double _x0 = 0;
        private double _y0 = 0.5;
        private double _x = 2;
        private double _h = 0.1;
        private List<ResultPoint> _results;
        private string _status = "Готов";
        private PlotModel _plotModel;

        public MainViewModel()
        {
            _solver = new OdeSolver();

            CalculateCommand = new RelayCommand(_ => ExecuteCalculate());
            ExportToExcelCommand = new RelayCommand(_ => ExecuteExportToExcel());
            CreateWordReportCommand = new RelayCommand(_ => ExecuteCreateWordReport());
            ExportToMatlabCommand = new RelayCommand(_ => ExecuteExportToMatlab());
            IntegrateWithMathcadCommand = new RelayCommand(_ => ExecuteIntegrateWithMathcad());
            Show3DInMatlabCommand = new RelayCommand(_ => ExecuteShow3DInMatlab());
        }

        public string Function
        {
            get => _function;
            set { _function = value; OnPropertyChanged(); }
        }

        public double X0
        {
            get => _x0;
            set { _x0 = value; OnPropertyChanged(); }
        }

        public double Y0
        {
            get => _y0;
            set { _y0 = value; OnPropertyChanged(); }
        }

        public double X
        {
            get => _x;
            set { _x = value; OnPropertyChanged(); }
        }

        public double H
        {
            get => _h;
            set { _h = value; OnPropertyChanged(); }
        }

        public List<ResultPoint> Results
        {
            get => _results;
            set { _results = value; OnPropertyChanged(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public PlotModel PlotModel
        {
            get => _plotModel;
            set { _plotModel = value; OnPropertyChanged(); }
        }

        public ICommand CalculateCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand CreateWordReportCommand { get; }
        public ICommand ExportToMatlabCommand { get; }
        public ICommand IntegrateWithMathcadCommand { get; }
        public ICommand Show3DInMatlabCommand { get; }

        private void ExecuteCalculate()
        {
            try
            {
                var parser = new ExpressionParser(Function);

                if (!parser.ValidateExpression(out string error))
                {
                    Status = $"Ошибка в формуле: {error}";
                    MessageBox.Show($"Ошибка в формуле: {error}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var f = parser.ParseFunction();
                Results = _solver.SolveModifiedEuler(f, X0, Y0, X, H);
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

        private void BuildGraph()
        {
            if (Results == null || Results.Count == 0) return;

            var plotModel = new PlotModel
            {
                Title = "Решение дифференциального уравнения",
                Subtitle = $"Модифицированный метод Эйлера, h = {H}"
            };

            plotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "x",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            plotModel.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "y(x)",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            var series = new LineSeries
            {
                Title = "y(x) - численное решение",
                Color = OxyColors.Blue,
                StrokeThickness = 2,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = OxyColors.Blue,
                LineStyle = LineStyle.Solid
            };

            foreach (var point in Results)
            {
                series.Points.Add(new DataPoint(point.X, point.Y));
            }

            plotModel.Series.Add(series);
            PlotModel = plotModel;
        }

        // 3D график в MATLAB
        private void ExecuteShow3DInMatlab()
        {
            try
            {
                if (Results == null || Results.Count == 0)
                {
                    MessageBox.Show("Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "MATLAB script (*.m)|*.m",
                    DefaultExt = ".m",
                    FileName = $"matlab_3d_visualization_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    Create3DMatlabScript(saveDialog.FileName);

                    MessageBox.Show(
                        $"3D MATLAB скрипт создан:\n{saveDialog.FileName}\n\n" +
                        $"Для просмотра 3D графика:\n" +
                        $"1. Откройте файл в MATLAB\n" +
                        $"2. Запустите скрипт (F5)\n" +
                        $"3. Вращайте график мышью для просмотра со всех сторон\n\n" +
                        $"Скрипт построит 3D линию решения в пространстве (X, Y, Шаг)",
                        "3D Визуализация в MATLAB",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    Status = $"3D скрипт для MATLAB создан: {Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Create3DMatlabScript(string filePath)
        {
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("% 3D Визуализация решения дифференциального уравнения");
            sb.AppendLine("% Создано программой решения ДУ модифицированным методом Эйлера");
            sb.AppendLine($"% Дата: {DateTime.Now}");
            sb.AppendLine($"% Уравнение: dy/dx = {Function}");
            sb.AppendLine($"% x0 = {X0}, y0 = {Y0}, X = {X}, h = {H}");
            sb.AppendLine();
            sb.AppendLine("% Данные");
            sb.AppendLine("x = [");
            foreach (var point in Results)
            {
                sb.AppendLine($"    {point.X.ToString(System.Globalization.CultureInfo.InvariantCulture)};");
            }
            sb.AppendLine("];");
            sb.AppendLine();
            sb.AppendLine("y = [");
            foreach (var point in Results)
            {
                sb.AppendLine($"    {point.Y.ToString(System.Globalization.CultureInfo.InvariantCulture)};");
            }
            sb.AppendLine("];");
            sb.AppendLine();
            sb.AppendLine("% Параметр t (шаги)");
            sb.AppendLine($"t = 0:{Results.Count - 1};");
            sb.AppendLine();
            sb.AppendLine("% 3D график");
            sb.AppendLine("figure('Name', '3D Визуализация решения ДУ', 'NumberTitle', 'off');");
            sb.AppendLine("plot3(x, y, t, 'b-o', 'LineWidth', 2, 'MarkerSize', 4);");
            sb.AppendLine("grid on;");
            sb.AppendLine("xlabel('x');");
            sb.AppendLine("ylabel('y(x)');");
            sb.AppendLine("zlabel('Step number');");
            sb.AppendLine($"title('3D Визуализация решения ДУ: dy/dx = {Function.Replace("*", ".*")}');");
            sb.AppendLine("legend('Численное решение (модиф. Эйлер)');");
            sb.AppendLine();
            sb.AppendLine("% Настройка 3D вида");
            sb.AppendLine("view(45, 30);");
            sb.AppendLine("rotate3d on;");
            sb.AppendLine();
            sb.AppendLine("% Сохранение графика в файл");
            sb.AppendLine("saveas(gcf, 'graph_3d.png');");
            sb.AppendLine();
            sb.AppendLine("disp('3D график построен и сохранён в graph_3d.png');");
            sb.AppendLine("disp('Используйте мышь для вращения графика');");

            System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
        }

        private void ExecuteExportToExcel()
        {
            try
            {
                if (Results == null || Results.Count == 0)
                {
                    MessageBox.Show("Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

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

                    MessageBox.Show($"Данные успешно экспортированы в файл:\n{saveDialog.FileName}",
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
                if (Results == null || Results.Count == 0)
                {
                    MessageBox.Show("Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var reporter = new WordReporter();
                    reporter.CreateReport(Results, X0, Y0, X, H, Function, saveDialog.FileName);

                    MessageBox.Show($"Отчёт успешно создан:\n{saveDialog.FileName}",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    Status = $"Отчёт в Word создан: {System.IO.Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании отчёта: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExecuteExportToMatlab()
        {
            try
            {
                if (Results == null || Results.Count == 0)
                {
                    MessageBox.Show("Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "MATLAB script (*.m)|*.m|Text file (*.txt)|*.txt",
                    DefaultExt = ".m",
                    FileName = $"matlab_script_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var exporter = new MatlabExportService();

                    if (saveDialog.FileName.EndsWith(".m"))
                    {
                        exporter.ExportToMFile(Results, X0, Y0, X, H, Function, saveDialog.FileName);
                        MessageBox.Show($"MATLAB скрипт создан:\n{saveDialog.FileName}\n\n" +
                            "Для использования:\n1. Откройте файл в MATLAB\n2. Запустите скрипт",
                            "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        exporter.ExportToDataFile(Results, saveDialog.FileName);
                        MessageBox.Show($"Файл с данными создан:\n{saveDialog.FileName}",
                            "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    Status = $"Экспорт в MATLAB завершён: {System.IO.Path.GetFileName(saveDialog.FileName)}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в MATLAB: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExecuteIntegrateWithMathcad()
        {
            try
            {
                if (Results == null || Results.Count == 0)
                {
                    MessageBox.Show("Сначала выполните расчёт!", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

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
                mathcadService.SetStringValue("func_str", functionWithQuotes);

                Status = $"Данные переданы в Mathcad Prime 10";

                MessageBox.Show(
                    $"Mathcad Prime 10 запущен!\n\n" +
                    $"Переданы параметры:\n" +
                    $"• x0 = {X0}\n" +
                    $"• y0 = {Y0}\n" +
                    $"• h = {H}\n" +
                    $"• X = {X}\n" +
                    $"• Функция: {functionWithQuotes}\n\n" +
                    $"Для построения графика:\n" +
                    $"1. Найдите переменную func_str\n" +
                    $"2. Скопируйте её значение без кавычек\n" +
                    $"3. Вставьте в определение функции f(x,y):= ...\n" +
                    $"4. Нажмите Enter для пересчёта",
                    "Интеграция с Mathcad",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}