using System.Collections.Generic;
using System.IO;
using Курсовой_проект.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Windows;

namespace Курсовой_проект.Services
{
    public class ExcelExporter
    {
        public void ExportToExcel(List<ResultPoint> data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                // Один лист для всего
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                // ========== ТАБЛИЦА ДАННЫХ ==========
                // Заголовки таблицы
                worksheet.Cells[1, 1].Value = "x";
                worksheet.Cells[1, 2].Value = "y";
                worksheet.Cells[1, 3].Value = "f(x,y)";

                using (var range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                }

                // Данные таблицы
                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = data[i].X;
                    worksheet.Cells[i + 2, 2].Value = data[i].Y;
                    worksheet.Cells[i + 2, 3].Value = data[i].F;
                }

                worksheet.Cells[1, 1, data.Count + 1, 3].AutoFitColumns();

                // ========== ГРАФИК ==========
                // Определяем диапазон данных для графика (столбцы x и y)
                int startRow = 1;
                int endRow = data.Count + 1;

                // Создаём график
                var chart = worksheet.Drawings.AddChart("SolutionGraph", eChartType.LineMarkers);
                chart.SetPosition(0, 0, 5, 0);
                chart.SetSize(800, 400);

                // Добавляем серию для y(x)
                var series = chart.Series.Add(
                    worksheet.Cells[startRow + 1, 2, endRow, 2],  // y значения
                    worksheet.Cells[startRow + 1, 1, endRow, 1]   // x значения
                );
                series.Header = "y(x) - численное решение";

                // Настройка внешнего вида графика
                chart.Title.Text = "График решения дифференциального уравнения";
                chart.XAxis.Title.Text = "x";
                chart.YAxis.Title.Text = "y(x)";

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}