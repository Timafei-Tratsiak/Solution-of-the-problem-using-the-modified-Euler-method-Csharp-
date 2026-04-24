using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Курсовой_проект.Models;

namespace Курсовой_проект.Services
{
    public class MatlabExportService
    {
        /// <summary>
        /// Экспорт в .m файл (скрипт MATLAB)
        /// </summary>
        public void ExportToMFile(List<ResultPoint> data, double x0, double y0, double X, double h, string function, string filePath)
        {
            var sb = new StringBuilder();

            sb.AppendLine("% MATLAB скрипт для визуализации результатов");
            sb.AppendLine("% Создано программой решения ДУ модифицированным методом Эйлера");
            sb.AppendLine("");
            sb.AppendLine("% Очистка рабочего пространства");
            sb.AppendLine("clear; clc; close all;");
            sb.AppendLine("");
            sb.AppendLine("% Параметры задачи");
            sb.AppendLine($"% Уравнение: dy/dx = {function}");
            sb.AppendLine($"% x0 = {x0}, y0 = {y0}, X = {X}, h = {h}");
            sb.AppendLine("");
            sb.AppendLine("% Данные");
            sb.AppendLine("x = [");

            foreach (var point in data)
            {
                sb.AppendLine($"    {point.X.ToString(System.Globalization.CultureInfo.InvariantCulture)};");
            }
            sb.AppendLine("];");
            sb.AppendLine("");
            sb.AppendLine("y = [");

            foreach (var point in data)
            {
                sb.AppendLine($"    {point.Y.ToString(System.Globalization.CultureInfo.InvariantCulture)};");
            }
            sb.AppendLine("];");
            sb.AppendLine("");
            sb.AppendLine("% Построение графика");
            sb.AppendLine("figure('Name', 'Решение ДУ модифицированным методом Эйлера', 'NumberTitle', 'off');");
            sb.AppendLine("plot(x, y, 'b-o', 'LineWidth', 2, 'MarkerSize', 4);");
            sb.AppendLine("grid on;");
            sb.AppendLine("xlabel('x');");
            sb.AppendLine("ylabel('y(x)');");
            sb.AppendLine($"title('Решение ДУ: dy/dx = {function.Replace("*", "\\cdot")}');");
            sb.AppendLine("legend('Численное решение (модиф. Эйлер)');");
            sb.AppendLine("");
            sb.AppendLine("% Сохранение графика в файл");
            sb.AppendLine("saveas(gcf, 'graph.png');");
            sb.AppendLine("");
            sb.AppendLine("disp('График построен и сохранён в graph.png');");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Экспорт в текстовый файл с данными
        /// </summary>
        public void ExportToDataFile(List<ResultPoint> data, string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("% x          y          f(x,y)");
            sb.AppendLine("% ------------------------------");

            foreach (var point in data)
            {
                sb.AppendLine($"{point.X.ToString(System.Globalization.CultureInfo.InvariantCulture),10} " +
                            $"{point.Y.ToString(System.Globalization.CultureInfo.InvariantCulture),10} " +
                            $"{point.F.ToString(System.Globalization.CultureInfo.InvariantCulture),10}");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}