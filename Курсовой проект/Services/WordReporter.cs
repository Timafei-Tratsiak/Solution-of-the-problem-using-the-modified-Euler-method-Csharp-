using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Курсовой_проект.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Курсовой_проект.Services
{
    public class WordReporter
    {
        public void CreateReport(List<ResultPoint> data, double x0, double y0, double X, double h, string function, string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Заголовок
                AddHeading(body, "Отчёт по курсовой работе", 1);
                AddHeading(body, "Решение дифференциального уравнения модифицированным методом Эйлера", 2);

                // Параметры задачи
                AddHeading(body, "1. Параметры задачи", 3);
                AddParagraph(body, $"Уравнение: dy/dx = {function}");
                AddParagraph(body, $"Начальные условия: x0 = {x0}, y0 = {y0}");
                AddParagraph(body, $"Отрезок интегрирования: [{x0}, {X}]");
                AddParagraph(body, $"Шаг интегрирования: h = {h}");
                AddParagraph(body, $"Количество точек: {data.Count}");

                // Метод решения
                AddHeading(body, "2. Метод решения", 3);
                AddParagraph(body, "Использован модифицированный метод Эйлера (метод Эйлера-Коши) второго порядка точности.");
                AddParagraph(body, "Формула метода: y_new = y + h * f(x + h/2, y + (h/2) * f(x, y))");

                // Результаты
                AddHeading(body, "3. Результаты расчёта", 3);
                AddTable(body, data);

                // Заключение
                AddHeading(body, "4. Заключение", 3);
                AddParagraph(body, $"Расчёт завершён успешно. Получено {data.Count} точек решения.");
            }
        }

        private void AddHeading(Body body, string text, int level)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));

            para.ParagraphProperties = new ParagraphProperties();
            para.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = $"Heading{level}" };

            run.RunProperties = new RunProperties();
            run.RunProperties.Bold = new Bold();
        }

        private void AddParagraph(Body body, string text)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));

            para.ParagraphProperties = new ParagraphProperties();
            para.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { After = "200" };
        }

        private void AddTable(Body body, List<ResultPoint> data)
        {
            Table table = new Table();

            TableProperties tableProps = new TableProperties();
            TableBorders borders = new TableBorders();
            borders.TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 1 };
            borders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 1 };
            borders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 1 };
            borders.RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 1 };
            borders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 1 };
            borders.InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 1 };
            tableProps.AppendChild(borders);
            table.AppendChild(tableProps);

            // Заголовки
            TableRow headerRow = new TableRow();
            AddCell(headerRow, "№");
            AddCell(headerRow, "x");
            AddCell(headerRow, "y");
            AddCell(headerRow, "f(x,y)");
            table.AppendChild(headerRow);

            // Данные (первые 15 строк и последние 5)
            int maxDisplay = 15;
            for (int i = 0; i < Math.Min(data.Count, maxDisplay); i++)
            {
                TableRow row = new TableRow();
                AddCell(row, (i + 1).ToString());
                AddCell(row, data[i].X.ToString("F6"));
                AddCell(row, data[i].Y.ToString("F6"));
                AddCell(row, data[i].F.ToString("F6"));
                table.AppendChild(row);
            }

            if (data.Count > maxDisplay)
            {
                TableRow dotsRow = new TableRow();
                AddCell(dotsRow, "...");
                AddCell(dotsRow, "...");
                AddCell(dotsRow, "...");
                AddCell(dotsRow, "...");
                table.AppendChild(dotsRow);

                for (int i = data.Count - 5; i < data.Count; i++)
                {
                    TableRow row = new TableRow();
                    AddCell(row, (i + 1).ToString());
                    AddCell(row, data[i].X.ToString("F6"));
                    AddCell(row, data[i].Y.ToString("F6"));
                    AddCell(row, data[i].F.ToString("F6"));
                    table.AppendChild(row);
                }
            }

            body.AppendChild(table);
        }

        private void AddCell(TableRow row, string text)
        {
            TableCell cell = new TableCell();
            Paragraph para = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text(text));
            para.AppendChild(run);
            cell.AppendChild(para);

            cell.TableCellProperties = new TableCellProperties();
            cell.TableCellProperties.TableCellWidth = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            row.AppendChild(cell);
        }
    }
}