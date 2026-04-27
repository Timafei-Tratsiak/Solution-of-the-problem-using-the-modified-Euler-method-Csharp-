using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Курсовой_проект.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Курсовой_проект.Services
{
    public class WordReporter
    {
        public void CreateReport(List<ResultPoint> data, double x0, double y0, double X, double h, string function, string filePath, string imagePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();
                mainPart.Document.Append(body);

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

                // Результаты расчёта
                AddHeading(body, "3. Результаты расчёта", 3);
                AddTable(body, data);

                // Заключение
                AddHeading(body, "4. Заключение", 3);
                AddParagraph(body, $"Расчёт завершён успешно. Получено {data.Count} точек решения.");
                AddParagraph(body, $"Конечное значение: y({X}) = {data.Last().Y:F6}");
                AddParagraph(body, "");

                // График решения
                AddHeading(body, "5. График решения", 3);

                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                {
                    AddImageToBody(mainPart, body, imagePath);
                }
                else
                {
                    AddParagraph(body, "График решения представлен в основном окне программы.");
                }

                mainPart.Document.Save();
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

            TableProperties tblProp = new TableProperties();
            TableBorders borders = new TableBorders
            {
                TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 1 },
                BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 1 },
                LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 1 },
                RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 1 },
                InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 1 },
                InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 1 }
            };
            tblProp.AppendChild(borders);
            table.AppendChild(tblProp);

            // Заголовки
            TableRow headerRow = new TableRow();
            AddCell(headerRow, "№", true);
            AddCell(headerRow, "x", true);
            AddCell(headerRow, "y", true);
            AddCell(headerRow, "f(x,y)", true);
            table.AppendChild(headerRow);

            // Данные (первые 15 строк и последние 5)
            int maxDisplay = 15;
            for (int i = 0; i < Math.Min(data.Count, maxDisplay); i++)
            {
                TableRow row = new TableRow();
                AddCell(row, (i + 1).ToString(), false);
                AddCell(row, data[i].X.ToString("F6"), false);
                AddCell(row, data[i].Y.ToString("F6"), false);
                AddCell(row, data[i].F.ToString("F6"), false);
                table.AppendChild(row);
            }

            if (data.Count > maxDisplay)
            {
                TableRow dotsRow = new TableRow();
                AddCell(dotsRow, "...", false);
                AddCell(dotsRow, "...", false);
                AddCell(dotsRow, "...", false);
                AddCell(dotsRow, "...", false);
                table.AppendChild(dotsRow);

                for (int i = data.Count - 5; i < data.Count; i++)
                {
                    TableRow row = new TableRow();
                    AddCell(row, (i + 1).ToString(), false);
                    AddCell(row, data[i].X.ToString("F6"), false);
                    AddCell(row, data[i].Y.ToString("F6"), false);
                    AddCell(row, data[i].F.ToString("F6"), false);
                    table.AppendChild(row);
                }
            }

            body.AppendChild(table);
        }

        private void AddCell(TableRow row, string text, bool isHeader)
        {
            TableCell cell = new TableCell();
            Paragraph para = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text(text));
            para.AppendChild(run);

            if (isHeader)
            {
                run.RunProperties = new RunProperties();
                run.RunProperties.Bold = new Bold();
            }

            cell.AppendChild(para);
            cell.TableCellProperties = new TableCellProperties();
            cell.TableCellProperties.TableCellWidth = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            row.AppendChild(cell);
        }

        private void AddImageToBody(MainDocumentPart mainPart, Body body, string imagePath)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = 5943600L, Cy = 3600000L },
                    new DW.DocProperties() { Id = 1U, Name = "Solution Graph" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Graph.png" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = 5943600L, Cy = 3600000L }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );

            body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}