using System.Collections.Generic;
using System.IO;
using Курсовой_проект.Models;
using OfficeOpenXml;
using System.Windows;

namespace Курсовой_проект.Services
{
    public class ExcelExporter
    {
        public void ExportToExcel(List<ResultPoint> data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                worksheet.Cells[1, 1].Value = "x";
                worksheet.Cells[1, 2].Value = "y";
                worksheet.Cells[1, 3].Value = "f(x,y)";

                using (var range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                }

                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = data[i].X;
                    worksheet.Cells[i + 2, 2].Value = data[i].Y;
                    worksheet.Cells[i + 2, 3].Value = data[i].F;
                }

                worksheet.Cells[1, 1, data.Count + 1, 3].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}