using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ImageToExcel
{
    public class ImageToExcelConverter
    {
        public void Convert(string imagePath, string excelPath)
        {
            if (imagePath == null) throw new ArgumentNullException("imagePath");
            if (excelPath == null) throw new ArgumentNullException("excelPath");

            if (!File.Exists(imagePath)) throw new ArgumentException("Image path does not exist", "imagePath");

            Bitmap bitmap = new Bitmap(imagePath);

            FileInfo newFile = new FileInfo(excelPath);
            ExcelPackage pck = new ExcelPackage(newFile);

            var ws = pck.Workbook.Worksheets.Add("Image");
            ws.View.ShowGridLines = false;

            for (int x = 0; x < bitmap.Width; x++)
            {
                for (int y = 0; y < bitmap.Height; y++)
                {
                    Color pixelColor = bitmap.GetPixel(x, y);
                    string cell = string.Format("{0}{1}", GetExcelColumnName(x+1), y+1);

                    ws.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[cell].Style.Fill.BackgroundColor.SetColor(pixelColor);
                }
            }
            
            ws.View.ZoomScale = 10;

            for (int x = 1; x < bitmap.Width; x++)
            {
                ws.Column(x).Width = 2;
            }
            for (int y = 1; y < bitmap.Height; y++)
            {
                ws.Row(y).Height = 5;
            }

            pck.Save();

        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = System.Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}