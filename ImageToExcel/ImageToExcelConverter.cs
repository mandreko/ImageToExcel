using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ImageToExcel
{
    public class ImageToExcelConverter
    {
        private const decimal DefaultResizeMultiplier = 1m;
        private const int DefaultZoomScale = 10;

        public void Convert(string imagePath, string excelPath, int? resizePercentage = null)
        {
            if (imagePath == null) throw new ArgumentNullException("imagePath");
            if (excelPath == null) throw new ArgumentNullException("excelPath");

            if (!File.Exists(imagePath)) throw new ArgumentException("Image path does not exist", "imagePath");
            if (File.Exists(excelPath)) throw new ArgumentException("Excel path already exists", "imagePath");

            Bitmap bitmap = new Bitmap(imagePath);

            // At least reduce the size by the default size. If user specified more, keep going based on the default
            bitmap = ResizeBitMap(bitmap, resizePercentage);

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

            //ws.View.ZoomScale = DefaultZoomScale;

            // Notes on excel units:
            // 9 excel units = 
            // 68px wide
            // 12px tall
               
            // 1 excel unit = 
            // 7.555555555555556 wide
            // .75 tall
               
            // 1px
            // 0.75 tall
            // 0.08 wide
               
            // 100px
            // 75 tall
            // 13.57 wide

            for (int x = 1; x <= bitmap.Width; x++)
            {
                ws.Column(x).Width = .42; // 5px
            }
            for (int y = 1; y <= bitmap.Height; y++)
            {
                ws.Row(y).Height = 3.75; // 5px
            }

            pck.Save();

        }

        private Bitmap ResizeBitMap(Bitmap bitmap, int? resizePercentage)
        {
            // Modify the size based on the default
            decimal newHeightDecimal = bitmap.Height * DefaultResizeMultiplier;
            decimal newWidthDecimal = (bitmap.Width * DefaultResizeMultiplier);

            if (resizePercentage.HasValue)
            {
                newHeightDecimal = newHeightDecimal * (((decimal)resizePercentage.Value) / 100);
                newWidthDecimal = newWidthDecimal * (((decimal)resizePercentage.Value) / 100);
            }

            int newHeight = (int)Math.Floor(newHeightDecimal);
            int newWidth = (int)Math.Floor(newWidthDecimal);

            Bitmap result = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(bitmap, 0, 0, newWidth, newHeight);
            }

            return result;
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