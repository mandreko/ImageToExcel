using System;
using System.Diagnostics;
using System.Linq;

namespace ImageToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 2)
            {
                var converter = new ImageToExcelConverter();
                converter.Convert(args[0], args[1]);
            }
            else if (args.Count() == 3)
            {
                int resizePercentage;
                try
                {
                    resizePercentage = Convert.ToInt32(args[2]);
                }
                catch (Exception)
                {
                    Console.WriteLine("Not a valid resize percentage");   
                    PrintUsage();
                    return;
                }

                var converter = new ImageToExcelConverter();
                converter.Convert(args[0], args[1], resizePercentage);
            }
            else
            {
                PrintUsage();
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: {0} image_filename output_excel_filename [resize_percentage]", Process.GetCurrentProcess().MainModule.FileName);
        }
    }
}
