using System;
using System.Diagnostics;
using System.Linq;

namespace ImageToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() != 2)
            {
                PrintUsage();
                return;
            }

            var converter = new ImageToExcelConverter();
            converter.Convert(args[0], args[1]);

        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: {0} image_filename output_excel_filename", Process.GetCurrentProcess().MainModule.FileName);
        }
    }
}
