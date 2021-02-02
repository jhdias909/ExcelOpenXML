using System;
using System.Diagnostics;
using System.IO;

namespace ExcelOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = new Stopwatch();

            if (File.Exists(@"c:\temp\teste.xlsx"))
                File.Delete(@"c:\temp\teste.xlsx");

            sw.Start();
            try
            {
                var obj = new ExcelOpenXML.GenerateExcel.GenerateExcel();
                obj.CreatePackage(@"c:\temp\teste.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                throw;
            }
            

            sw.Stop();

            Console.WriteLine($"Tempo de processamento: {sw.ElapsedMilliseconds}");
            Console.ReadKey();
        }
    }
}
