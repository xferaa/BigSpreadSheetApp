using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace BigSpreadSheetApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Stopwatch stopWatch = new Stopwatch();

            stopWatch.Start();

            BigSpreadSheetParser parser = new BigSpreadSheetParser(@"c:/temp/BigSpreadsheetAllTypes.xlsx");
            List<List<object>> results = parser.ParseSpreadSheet();

            stopWatch.Stop();

            Console.WriteLine("Open XML SDK parsed {0} row(s) in {1} seconds", results.Count, stopWatch.ElapsedMilliseconds / 1000);

            Console.ReadLine();
        }
    }
}
