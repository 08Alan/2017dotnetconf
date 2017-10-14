using System;
using System.IO;
using OfficeOpenXml;

namespace dotnetconflib.EpplusSample.Sample02
{
    /// <summary>
    /// Simply opens an existing file and reads some values and properties
    /// </summary>
    public class Sample02
    {
        public void RunSample2()
        {
            var filePath = Path.Combine("", @"Sample1.xlsx");
            Console.WriteLine("Reading column 2 of {0}", filePath);
            Console.WriteLine();

            FileInfo existingFile = new FileInfo(filePath);
            if(!existingFile.Exists)
            {
                return;
            }

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int col = 2; //The item description
                             // output the data in column 2
                for (int row = 2; row < 5; row++)
                    Console.WriteLine("\tCell({0},{1}).Value={2}", row, col, worksheet.Cells[row, col].Value);

                // output the formula in row 5
                Console.WriteLine("\tCell({0},{1}).Formula={2}", 3, 5, worksheet.Cells[3, 5].Formula);
                Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells[3, 5].FormulaR1C1);

                // output the formula in row 5
                Console.WriteLine("\tCell({0},{1}).Formula={2}", 5, 3, worksheet.Cells[5, 3].Formula);
                Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells[5, 3].FormulaR1C1);

            } // the using statement automatically calls Dispose() which closes the package.

            Console.WriteLine();
            Console.WriteLine("Sample 2 complete");
            Console.WriteLine();
        }
    }
}