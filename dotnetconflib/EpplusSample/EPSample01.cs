using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace dotnetconflib.EpplusSample.Sample01
{
    public class Sample01
    {
        /// <summary>
        /// Sample 1 - simply creates a new workbook from scratch.
        /// The workbook contains one worksheet with a simple invertory list
        /// </summary>
        public void RunSample1()
        {
            FileInfo newFile = new FileInfo(Path.Combine("", "Sample1.xlsx"));
            if(newFile.Exists)
            {
                newFile.Delete();
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                //Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
                //Add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Value";

                //Add some items...
                worksheet.Cells["A2"].Value = 12001;
                worksheet.Cells["B2"].Value = "Nails";
                worksheet.Cells["C2"].Value = 37;
                worksheet.Cells["D2"].Value = 3.99;

                worksheet.Cells["A3"].Value = 12002;
                worksheet.Cells["B3"].Value = "Hammer";
                worksheet.Cells["C3"].Value = 5;
                worksheet.Cells["D3"].Value = 12.10;

                worksheet.Cells["A4"].Value = 12003;
                worksheet.Cells["B4"].Value = "Saw";
                worksheet.Cells["C4"].Value = 12;
                worksheet.Cells["D4"].Value = 15.37;
                //Add a formula for the value-column
                worksheet.Cells["E2:E4"].Formula = "C2*D2";
                // //Ok now format the values;
                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    // Set Font to Bold
                    range.Style.Font.Bold = true;
                    // Set BackgroundColor to Gray
                    range.Style.Fill.PatternType = ExcelFillStyle.Gray0625;
                }
                // Set Top Border
                worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                // Set Font to Bold
                worksheet.Cells["A5:E5"].Style.Font.Bold = true;
                worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 3, 4, 3).Address);
                worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";
                // //Create an autofilter for the range
                worksheet.Cells["A1:E4"].AutoFilter = true;
                worksheet.Cells["A2:A4"].Style.Numberformat.Format = "@";   //Format as text
                worksheet.Calculate();
                package.Save();

            }
        }
    }
}