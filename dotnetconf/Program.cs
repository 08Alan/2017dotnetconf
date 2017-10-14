using System;
using dotnetconflib;
using dotnetconflib.EpplusSample.Sample01;
using dotnetconflib.DBModel.DBTools;

namespace dotnetconf
{
    class Program
    {
        static void Main(string[] args)
        {
            Sample01 demo = new Sample01();
            demo.RunSample1();

            // ------ Entity Framework Code ------ 
            DBTools db = new DBTools();
            //------ Demo connection with DB ------ 
            //db.storeData();
            db.getData();

            // ------ DB export to Excel ------ 
             db.DataTableSFToExcelFile("DB_To_Excel.xlsx");

        }
    }
}
