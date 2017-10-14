using System;
using System.Drawing;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

// EPPLUS
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

using dotnetconflib.DBModel.dotnetconfDBContext;
using dotnetconflib.Entity.Blog;
using dotnetconflib.Entity.Post;
using System.IO;

namespace dotnetconflib.DBModel.DBTools
{
    public class DBTools
    {
        List<Blog> blog_list = new List<Blog>()
        {
            new Blog { Url = "http://blogs.msdn.com/adonet", Rating = 5 },
            new Blog { Url = "https://08alan.github.io/", Rating = 4 },
            new Blog { Url = "https://kji0.blogspot.tw/", Rating = 3 },
        };
        public void storeData()
        {
            using (var db = new dotnetconfDBContext.dotnetconfContext())
            {
                foreach (var sub_blog in blog_list)
                {
                    db.Blogs.Add(sub_blog);
                }
                db.SaveChanges();
            }
        }

        public void getData()
        {
            using (var db = new dotnetconfDBContext.dotnetconfContext())
            {
                Console.WriteLine("All blogs in database:");

                int count = 1;
                foreach (var blog in db.Blogs)
                {
                    Console.WriteLine(" - {0}", blog.Url);
                    count++;
                }

                Console.WriteLine("{0} records saved to database", count);
                Console.WriteLine();
            }
        }

        // Export from DB to Excel
        public void DataTableSFToExcelFile(string file_name)
        {
            FileInfo getFromTable = new FileInfo(file_name);
            if (getFromTable.Exists)
            {
                getFromTable.Delete();
            }

            using(var db = new dotnetconfDBContext.dotnetconfContext())
            {
                ExcelPackage ep = new ExcelPackage(getFromTable);
                ExcelWorksheet ws;

                if (db.Blogs.GetType() != null)
                {
                    ws = ep.Workbook.Worksheets.Add("AzureTaiwan");
                    var blog_set = db.Blogs;
                    
                    List<Blog> blog_list = new List<Blog>();
                    int count = 1;
                    foreach (var item in blog_set)
                    {
                        blog_list.Add(item);
                        count++;
                    }

                    // 設定 Header 樣式
                    using(ExcelRange rng = ws.Cells["A1:C" + count.ToString()])
                    {
                        // 設定粗體
                        rng.Style.Font.Bold = true;

                        // 
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;

                        // 設定框格底色
                        rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                        // rng.Style.Font.Color.SetColor(Color.White);

                        // 設定文字置中
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // Set Border
                        rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                    ws.Cells["A1"].Value = "Blog ID";
                    ws.Cells["B1"].Value = "Url";
                    ws.Cells["C1"].Value = "Rating";

                    // EF to Excel data
                    ws.Cells["A2"].LoadFromCollection(blog_list);

                    // Download Mono Frameworks, copy below files to usr/local/lib.
                    // /Library/Frameworks/Mono.framework/Versions/5.0.1/lib/libgdiplus.0.dylib
                    // /Library/Frameworks/Mono.framework/Versions/5.0.1/lib/libgdiplus.0.dylib.dSYM
                    // /Library/Frameworks/Mono.framework/Versions/5.0.1/lib/libgdiplus.dylib
                    // /Library/Frameworks/Mono.framework/Versions/5.0.1/lib/libgdiplus.la
                    ws.Cells["A1:C" + count.ToString()].AutoFitColumns();
                }
                else
                {
                    ws = ep.Workbook.Worksheets.Add("Nothing");
                    ws.Cells["A2"].Value = "There is no data table data.";
                }
                ep.Save();
            }
        }
    }
}