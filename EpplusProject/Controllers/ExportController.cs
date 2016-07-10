using EpplusProject.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CRUDDeom.Controllers
{
    public class ExportController : Controller
    {
        //private DbContextModel db = new DbContextModel();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Excel()
        {

            //List accounts = db.Accounts.ToList();

            DataTable dt = new DataTable("Customers");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("Phone", typeof(string));
            dt.Rows.Add("Ram", "ram@techbrij.com", "111-222-3333");
            dt.Rows.Add("Shyam", "shyam@techbrij.com", "159-222-1596");
            dt.Rows.Add("Mohan", "mohan@techbrij.com", "456-222-4569");
            dt.Rows.Add("Sohan", "sohan@techbrij.com", "789-456-3333");
            dt.Rows.Add("Karan", "karan@techbrij.com", "111-222-1234");
            dt.Rows.Add("Brij", "brij@techbrij.com", "111-222-3333");

            var data = new[]{
                               new{ Name="Ram", Email="ram@techbrij.com", Phone="111-222-3333" },
                               new{ Name="Shyam", Email="shyam@techbrij.com", Phone="159-222-1596" },
                               new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                               new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                               new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                               new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }
                      };

            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Accounts");
                //ws.Cells["A1"].LoadFromCollection(dt, true);
                ws.Cells["A1"].LoadFromDataTable(dt,true);
                // Load your collection "accounts"

                Byte[] fileBytes = pck.GetAsByteArray();
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment;filename=DataTable.xlsx");
                // Replace filename with your custom Excel-sheet name.

                Response.Charset = "";
                Response.ContentType = "application/vnd.ms-excel";
                StringWriter sw = new StringWriter();
                Response.BinaryWrite(fileBytes);
                Response.End();
            }

            return RedirectToAction("Index");
        }

        public ActionResult CSV()
        {
            DataTable dt = new DataTable("Customers");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("Phone", typeof(string));
            dt.Rows.Add("Ram", "ram@techbrij.com", "111-222-3333");
            dt.Rows.Add("Shyam", "shyam@techbrij.com", "159-222-1596");
            dt.Rows.Add("Mohan", "mohan@techbrij.com", "456-222-4569");
            dt.Rows.Add("Sohan", "sohan@techbrij.com", "789-456-3333");
            dt.Rows.Add("Karan", "karan@techbrij.com", "111-222-1234");
            dt.Rows.Add("Brij", "brij@techbrij.com", "111-222-3333");

            var data = new[]{
                               new{ Name="Ram", Email="ram@techbrij.com", Phone="111-222-3333" },
                               new{ Name="Shyam", Email="shyam@techbrij.com", Phone="159-222-1596" },
                               new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                               new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                               new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                               new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }
                      };

            StringWriter sw = new StringWriter();
            sw.WriteLine("\"Name\",\"Email\",\"Phone\"");
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment;filename=Exported_Users.csv");
            Response.ContentType = "text/csv";

            foreach (var line in data)
            {
                sw.WriteLine(string.Format("\"{0}\",\"{1}\",\"{2}\"",
                                           line.Name,
                                           line.Email,
                                           line.Phone));
            }

            Response.Write(sw.ToString());
            Response.End();

            return RedirectToAction("Index");
        }

    }
}