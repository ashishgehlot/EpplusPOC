using OfficeOpenXml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Epplus.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public FileResult Generate()
        {
            FileInfo file = new FileInfo(Path.Combine(Server.MapPath("~/"), "template.xlsx"));
            MemoryStream stream = new MemoryStream();


            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                excelWorksheet.Cells[12, 6].Value = DateTime.Now.Date.ToString();


                var outputPath = Path.Combine(Server.MapPath("~/"), $"{Guid.NewGuid()}.xlsx").ToString();
                FileInfo fn = new FileInfo(outputPath);
                excelPackage.SaveAs(fn);

                Workbook workbook = new Workbook();
                ////Load excel file  
                workbook.LoadFromFile(outputPath);
                ////Save excel file to pdf file.  
                var pdfPath = Path.Combine(Server.MapPath("~/"), $"{Guid.NewGuid()}.pdf").ToString();
                workbook.SaveToFile(pdfPath);

                //this.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //this.Response.AddHeader("content-disposition", string.Format("attachment;  filename={0}", "ExcellData.pdf"));
                //return File(stream, System.Net.Mime.MediaTypeNames.Application.Octet, "demo.pdf");

                var cd = new System.Net.Mime.ContentDisposition
                {
                    // for example foo.bak
                    FileName = "dmeo.pdf",

                    // always prompt the user for downloading, set to true if you want 
                    // the browser to try to show the file inline
                    Inline = false,
                };
                Response.AppendHeader("Content-Disposition", cd.ToString());
                return File(pdfPath, "application/pdf");
            }



            return File(stream, System.Net.Mime.MediaTypeNames.Application.Octet, "demo.xlsx");
        }

        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}