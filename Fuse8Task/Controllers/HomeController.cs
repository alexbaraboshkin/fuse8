using Fuse8Task.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fuse8Task.Controllers
{
    public class HomeController : Controller
    {
        public const string FROM = "abaraboshkin@yandex.ru";
        public const string PASSWORD = "*********************";

        public ViewResult Index()
        {
            return View();
        }

        private static void SendMail(string smtpServer, string from, string password, string mailto, string caption, string message, string attachFile = null)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(from);
                mail.To.Add(new MailAddress(mailto));
                mail.Subject = caption;
                mail.Body = message;
                if (!string.IsNullOrEmpty(attachFile))
                    mail.Attachments.Add(new Attachment(attachFile));
                SmtpClient client = new SmtpClient();
                client.Host = smtpServer;
                //client.Port = 587; // для gmail
                client.Port = 25;    // для yandex
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(from.Split('@')[0], password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(mail);
                mail.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Mail.Send: " + e.Message);
            }
        }

        [HttpPost]
        public ViewResult Report(EmailDateViewModel emailDateViewModel)
        {
            Context context = new Context();

            if (!ModelState.IsValidField("emailForSendReport"))
            {
                ViewBag.Error = "неверно введен email";
                return View("Error");
            }
            if (!ModelState.IsValidField("datepicker"))
            {
                ViewBag.Error = "неверно введена начальная дата";
                return View("Error");
            }
            if (!ModelState.IsValidField("datepicker1"))
            {
                ViewBag.Error = "неверно введена конечная дата";
                return View("Error");
            }
            string formatString = "dd-mm-yyyy";
            DateTime datetime0 = DateTime.ParseExact(emailDateViewModel.datepicker, formatString, CultureInfo.InvariantCulture);
            DateTime datetime1 = DateTime.ParseExact(emailDateViewModel.datepicker1, formatString, CultureInfo.InvariantCulture);

            var products = from ep in context.Product
                           join e in context.OrderDetail on ep.ID equals e.ProductID
                           join t in context.Order on e.OrderID equals t.ID
                           where t.OrderDate > datetime0 && t.OrderDate < datetime1
                           select new MyViewModel
                           {
                               ProductId = ep.ID,
                               OderID = t.ID,
                               OrderDate = t.OrderDate,
                               Quantity = e.Quantity,
                               UnitPrice = e.UnitPrice
                           };

            // инициализируем объект приложения Excel
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            // создаем новую книгу
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            // получаем доступ к таблице, т.е. листу 1
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "OderID";
            xlWorkSheet.Cells[1, 2] = "OrderDate";
            xlWorkSheet.Cells[1, 3] = "ProductId";
            xlWorkSheet.Cells[1, 4] = "Quantity";
            xlWorkSheet.Cells[1, 5] = "UnitPrice";
            xlWorkSheet.Cells[1, 6] = "Cost";
            int i = 2;
            foreach (MyViewModel myViewModel in products)
            {
                xlWorkSheet.Cells[i, 1] = myViewModel.OderID;
                xlWorkSheet.Cells[i, 2] = myViewModel.OrderDate;
                Excel.Range c2 = xlWorkSheet.Cells[i, 2];
                c2.EntireColumn.AutoFit();
                xlWorkSheet.Cells[i, 3] = myViewModel.ProductId;
                xlWorkSheet.Cells[i, 4] = myViewModel.Quantity;
                xlWorkSheet.Cells[i, 5] = myViewModel.UnitPrice;
                xlWorkSheet.Cells[i, 6] = "=RC[-2]*RC[-1]";
                i++;
            }

            // предварительно удалим файл, если существует
            string path = @"C:\\csharp-Excel.xls";
            FileInfo fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }

            xlWorkBook.SaveAs("C:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ViewBag.emailForSendReport = emailDateViewModel.emailForSendReport;
            SendMail("smtp.yandex.ru", FROM, PASSWORD, emailDateViewModel.emailForSendReport, "Отчет Excel", "Ваш отчет Excel", "C:\\csharp-Excel.xls");

            return View("Report", products);
        }

    }
}
