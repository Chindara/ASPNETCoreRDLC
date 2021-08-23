using AspNetCore.Reporting;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Data;

namespace MVCCoreRDLC.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public HomeController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult PDF()
        {
            DataTable dt = GetData();
            string parameter = "Report Date: "+DateTime.Today.ToString("yyyy-MM-dd");

            string mimetype = "";
            int extension = 1;
            var path = $"{this._webHostEnvironment.WebRootPath}\\reports\\Report1.rdlc";

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("prm", parameter);

            LocalReport localReport = new LocalReport(path);
            localReport.AddDataSource("dsEmployee", dt);

            var result = localReport.Execute(RenderType.Pdf, extension, parameters, mimetype);

            return File(result.MainStream, "application/pdf", "Employees.pdf");
        }

        public IActionResult Excel()
        {
            DataTable dt = GetData();
            string parameter = "Report Date: " + DateTime.Today.ToString("yyyy-MM-dd");

            string mimetype = "";
            int extension = 1;
            var path = $"{this._webHostEnvironment.WebRootPath}\\reports\\Report1.rdlc";

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("prm", parameter);

            LocalReport localReport = new LocalReport(path);
            localReport.AddDataSource("dsEmployee", dt);

            var result = localReport.Execute(RenderType.Excel, extension, parameters, mimetype);

            return File(result.MainStream, "application/msexcel", "Employees.xls");
        }

        public IActionResult Word()
        {
            DataTable dt = GetData();
            string parameter = "Report Date: " + DateTime.Today.ToString("yyyy-MM-dd");

            string mimetype = "";
            int extension = 1;
            var path = $"{this._webHostEnvironment.WebRootPath}\\reports\\Report1.rdlc";

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("prm", parameter);

            LocalReport localReport = new LocalReport(path);
            localReport.AddDataSource("dsEmployee", dt);

            var result = localReport.Execute(RenderType.Word, extension, parameters, mimetype);

            return File(result.MainStream, "application/msword", "Employees.doc");
        }

        public DataTable GetData()
        {
            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Email");
            dt.Columns.Add("Age");
            dt.Columns.Add("Designation");

            DataRow dr;
            for (int i = 0; i < 100; i++)
            {
                dr = dt.NewRow();
                dr["Name"] = "Name " + i;
                dr["Email"] = "Email " + i;
                dr["Age"] = "Age " + i;
                dr["Designation"] = "Designation " + i;

                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}