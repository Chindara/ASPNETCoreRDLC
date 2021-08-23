﻿using AspNetCore.Reporting;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
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

        public IActionResult Print()
        {
            DataTable dt = GetData();

            string mimetype = "";
            int extension = 1;
            var path = $"{this._webHostEnvironment.WebRootPath}\\reports\\Report1.rdlc";

            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("prm", "DMS Software Engineering");

            LocalReport localReport = new LocalReport(path);
            localReport.AddDataSource("dsEmployee", dt);

            var result = localReport.Execute(RenderType.Pdf, extension, parameters, mimetype);

            return File(result.MainStream, "application/pdf");
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