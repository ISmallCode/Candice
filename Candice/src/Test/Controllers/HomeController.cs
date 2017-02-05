using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System.IO;
using Candice;

namespace Test.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment hostingEnv;
        public HomeController(IHostingEnvironment env)
        {
            this.hostingEnv = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file, bool notUsed = false)
        {
            if (file != null)
            {
                var guidImportFileName = $"{Guid.NewGuid()}.csv";
                //ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');   //获取文件名
                var importFilePath = hostingEnv.WebRootPath + $@"\{guidImportFileName}";
                using (FileStream fs = System.IO.File.Create(importFilePath))
                {
                    file.CopyTo(fs);
                    fs.Flush();
                }

                var guidExportFileName = $"{Guid.NewGuid()}.txt";
                var exportFilePath = hostingEnv.WebRootPath + $@"\{guidExportFileName}";
                ImportCSV importCSV = new ImportCSV(importFilePath, exportFilePath);
            }

            return RedirectToAction("Index");
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
