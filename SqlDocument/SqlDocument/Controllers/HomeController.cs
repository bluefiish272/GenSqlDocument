using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using SqlDocument.Models;

namespace SqlDocument.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            var spInfos = new DbConnection().GetSpDescription();
            // 選擇要用哪種套件產文件
            // new NpoiGenerateDocument().GenerateWord(spInfos);
            new OenpXmlGenDoc().GenerateWord(spInfos);
            return View(spInfos);
        }

        public IActionResult GenerateWord()
        {
            return View();
            //return File();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
