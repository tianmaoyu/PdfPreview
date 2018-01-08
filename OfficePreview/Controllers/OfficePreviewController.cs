using OfficeConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficePreview.Controllers
{
    public class OfficePreviewController : Controller
    {
       
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile()
        {
            var pdfFileUrl = "";

            HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
            foreach (string key in files.AllKeys)
            {
                HttpPostedFile file = files[key];
                if (file != null && file.ContentLength > 0)
                {
                    var fileName = Path.Combine(Server.MapPath("~/Upload/"), file.FileName);
                    file.SaveAs(fileName);
                    var extractor = new Converter();
                    var _name = Path.GetFileNameWithoutExtension(fileName);
                    var pdfFile = Path.Combine(Server.MapPath("~/Upload/"), _name + ".pdf");
                    extractor.Convert(fileName, pdfFile);
                    pdfFileUrl = "/Upload/" + _name + ".pdf";
                    break;
                }
            }
            return Json(pdfFileUrl);
        }



        public ActionResult Upload()
        {
            return View();
        }
        public ActionResult Preview(string url = "/Upload/00.pdf")
        {
            ViewData["url"] = url;
            return View();
        }



    }
}