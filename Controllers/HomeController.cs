using System;
using System.IO;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using System.Drawing.Imaging;

namespace FileViewer.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            object documentFormat = 8;
            string randomName = DateTime.Now.Ticks.ToString();
            object htmlFilePath = Server.MapPath("~/Temp/") + randomName + ".htm";
            string directoryPath = Server.MapPath("~/Temp/") + randomName + "_files";
            if (!System.IO.Directory.Exists("~/Temp/"))
            {
                System.IO.Directory.CreateDirectory("~/Temp/");
            }
            object fileSavePath = Server.MapPath("~/Temp/") + System.IO.Path.GetFileName(postedFile.FileName);
            //If Directory not present, create it.
            if (!Directory.Exists(Server.MapPath("~/Temp/")))
            {
                Directory.CreateDirectory(Server.MapPath("~/Temp/"));
            }
            var exention = System.IO.Path.GetExtension(postedFile.FileName);
            if (exention.ToLower() == ".docx".ToLower())
            {
                postedFile.SaveAs(fileSavePath.ToString());

                _Application applicationclass = new Application();
                applicationclass.Documents.Open(ref fileSavePath);
                applicationclass.Visible = false;
                var document = applicationclass.ActiveDocument;

                document.SaveAs(ref htmlFilePath, ref documentFormat);
                document.Close();
                string wordHTML = System.IO.File.ReadAllText(htmlFilePath.ToString());

                System.IO.File.Delete(fileSavePath.ToString());
                ViewBag.WordHtml = wordHTML;
            }
            else if (exention.ToLower() == ".pdf".ToLower())
            {
                var FileName = Path.GetFileName(postedFile.FileName);
                var Extenions = Path.GetExtension(postedFile.FileName);
                string outputImagePath = HttpContext.Server.MapPath("~/DestinationFiles/"+ FileName+"/"+ Extenions);

                // Convert PDF to image
                using (var rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
                {
                    rasterizer.Open(postedFile.InputStream);

                    // Set the resolution (in DPI) for the output image
                    int dpi = 300;

                    // Set the page index (starting from 1)
                    int pageIndex = 1;

                    // Render the PDF page to an image
                    using (var image = rasterizer.GetPage(dpi, pageIndex))
                    {
                        // Save the image as PNG
                        image.Save(outputImagePath, ImageFormat.Png);
                    }
                }
                ViewBag.ImagePath = FileName + Extenions;
            }
            return View();
        }
    }
}