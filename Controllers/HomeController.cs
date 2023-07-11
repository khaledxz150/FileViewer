using System;
using System.IO;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using System.Drawing.Imaging;
using System.Drawing;

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
                string outputImagePath = HttpContext.Server.MapPath("~/DestinationFiles/"+ FileName);

                using (var rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
                {
                    rasterizer.Open(postedFile.InputStream);

                    // Set the resolution (in DPI) for the output image
                    int dpi = 300;

                    // Get the dimensions of the first page
                    var firstPageImage = rasterizer.GetPage(dpi, 1);
                    int imageWidth = firstPageImage.Width;
                    int imageHeight = firstPageImage.Height;

                    // Create a blank canvas to hold the stacked image
                    using (var stackedImage = new Bitmap(imageWidth, imageHeight * rasterizer.PageCount))
                    using (var graphics = Graphics.FromImage(stackedImage))
                    {
                        // Set the background color (optional)
                        graphics.Clear(Color.White);

                        // Stack the pages on the canvas
                        for (int pageIndex = 1; pageIndex <= rasterizer.PageCount; pageIndex++)
                        {
                            using (var pageImage = rasterizer.GetPage(dpi, pageIndex))
                            {
                                // Calculate the position to draw the current page
                                int yPos = (pageIndex - 1) * imageHeight;

                                // Draw the page onto the stacked image
                                graphics.DrawImage(pageImage, 0, yPos);
                            }
                        }

                        // Save the stacked image as PNG
                        stackedImage.Save(outputImagePath, ImageFormat.Png);
                    }
                }
                ViewBag.ImagePath = "/DestinationFiles/"+FileName;
            }
            return View();
        }
    }
}