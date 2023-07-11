using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using Aspose.Pdf;
using static iTextSharp.text.pdf.AcroFields;
using Aspose.Pdf.Operators;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

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
                //Upload the word document and save to Temp folder.
                postedFile.SaveAs(fileSavePath.ToString());

                //Open the word document in background.
                _Application applicationclass = new Application();
                applicationclass.Documents.Open(ref fileSavePath);
                applicationclass.Visible = false;
                var document = applicationclass.ActiveDocument;

                //Save the word document as HTML file.
                document.SaveAs(ref htmlFilePath, ref documentFormat);

                //Close the word document.
                document.Close();

                //Read the saved Html File.
                string wordHTML = System.IO.File.ReadAllText(htmlFilePath.ToString());

                //Loop and replace the Image Path.
                //foreach (Match match in Regex.Matches(wordHTML, "<v:imagedata.+?src=[\"'](.+?)[\"'].*?>", RegexOptions.IgnoreCase))
                //{
                //    wordHTML = Regex.Replace(wordHTML, match.Groups[1].Value, "Temp/" + match.Groups[1].Value);
                //}

                //Delete the Uploaded Word File.
                System.IO.File.Delete(fileSavePath.ToString());

                ViewBag.WordHtml = wordHTML;
            }
            else if (exention.ToLower() == ".pdf".ToLower())
            {
                if (!System.IO.Directory.Exists("~/DestinationFiles/"))
                {
                    System.IO.Directory.CreateDirectory("~/DestinationFiles/");
                }
                string outputPath = HttpContext.Server.MapPath("~/DestinationFiles/");
                StringBuilder text = new StringBuilder();
                using (PdfReader reader = new PdfReader(postedFile.InputStream))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        var x = PdfTextExtractor.GetTextFromPage(reader, i);
                        text.AppendLine(PdfTextExtractor.GetTextFromPage(reader, i));
                    }
                }
                using (StreamWriter outputFile = new StreamWriter(System.IO.Path.Combine(outputPath, "Pdf2Text.txt")))
                {
                    outputFile.WriteLine(text);
                }
            }

            return View();
        }
    }
}