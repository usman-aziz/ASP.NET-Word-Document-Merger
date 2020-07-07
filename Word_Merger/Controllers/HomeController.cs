using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Word_Merger.Models;
using Aspose.Words;
using Aspose.Words.Saving;
using Microsoft.AspNetCore.Http;
using System.IO;

namespace Word_Merger.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public FileResult UploadFiles(List<IFormFile> files, string outputFormat)
        {
            if (files.Count() <= 1)
            {
                // display some message
                return null;
            }
            string fileName = "merged-document.docx";
            string path = "wwwroot/uploads";
            List<Document> documents = new List<Document>();
            // upload files 
            foreach (IFormFile file in files)
            {
                string filePath = Path.Combine(path, file.FileName);
                // Save files
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                // Add all documents to the list
                documents.Add(new Document(filePath));
            }
            // Load first Word document
            Document doc1 = documents[0];
            for (int i = 1; i < documents.Count(); i++)
            {
                doc1.AppendDocument(documents[i], ImportFormatMode.KeepSourceFormatting);
            }           

            var outputStream = new MemoryStream(); 
            if (outputFormat == "DOCX")
            {
                doc1.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0;
                // Return generated Word file
                return File(outputStream, System.Net.Mime.MediaTypeNames.Application.Rtf, fileName);
            }
            else
            {
                fileName = "merged-document.pdf";
                doc1.Save(outputStream, SaveFormat.Pdf);
                outputStream.Position = 0;
                // Return generated PDF file
                return File(outputStream, System.Net.Mime.MediaTypeNames.Application.Pdf, fileName);
            }
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
