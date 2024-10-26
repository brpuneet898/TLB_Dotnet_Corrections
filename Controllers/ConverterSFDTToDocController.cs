using System;
using System.IO;
using System.Net;
using System.Net.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Microsoft.AspNetCore.Cors;
using WFormatType = Syncfusion.DocIO.FormatType;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;

namespace TheLegalBook.Controllers
{
    [ApiController]
    [Route("api/convert")]
    public class ConverterSFDTToDocController : ControllerBase
    {
        private MemoryStream _docxStream;

        [HttpPost]
        [Route("Save")]
        public IActionResult Save([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            string format = RetrieveFileType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Save(data.Content);
            FileStream fileStream = new FileStream(name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, GetWFormatType(format));
            document.Close();

            // Return the document as a file stream response
            return File(fileStream, "application/msword", name);
        }
        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }
        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
        }
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        //[HttpPost]
        //[Route("sfdt-to-docx")]
        //public IActionResult ConvertSFDTToDOCX([FromBody] string sfdtContent)
        //{
        //    try
        //    {
        //        // Load the SFDT content
        //        using (MemoryStream sfdtStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(sfdtContent)))
        //        {
        //            // Create a new Word document
        //            Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument();
        //            // Open the SFDT stream
        //            document.Open(sfdtStream, Syncfusion.DocIO.FormatType.Txt);
        //            // Save the document as DOCX
        //            _docxStream = new MemoryStream();
        //            document.Save(_docxStream, Syncfusion.DocIO.FormatType.Docx);
        //            // Reset the position of the DOCX stream
        //            _docxStream.Position = 0;
        //            return File(_docxStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "result.docx");
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        // Handle exceptions here
        //        return BadRequest("An error occurred.");
        //    }
        //}
    }
}