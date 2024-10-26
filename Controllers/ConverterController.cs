using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using Microsoft.AspNetCore.Http;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Web;
using Syncfusion.EJ2.DocumentEditor;

namespace TheLegalBook.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ConverterController : Controller
    {        [HttpPost]
        [Route("DocToSfdt")]
        public IActionResult ConvertDocToSfdt(IFormFile docFile)
        {
            if (docFile == null || docFile.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            try
            {
                string fileExtension = Path.GetExtension(docFile.FileName)?.ToLower();
                FormatType formatType;

                try
                {
                    formatType = GetFormatType(fileExtension);
                }
                catch (NotSupportedException ex)
                {
                    return BadRequest(ex.Message);
                }

                string sfdt = ConvertDocumentToSfdt(docFile, formatType);
                return Ok(sfdt);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return StatusCode(500, $"An error occurred during conversion: {ex.Message}");
            }
        }

        private static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("File format is empty or null.");

            return format switch
            {
                ".docx" or ".dotx" or ".docm" or ".dotm" => FormatType.Docx,
                ".doc" or ".dot" => FormatType.Doc,
                ".rtf" => FormatType.Rtf,
                ".txt" => FormatType.Txt,
                ".xml" => FormatType.WordML,
                _ => throw new NotSupportedException($"The file format '{format}' is not supported.")
            };
        }

        private string ConvertDocumentToSfdt(IFormFile docFile, FormatType formatType)
        {
            string sfdt = "";
            using (MemoryStream memoryStream = new MemoryStream())
            {

                docFile.CopyTo(memoryStream);
                memoryStream.Position = 0;

                try
                {
                    switch (formatType)
                    {
                        case FormatType.Docx:
                        case FormatType.Doc:

                            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, false))
                            {
                                if (wordDoc.MainDocumentPart == null)
                                {
                                    throw new Exception("MainDocumentPart is missing. The file may be corrupted.");
                                }

                                Body body = wordDoc.MainDocumentPart.Document.Body;
                                string text = body.InnerText;

                                int imageCount = wordDoc.MainDocumentPart.ImageParts.Count();

                                var content = new
                                {
                                    Text = text,
                                    Images = imageCount
                                };

                                sfdt = JsonConvert.SerializeObject(content);
                            }
                            break;

                        case FormatType.Rtf:

                            sfdt = ConvertRtfToSfdt(memoryStream);
                            break;

                        case FormatType.Txt:

                            sfdt = ConvertTxtToSfdt(memoryStream);
                            break;

                        case FormatType.WordML:

                            sfdt = ConvertWordMLToSfdt(memoryStream);
                            break;

                        default:
                            throw new NotSupportedException("Only DOCX, DOC, RTF, TXT, and WordML formats are supported for SFDT conversion.");
                    }

                    if (string.IsNullOrEmpty(sfdt))
                    {
                        throw new Exception("Conversion to SFDT format failed.");
                    }

                    SaveSfdtToLocalFile(sfdt, docFile.FileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing document: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                    throw;
                }
            }

            return sfdt;
        }

        private string ConvertRtfToSfdt(Stream rtfStream)
        {

            using (StreamReader reader = new StreamReader(rtfStream))
            {
                string text = reader.ReadToEnd();
                return JsonConvert.SerializeObject(new { Text = text, Images = 0 });
            }
        }

        private string ConvertTxtToSfdt(Stream txtStream)
        {

            using (StreamReader reader = new StreamReader(txtStream))
            {
                string text = reader.ReadToEnd();
                return JsonConvert.SerializeObject(new { Text = text, Images = 0 });
            }
        }

        private string ConvertWordMLToSfdt(Stream wordMLStream)
        {

            using (StreamReader reader = new StreamReader(wordMLStream))
            {
                string text = reader.ReadToEnd();
                return JsonConvert.SerializeObject(new { Text = text, Images = 0 });
            }
        }
        private void SaveSfdtToLocalFile(string sfdt, string originalFileName)
        {
            try
            {
                string directoryPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedFiles");
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFileName);
                string sfdtFileName = $"{fileNameWithoutExtension}.sfdt";
                string filePath = Path.Combine(directoryPath, sfdtFileName);

                System.IO.File.WriteAllText(filePath, sfdt);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving SFDT to file: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                throw;
            }
        }
    }

    public enum FormatType
    {
        Docx,
        Doc,
        Rtf,
        Txt,
        WordML
    }
}