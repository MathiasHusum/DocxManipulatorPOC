using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxManipulator
{
    public class Modifier
    {
        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("Dato");
                docText = regexText.Replace(docText, "Niels");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        public static void ConvertToPdf(string docxFile, string odtFile, string path)
        {
            try
            {
                using (Process pdfProcess = new Process())
                {
                    pdfProcess.StartInfo.UseShellExecute = false;
                    pdfProcess.StartInfo.RedirectStandardOutput = true;
                    pdfProcess.StartInfo.WorkingDirectory = path;
                    pdfProcess.StartInfo.FileName = "soffice";
                    pdfProcess.StartInfo.Arguments = $" --headless --convert-to pdf {odtFile}";
//                    pdfProcess.StartInfo.Verb = "runas";
                    pdfProcess.Start();
                    string output = pdfProcess.StandardOutput.ReadToEnd();
                    Trace.WriteLine(output);
                    pdfProcess.WaitForExit();
                    
                }
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
        
    }
    
}