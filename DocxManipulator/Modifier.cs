using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxManipulator
{
    public class Modifier
    {
        public static void SearchAndReplace(string document)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                List<KeyValuePair<int, string>> fieldValues = new List<KeyValuePair<int, string>>();
            
            
                fieldValues.Add(new KeyValuePair<int, string>(1 , "wolf"));
                fieldValues.Add(new KeyValuePair<int, string>(2 , "human"));
                fieldValues.Add(new KeyValuePair<int, string>(3 , "loch ness monster"));
                fieldValues.Add(new KeyValuePair<int, string>(4 , "Ape"));
                fieldValues.Add(new KeyValuePair<int, string>(5 , "Qutie"));
                fieldValues.Add(new KeyValuePair<int, string>(6 , "Klima"));
                fieldValues.Add(new KeyValuePair<int, string>(7 , "Oof"));
                
//                var body = wordDoc.MainDocumentPart.Document.Body;

                
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                
                foreach (var fieldValue in fieldValues)
                {
                    Regex regexText = new Regex($"F_{fieldValue.Key}");
                    docText = regexText.Replace(docText, fieldValue.Value);
                }

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                    sw.Flush();
                    sw.Close();
                }
                
            }
        }

        public static void ConvertToPdf(string docxFile, string path)
        {
            try
            {
                using (Process pdfProcess = new Process())
                {
                    pdfProcess.StartInfo.UseShellExecute = false;
                    pdfProcess.StartInfo.RedirectStandardOutput = true;
                    pdfProcess.StartInfo.WorkingDirectory = path;
                    pdfProcess.StartInfo.FileName = "soffice";
                    pdfProcess.StartInfo.Arguments = $" --headless --convert-to pdf {docxFile}";
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

        public static void InsertPicture(string document, string fileName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordDoc, mainPart.GetIdOfPart(imagePart));

            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
    {
        
        // Define the reference of the image.
        var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, 
                         RightEdge = 0L, BottomEdge = 0L },
                     new DW.DocProperties() { Id = (UInt32Value)1U, 
                         Name = "Picture 1" },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties() 
                                        { Id = (UInt32Value)0U, 
                                            Name = "New Bitmap Image.jpg" },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension() 
                                                { Uri = 
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}" })
                                     ) 
                                     { Embed = relationshipId, 
                                         CompressionState = 
                                         A.BlipCompressionValues.Print },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     ) { Preset = A.ShapeTypeValues.Rectangle }))
                         ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                 ) { DistanceFromTop = (UInt32Value)0U, 
                     DistanceFromBottom = (UInt32Value)0U, 
                     DistanceFromLeft = (UInt32Value)0U, 
                     DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });
        
        // Append the reference to body, the element should be in a Run.
       wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
    }
    }
    
}