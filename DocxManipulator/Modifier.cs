using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
                
                var body = wordDoc.MainDocumentPart.Document.Body;

                
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
                    sw.Close();
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
            
           

            
          
            
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true)) 
            { 
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
                
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            List<KeyValuePair<int, string>> fieldValues = new List<KeyValuePair<int, string>>();

            fieldValues.Add(new KeyValuePair<int, string>(1 , "https://481.microting.com/api/template-files/get-image/7_6d9ee36ecf4ef6b56ad67c25d583ef52.jpeg"));
            fieldValues.Add(new KeyValuePair<int, string>(2 , "https://481.microting.com/api/template-files/get-image/8_a0149c2fa5349977529f84be3a589416.png"));
            fieldValues.Add(new KeyValuePair<int, string>(3 , "https://481.microting.com/api/template-files/get-image/4_07822c58643c467610a68d55665a04c5.jpeg"));
            fieldValues.Add(new KeyValuePair<int, string>(4 , "https://481.microting.com/api/template-files/get-image/87_e9ef7bf68d9d079da6a42b8846166b78.jpeg"));
            fieldValues.Add(new KeyValuePair<int, string>(5 , "https://481.microting.com/api/template-files/get-image/88_818222d2276c2e66955d3d5f95abfe9c.jpeg"));
            fieldValues.Add(new KeyValuePair<int, string>(6 , "https://481.microting.com/api/template-files/get-image/89_69da16266fa1d3ce22a9c133f5feef0b.jpeg"));
            fieldValues.Add(new KeyValuePair<int, string>(7 , "https://481.microting.com/api/template-files/get-image/90_4ebf3bb168c6efe4c880ce8d86c5fdf5.jpeg"));

            foreach (var fieldValue in fieldValues)
            {
                
           var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, 
                         RightEdge = 0L, BottomEdge = 0L },
                     new DW.DocProperties() { Id = (UInt32Value)1U, 
                         Name = $"Picture {fieldValue.Key}" },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties() 
                                        { Id = (UInt32Value)0U, 
                                            Name = "New Bitmap Image.png" },
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
                         ) { Uri = fieldValue.Value})
                 ) { DistanceFromTop = (UInt32Value)0U, 
                     DistanceFromBottom = (UInt32Value)0U, 
                     DistanceFromLeft = (UInt32Value)0U, 
                     DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

       // Append the reference to body, the element should be in a Run.
       wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
            }
        }
    }
    
}