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
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxManipulator
{
    public class Modifier
    {
        public static void SearchAndReplace(string fullPathToDocument)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fullPathToDocument, true))
            {
          
                List<KeyValuePair<int, string>> fieldValues = new List<KeyValuePair<int, string>>();
            
            
                fieldValues.Add(new KeyValuePair<int, string>(1 , "wolf"));
                fieldValues.Add(new KeyValuePair<int, string>(2 , "human"));
                fieldValues.Add(new KeyValuePair<int, string>(3 , "loch ness monster"));
                fieldValues.Add(new KeyValuePair<int, string>(4 , "Ape"));
                fieldValues.Add(new KeyValuePair<int, string>(5 , "&#10004;"));
                fieldValues.Add(new KeyValuePair<int, string>(6 , "Monkey"));
                fieldValues.Add(new KeyValuePair<int, string>(7 , "Lords"));
                
//                var body = wordDoc.MainDocumentPart.Document.Body;

                var body = wordDoc.MainDocumentPart.Document.Body;

                var paragraphs = body.Descendants<Paragraph>();
              
                
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                foreach (var fieldValue in fieldValues)
                {  
                    Regex regexText = new Regex($"F_{fieldValue.Key}");
                    docText = regexText.Replace(docText, fieldValue.Value);
                    if (fieldValue.Key == 4)
                    {
                        HighlightWord(wordDoc, $"F_{fieldValue.Key}");
                    }
                }

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                    sw.Flush();
                    sw.Close();
                }
                
            }
        }
        
        public static void HighlightWord(WordprocessingDocument wordDoc, string word)
        {
            Body body = wordDoc.MainDocumentPart.Document.Body;
            var paragraph = body.Descendants<Paragraph>().Where(x => x.InnerText == word);
            
            foreach (var para in paragraph)
            {
                var subRuns = para.Descendants<Run>().ToList();
                foreach (var run in subRuns)
                {
                    var subRunProp = run.Descendants<RunProperties>().ToList().FirstOrDefault();
                    var newColor = new Color();
                    newColor.Val = "EF413D";
                    
                    if (subRunProp != null)
                    {
                        var color = subRunProp.Descendants<Color>().FirstOrDefault();
                        subRunProp.ReplaceChild(newColor, color);
                    }
                    else
                    {
                        var tmpSubRunProp = new RunProperties();
                        tmpSubRunProp.AppendChild(newColor);
                        run.AppendChild(tmpSubRunProp);
                    }
                }
            }
            wordDoc.MainDocumentPart.Document.Save();
        }
        public static void ConvertToPdf(string docxFileName, string fullPathToDocument)
        {
            try
            {
                using (Process pdfProcess = new Process())
                {
                    pdfProcess.StartInfo.UseShellExecute = false;
                    pdfProcess.StartInfo.RedirectStandardOutput = true;
                    pdfProcess.StartInfo.WorkingDirectory = fullPathToDocument;
                    pdfProcess.StartInfo.FileName = "soffice";
                    pdfProcess.StartInfo.Arguments = $" --headless --convert-to pdf {docxFileName}";
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

        public static void AppendStyle(WordprocessingDocument wordDoc, string word, string col)
        {
            try
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var paragraphs = body.Elements<Paragraph>();
                var color = new Color();

                foreach (var para in paragraphs)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text.Contains(word))
                            {
                                color.Val = col;
                                run.AppendChild(color);
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
        
        
        public static void Insert(string fullPathToDocument, string fullPathToImageFile)
        {
            List<KeyValuePair<string, string>> keyValuePairs = new List<KeyValuePair<string, string>>();
            keyValuePairs.Add(new KeyValuePair<string, string>("Deponi", "Picture 1"));
            keyValuePairs.Add(new KeyValuePair<string, string>("Deponi", "Picture 2"));
            keyValuePairs.Add(new KeyValuePair<string, string>("Deponi", "Picture 3"));
            keyValuePairs.Add(new KeyValuePair<string, string>("Have", "Picture 4"));
            keyValuePairs.Add(new KeyValuePair<string, string>("Metal", "Picture 5"));
            keyValuePairs.Add(new KeyValuePair<string, string>("Metal", "Picture 6"));
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fullPathToDocument, true))
            {
                string currentHeader = "";

                foreach (var keyValuePair in keyValuePairs)
                {
                    if (currentHeader != keyValuePair.Key)
                    {
                        if (!string.IsNullOrEmpty(currentHeader))
                        {
                            // insert pakebreak
                            Body body = wordDoc.MainDocumentPart.Document.Body;

                            Paragraph para = body.AppendChild(new Paragraph());
                            Run run = para.AppendChild(new Run());
                            Break pageBreak = run.AppendChild(new Break());
                            pageBreak.Type = BreakValues.Page;
                        }
                        InsertHeader(keyValuePair.Key, wordDoc, currentHeader);
                        currentHeader = keyValuePair.Key;
                    }
                    
                    InsertPicture(keyValuePair.Value, wordDoc);
                }
            }

        }

        public static void InsertHeader(string header, WordprocessingDocument wordDoc, string currentHeader)
        {
            if (header != currentHeader)
            {
                //if currentHeader is not equal to new header, insert new header.
                currentHeader = header;
                Body body = wordDoc.MainDocumentPart.Document.Body;

                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(currentHeader));
            }
        }

        public static void InsertPicture(string fullPathToImageFile, WordprocessingDocument wordDoc)
        {
           
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fullPathToImageFile, FileMode.Open)) {
                imagePart.FeedData(stream);
            }
            AddImageToBody(wordDoc, mainPart.GetIdOfPart(imagePart));
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
        
        // Define the reference of the image.
        var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 6000000L, Cy = 4000000L },
                     new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, 
                         RightEdge = 0L, BottomEdge = 0L },
                     new DW.DocProperties() { Id = (UInt32Value)1U, 
                         Name = "Picture" },
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