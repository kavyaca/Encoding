using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using CSV.Models.ApiData;

namespace CSV.Models.Utilities
{
    public class WordDocumentClass
    {


        public static void CreateWordprocessingDocumentFromApi(string filepath, List<StudentNewModel> st)
        {



            String fileName = "Users/kavyaarora/Downloads/CSV-2 3/Content/Images/MyImage.jpg";
            using (WordprocessingDocument wordDocument =
           WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();

                Body body = mainPart.Document.AppendChild(new Body());



                foreach (var stu in st)
                {



                    Paragraph para = body.AppendChild(new Paragraph());


                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Id: " + stu.Id));

                    Paragraph para0 = body.AppendChild(new Paragraph());


                    Run run0 = para0.AppendChild(new Run());
                    run0.AppendChild(new Text("Username: " + stu.Username));


                    Paragraph para1 = body.AppendChild(new Paragraph());

                    Run run1 = para1.AppendChild(new Run());
                    run1.AppendChild(new Text("Name: " + stu.Name));

                    Paragraph para2 = body.AppendChild(new Paragraph());

                    Run run2 = para2.AppendChild(new Run());
                    run2.AppendChild(new Text("Phone: " + stu.Phone));

                    Paragraph para3 = body.AppendChild(new Paragraph());

                    Run run3 = para3.AppendChild(new Run());
                    run3.AppendChild(new Text("Email: " + stu.Email));

                    Paragraph para4 = body.AppendChild(new Paragraph());

                    Run run4 = para4.AppendChild(new Run());
                    run4.AppendChild(new Text("Website: " + stu.Website));


                    Paragraph PageBreakParagraph = new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page }));




                    body.Append(PageBreakParagraph);

                }

                // var styles = ExtractStylesPart(filepath, true);



            }



        }
        public static void CreateWordprocessingDocument(string filepath, List<Student> st)
        {



            String fileName = Constants.Locations.ImagesFolder + "/" + Constants.Locations.ImageFile;
            using (WordprocessingDocument wordDocument =
           WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();

                Body body = mainPart.Document.AppendChild(new Body());



                foreach (var stu in st)
                {
                    Paragraph para = body.AppendChild(new Paragraph());


                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Hello , my name is " + stu.FirstName + " " + stu.LastName + "\n\n"));



                    string imageFilePath = stu.FullPathUrl + "/" + Constants.Locations.ImageFile;


                    var ImageBytes = FTP.DownloadFileBytes(imageFilePath);

                    SaveBytesToFile(fileName, ImageBytes);


                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                    using (FileStream stream = new FileStream(fileName, FileMode.Open))
                    {
                        imagePart.FeedData(stream);

                    }


                    AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));

                    Paragraph PageBreakParagraph = new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page }));


                    body.Append(PageBreakParagraph);

                }

                // var styles = ExtractStylesPart(filepath, true);



            }



        }


        // Extract the styles or stylesWithEffects part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractStylesPart(
          string fileName,
          bool getStylesWithEffectsPart = true)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            // Open the document for read access and get a reference.
            using (var document =
                WordprocessingDocument.Open(fileName, false))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the
                // stylesPart variable.
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                    stylesPart = docPart.StylesWithEffectsPart;
                else
                    stylesPart = docPart.StyleDefinitionsPart;

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                      stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                    }
                }
            }
            // Return the XDocument instance.
            return styles;
        }

        public static void SaveBytesToFile(string filename, byte[] bytesToWrite)
        {
            if (filename != null && filename.Length > 0 && bytesToWrite != null)
            {
                if (!Directory.Exists(System.IO.Path.GetDirectoryName(filename)))
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(filename));

                FileStream file = File.Create(filename);

                file.Write(bytesToWrite, 0, bytesToWrite.Length);

                file.Close();
            }

        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 50L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(element)));
        }

    }
}