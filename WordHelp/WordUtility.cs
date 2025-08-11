using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordHelp
{
    public class WordUtility
    {
        public WordprocessingDocument wordDoc { get; set; }

        /// <summary>
        /// Opens a WordprocessingDocument at the specified file path.
        /// </summary>
        /// <param name="wordUtilityObj"></param>
        /// <param name="filePath"></param>
        /// <param name="isEditable"></param>
        public static void OpenWordDocument(WordUtility wordUtilityObj,string filePath, bool isEditable = true)
        {
            try
            {
                wordUtilityObj.wordDoc = WordprocessingDocument.Open(filePath, isEditable);
            }
            catch (Exception ex)
            {
                CloseWordProcessDocument(wordUtilityObj.wordDoc);
                throw ex;
            }
        }

        /// <summary>
        /// Finds and replaces text in a WordprocessingDocument.
        /// </summary>
        /// <param name="wordprocessingDocument"></param>
        /// <param name="searchText"></param>
        /// <param name="replacementText"></param>
        public static void FindAndReplaceText(WordprocessingDocument wordprocessingDocument, string searchText, string replacementText)
        {
            //// Open the WordprocessingDocument for editing
            //using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, true))
            //{
                // Access the MainDocumentPart and make sure it is not null
                var mainDocumentPart = wordprocessingDocument.MainDocumentPart;

                if (mainDocumentPart != null)
                {
                    // Create a MemoryStream to store the updated MainDocumentPart
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        // Create an OpenXmlReader to read the main document part
                        // and an OpenXmlWriter to write to the MemoryStream
                        using (OpenXmlReader reader = OpenXmlPartReader.Create(mainDocumentPart))
                        using (OpenXmlWriter writer = OpenXmlPartWriter.Create(memoryStream))
                        {
                            // Write the XML declaration with the version "1.0".
                            writer.WriteStartDocument();

                            // Read the elements from the MainDocumentPart
                            while (reader.Read())
                            {
                                // Check if the element is of type Text
                                if (reader.ElementType == typeof(Text))
                                {
                                    // If it is the start of an element write the start element and the updated text
                                    if (reader.IsStartElement)
                                    {
                                        writer.WriteStartElement(reader);

                                        string text = reader.GetText().Replace(searchText, replacementText);

                                        writer.WriteString(text);

                                    }
                                    else
                                    {
                                        // Close the element
                                        writer.WriteEndElement();
                                    }
                                }
                                else
                                // Write the other XML elements without editing
                                {
                                    if (reader.IsStartElement)
                                    {
                                        writer.WriteStartElement(reader);
                                    }
                                    else if (reader.IsEndElement)
                                    {
                                        writer.WriteEndElement();
                                    }
                                }
                            }
                        }
                        // Set the MemoryStream's position to 0 and replace the MainDocumentPart
                        memoryStream.Position = 0;
                        mainDocumentPart.FeedData(memoryStream);
                    }
                }
            //}
        }

        public static void MergeDocuments(string[] documentsToMerge, string destinationFile)
        {
            try
            {
                //Copy the first document as base doc
                File.Copy(documentsToMerge[0], destinationFile);

                using (WordprocessingDocument destinationDoc = WordprocessingDocument.Open(destinationFile, true))
                {
                    var mainPart = destinationDoc.MainDocumentPart;
                    var body = mainPart.Document.Body;

                    for (int i = 1; i < documentsToMerge.Length; i++)
                    {
                        using(WordprocessingDocument srcDoc = WordprocessingDocument.Open(documentsToMerge[i], true))
                        {
                            var srcPart = srcDoc.MainDocumentPart;

                            CopyStyles(srcPart, mainPart);

                            var imageMapping = CopyImageWithMapping(srcPart, mainPart);

                            CopyHeaderFooter(srcDoc, destinationDoc);

                            //clone and remap body content
                            foreach (var element in srcPart.Document.Body.Elements())
                            {
                                var clonedElement = element.CloneNode(true);
                                ReMapImageReferences(clonedElement, imageMapping);
                                body.AppendChild(clonedElement);
                            }

                            body.AppendChild(new Paragraph(new Run(new Break()))); // Add a break between documents
                        }
                    }
                    // Save changes to the destination document
                    destinationDoc.MainDocumentPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void CopyStyles(MainDocumentPart sourcePart, MainDocumentPart destinationPart)
        {
            try
            {
                if (sourcePart.StyleDefinitionsPart != null)
                {
                    // Ensure the destination part has a StyleDefinitionsPart
                    if (destinationPart.StyleDefinitionsPart == null)
                    {
                        destinationPart.AddNewPart<StyleDefinitionsPart>();
                    }
                    // Copy styles from source to destination
                    using (var stream = new MemoryStream())
                    {
                        sourcePart.StyleDefinitionsPart.GetStream().CopyTo(stream);
                        stream.Position = 0;
                        destinationPart.StyleDefinitionsPart.FeedData(stream);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static Dictionary<string, string> CopyImageWithMapping(MainDocumentPart sourcePart,MainDocumentPart desttinationPart)
        {
            var imageMapping = new Dictionary<string, string>();
            try
            {
                foreach (var imagePart in sourcePart.ImageParts)
                {
                    string oldRelId = sourcePart.GetIdOfPart(imagePart);

                    var newImagePart = desttinationPart.AddImagePart(imagePart.ContentType);
                    newImagePart.FeedData(imagePart.GetStream());

                    string newRelId = desttinationPart.GetIdOfPart(newImagePart);
                    imageMapping[oldRelId] = newRelId;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return imageMapping;
        }

        public static void ReMapImageReferences(OpenXmlElement element, Dictionary<string, string> imageMapping)
        {
            try
            {
                foreach (var drawing in element.Descendants<Drawing>())
                {
                    var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                    if (blip != null)
                    {
                        var oldRelId = blip.Embed.Value;
                        if (imageMapping.ContainsKey(oldRelId))
                        {
                            blip.Embed = imageMapping[blip.Embed]; 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void CopyHeaderFooter(WordprocessingDocument sourcePart, WordprocessingDocument destinationPart)
        {
            try
            {
                var sourceHeaderParts = sourcePart.MainDocumentPart.HeaderParts;
                var sourceFooterParts = sourcePart.MainDocumentPart.FooterParts;

                foreach (var sourceHeaderPart in sourceHeaderParts)
                {
                    var newHeaderPart = destinationPart.MainDocumentPart.AddNewPart<HeaderPart>();
                    newHeaderPart.FeedData(sourceHeaderPart.GetStream());
                }

                foreach(var sourceFooterPart in sourceFooterParts)
                {
                    var newFooterPart = destinationPart.MainDocumentPart.AddNewPart<FooterPart>();
                    newFooterPart.FeedData(sourceFooterPart.GetStream());
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void InsertAPicture(string document, string fileName)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
            {
                if (wordprocessingDocument.MainDocumentPart is null)
                {
                    throw new ArgumentNullException("MainDocumentPart is null.");
                }

                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
            }
        }

        public static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
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

            if (wordDoc.MainDocumentPart is null || wordDoc.MainDocumentPart.Document.Body is null)
            {
                throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
            }

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        public static void SaveWordProcessDocument(WordprocessingDocument wordprocessingDocument)
        {
            try
            {
                wordprocessingDocument.MainDocumentPart.Document.Save();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void SaveAs(WordprocessingDocument wordprocessingDocument, string filePath)
        {
            try
            {
                wordprocessingDocument.Clone(filePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void CloseWordProcessDocument(WordprocessingDocument wordprocessingDocument)
        {
            try
            {
                wordprocessingDocument.Dispose();
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
