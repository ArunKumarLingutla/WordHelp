using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordHelp
{
    public class WordUtility
    {
        public WordprocessingDocument wordDoc { get; set; }
        public static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(""));
            }
        }
        /// <summary>
        /// Opens a WordprocessingDocument at the specified file path.
        /// </summary>
        /// <param name="wordUtilityObj"></param>
        /// <param name="filePath"></param>
        /// <param name="isEditable"></param>
        public static void OpenWordDocument(WordUtility wordUtilityObj, string filePath, bool isEditable = true)
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
        /// Replaces all occurrences of a given text in a Word document (.docx).
        /// </summary>
        /// <param name="filePath">Full path to the Word document</param>
        /// <param name="searchText">The text to search for</param>
        /// <param name="replaceText">The text to replace with</param>
        public static void ReplaceText(WordprocessingDocument wordDoc, string searchText, string replaceText)
        {
            var body = wordDoc.MainDocumentPart.Document.Body;

            foreach (var text in body.Descendants<Text>())
            {
                if (text.Text.Contains(searchText))
                {
                    text.Text = text.Text.Replace(searchText, replaceText);
                }
            }
            wordDoc.MainDocumentPart.Document.Save();
        }

        public static void SetMarginsForAllSections(MainDocumentPart mainPart, int top = 720, int bottom = 720, int left = 720, int right = 720)
        {
            var body = mainPart.Document.Body;

            // Get all SectionProperties in the document
            var sections = body.Descendants<SectionProperties>();

            foreach (var sectProps in sections)
            {
                var margin = sectProps.GetFirstChild<PageMargin>();

                if (margin == null)
                {
                    margin = new PageMargin();
                    sectProps.Append(margin);
                }

                //margin.Top = top;
                //margin.Bottom = bottom;
                margin.Left = 720;
                margin.Right = 720;
            }
        }

        public static void MergeDocuments(string templateFile, string[] documentsToMerge, string destinationFile)
        {
            try
            {
                // Start from template
                File.Copy(templateFile, destinationFile, true);

                using (WordprocessingDocument destinationDoc = WordprocessingDocument.Open(destinationFile, true))
                {
                    var mainPart = destinationDoc.MainDocumentPart;
                    var body = mainPart.Document.Body;

                    for (int i = 0; i < documentsToMerge.Length; i++)
                    {
                        using (WordprocessingDocument srcDoc = WordprocessingDocument.Open(documentsToMerge[i], true))
                        {
                            var srcPart = srcDoc.MainDocumentPart;

                            // Copy styles and images
                            CopyStyles(srcPart, mainPart);
                            var imageMapping = CopyImageWithMapping(srcPart, mainPart);

                            // Copy headers and footers
                            CopyHeaderFooter(srcDoc, destinationDoc);

                            // Ensure each document starts on new page (except first one)
                            if (i > 0)
                            {
                                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                            }

                            // Get usable page width for images
                            int maxWidthEMU = GetUsablePageWidth(mainPart);

                            // Clone and import elements
                            foreach (var element in srcPart.Document.Body.Elements())
                            {
                                //if (element is SectionProperties) continue;

                                var clonedElement = (OpenXmlElement)element.CloneNode(true);

                                // Detach to avoid "part of a tree" error
                                clonedElement = clonedElement.CloneNode(true);

                                ForceLandscapeOnSection(clonedElement);

                                // Remap image IDs
                                ReMapImageReferences(clonedElement, imageMapping);

                                // Fix image sizes
                                FixImageSizes(clonedElement, maxWidthEMU);

                                body.AppendChild(clonedElement);
                            }
                        }
                    }
                    SetMarginsForAllSections(destinationDoc.MainDocumentPart);

                    // Save merged result
                    destinationDoc.MainDocumentPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error merging documents", ex);
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

                    //copy all styles from source to destination if not present already
                    foreach (var style in sourcePart.StyleDefinitionsPart.Styles.Elements<Style>())
                    {
                        if (!destinationPart.StyleDefinitionsPart.Styles.Elements<Style>().Any(s => s.StyleId == style.StyleId))
                        {
                            destinationPart.StyleDefinitionsPart.Styles.Append(style.CloneNode(true));
                        }
                    }

                    //// Copy styles from source to destination, will replace all styles in previous doc
                    //using (var stream = new MemoryStream())
                    //{
                    //    sourcePart.StyleDefinitionsPart.GetStream().CopyTo(stream);
                    //    stream.Position = 0;
                    //    destinationPart.StyleDefinitionsPart.FeedData(stream);
                    //}
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static Dictionary<string, string> CopyImageWithMapping(MainDocumentPart sourcePart, MainDocumentPart desttinationPart)
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

                foreach (var sourceFooterPart in sourceFooterParts)
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


        //public static void InsertDocumentWithLayout(MainDocumentPart mainPart, string filePath, bool isLandscape)
        //{
        //    if (mainPart?.Document?.Body == null)
        //        throw new ArgumentNullException(nameof(mainPart));

        //    // Section break before the inserted document
        //    SectionProperties startSectProps = new SectionProperties(
        //        new PageSize
        //        {
        //            Width = isLandscape ? (UInt32Value)16838U : (UInt32Value)11906U,
        //            Height = isLandscape ? (UInt32Value)11906U : (UInt32Value)16838U,
        //            Orient = isLandscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait
        //        },
        //        new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }
        //    );

        //    mainPart.Document.Body.Append(
        //        new Paragraph(new Run(new Break() { Type = BreakValues.Page })),
        //        new Paragraph(new Run()) { ParagraphProperties = new ParagraphProperties(startSectProps.CloneNode(true) as SectionProperties) }
        //    );

        //    // Section break to return to portrait layout
        //    SectionProperties endSectProps = new SectionProperties(
        //        new PageSize
        //        {
        //            Width = 11906U,
        //            Height = 16838U,
        //            Orient = PageOrientationValues.Portrait
        //        },
        //        new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720 }
        //    );

        //    mainPart.Document.Body.Append(
        //        new Paragraph(new Run(new Break() { Type = BreakValues.Page })),
        //        new Paragraph(new Run()) { ParagraphProperties = new ParagraphProperties(endSectProps.CloneNode(true) as SectionProperties) }
        //    );
        //}

        public static void InsertAPicture(WordprocessingDocument wordDoc, string fileName)
        {
            if (wordDoc.MainDocumentPart == null)
                throw new ArgumentNullException("MainDocumentPart is null.");

            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            long widthEmu;
            long heightEmu;

            using (System.Drawing.Image img = System.Drawing.Image.FromFile(fileName))
            {
                widthEmu = (long)(img.Width * 9525);  // px to EMUs
                heightEmu = (long)(img.Height * 9525);
            }

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToBody(wordDoc, mainPart.GetIdOfPart(imagePart), widthEmu, heightEmu);
        }

        public static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long widthEmu, long heightEmu)
        {
            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Inserted Image" },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );

            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }


        public static void ReplaceImage(WordprocessingDocument wordDoc, string newImagePath)
        {

            // Get all ImageParts in the main document
            var imageParts = wordDoc.MainDocumentPart.ImageParts;

            foreach (ImagePart imagePart in imageParts)
            {
                using (FileStream newImageStream = new FileStream(newImagePath, FileMode.Open))
                {
                    // Replace the image data in the ImagePart
                    imagePart.FeedData(newImageStream);
                }
            }

        }
        public static int GetUsablePageWidth(MainDocumentPart mainPart)
        {
            var sectProps = mainPart.Document.Body.Descendants<SectionProperties>().LastOrDefault();
            var pageSize = sectProps?.GetFirstChild<PageSize>();
            var pageMargin = sectProps?.GetFirstChild<PageMargin>();

            var pageWidthTwips = pageSize?.Width ?? 11906;   // default A4 portrait width in twips
            var leftMarginTwips = pageMargin?.Left ?? 1440;  // default 1 inch
            var rightMarginTwips = pageMargin?.Right ?? 1440;

            // Convert twips → EMUs (1 twip = 635 EMUs)
            return (int)((pageWidthTwips - leftMarginTwips - rightMarginTwips) * 635);
        }
        public static void FixImageSizes(OpenXmlElement element, int maxWidthEMU)
        {
            foreach (var extent in element.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>())
            {
                if (extent.Cx > maxWidthEMU)
                {
                    double scale = (double)maxWidthEMU / extent.Cx;
                    extent.Cx = (long)(extent.Cx * scale);
                    extent.Cy = (long)(extent.Cy * scale);
                }
            }
        }

        public static void ConvertNcToWord(string ncFilePath, string docxFilePath)
        {
            // Read all lines from the .nc file
            string[] ncLines = File.ReadAllLines(ncFilePath);

            // Create Word document
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))
            {
                // Add main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                // Add each line as a paragraph
                foreach (string line in ncLines)
                {
                    Paragraph para = new Paragraph(new Run(new Text(line)));
                    body.AppendChild(para);
                }

                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }

        public static void ForceLandscapeOnSection(OpenXmlElement element)
        {
            var sectPropsList = element.Descendants<SectionProperties>();

            foreach (var sectProps in sectPropsList)
            {
                var pageSize = sectProps.GetFirstChild<PageSize>();

                if (pageSize != null)
                {
                    // Swap width & height if portrait
                    if (pageSize.Width < pageSize.Height)
                    {
                        UInt32Value temp = pageSize.Width;
                        pageSize.Width = pageSize.Height;
                        pageSize.Height = temp;
                    }

                    pageSize.Orient = PageOrientationValues.Landscape;
                }
                else
                {
                    // If no PageSize, create one
                    sectProps.Append(new PageSize()
                    {
                        Width = 16838U,
                        Height = 11906U,
                        Orient = PageOrientationValues.Landscape
                    });
                }
            }
        }


        public static void ConvertToHTML(string strFilePath, string strOutputDirectory)
        {
            // Setup variables and file paths
            //string strFilePath = "Path to your .docx file"; // Replace with your file path
            //string strOutputDirectory = "OutputPath"; // Replace with your output path

            Directory.CreateDirectory(strOutputDirectory);

            WordprocessingDocument wdDoc = WordprocessingDocument.Open(strFilePath, true);

            // Set image directory and HTML settings
            string strImageDirectoryName = Path.Combine(strOutputDirectory, "_files");
            Directory.CreateDirectory(strImageDirectoryName);

            // Replace this block in ConvertToHTML method:
            HtmlConverterSettings settings = new HtmlConverterSettings()
            {
                ImageHandler = imageInfo =>
                {
                    // Get image bytes from Bitmap
                    byte[] imageBytes;
                    using (var ms = new MemoryStream())
                    {
                        imageInfo.Bitmap.Save(ms, imageInfo.Bitmap.RawFormat);
                        imageBytes = ms.ToArray();
                    }

                    string base64 = Convert.ToBase64String(imageBytes);

                    // Detect MIME type from extension
                    string mimeType = imageInfo.ContentType; // example: "image/png", "image/jpeg"

                    return new XElement(Xhtml.img,
                        new XAttribute(NoNamespace.src, $"data:{mimeType};base64,{base64}"),
                        imageInfo.ImgStyleAttribute,
                        imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null
                    );
                }
            };

            // Convert the document to HTML
            XElement htmlElement = HtmlConverter.ConvertToHtml(wdDoc, settings);
            wdDoc.Dispose();

            // Write to HTML file
            File.WriteAllText(Path.Combine(strOutputDirectory, "output.html"), htmlElement.ToString(), System.Text.Encoding.UTF8);

            Console.WriteLine("Conversion complete.");
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
        public static void InsertImagesWithCaptions(
        WordprocessingDocument wordDoc,
        List<string> imagePaths)
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            Body body = mainPart.Document.Body;

            const int maxImagesPerPage = 8;
            uint imageId = 1;

            for (int pageStart = 0; pageStart < imagePaths.Count; pageStart += maxImagesPerPage)
            {
                List<string> pageImages =
                    imagePaths.Skip(pageStart).Take(maxImagesPerPage).ToList();

                SectionProperties section =
                    body.Elements<SectionProperties>().LastOrDefault()
                    ?? body.AppendChild(new SectionProperties());

                PageSize pageSize =
                    section.GetFirstChild<PageSize>() ??
                    new PageSize() { Width = 12240, Height = 15840 };

                PageMargin margin =
                    section.GetFirstChild<PageMargin>() ??
                    new PageMargin()
                    {
                        Left = 1440,
                        Right = 1440,
                        Top = 1440,
                        Bottom = 1440
                    };

                long usableWidthTwips =
                    pageSize.Width - margin.Left - margin.Right;

                long usableHeightTwips =
                    pageSize.Height - margin.Top - margin.Bottom;

                long usableWidthEMU = usableWidthTwips * 635;
                long usableHeightEMU = usableHeightTwips * 635;

                int cols = 4;
                int rows = 2;

                long cellWidthEMU = usableWidthEMU / cols;
                long cellHeightEMU = usableHeightEMU / rows;

                long imageWidth = (long)(cellWidthEMU * 0.9);
                long imageHeight = (long)(cellHeightEMU * 0.7);

                DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

                table.AppendChild(
                    new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = BorderValues.None },
                            new BottomBorder { Val = BorderValues.None },
                            new LeftBorder { Val = BorderValues.None },
                            new RightBorder { Val = BorderValues.None },
                            new InsideHorizontalBorder { Val = BorderValues.None },
                            new InsideVerticalBorder { Val = BorderValues.None }
                        )));

                int index = 0;

                for (int r = 0; r < rows; r++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                    for (int c = 0; c < cols; c++)
                    {
                        DocumentFormat.OpenXml.Wordprocessing.TableCell tc = new DocumentFormat.OpenXml.Wordprocessing.TableCell();

                        if (index < pageImages.Count)
                        {
                            string imagePath = pageImages[index];

                            Drawing drawing =
                                AddImage(mainPart, imagePath, imageWidth, imageHeight, imageId++);

                            Paragraph imagePara =
                                new Paragraph(
                                    new Run(drawing))
                                {
                                    ParagraphProperties =
                                        new ParagraphProperties(
                                            new Justification()
                                            {
                                                Val = JustificationValues.Center
                                            })
                                };

                            string caption =
                                Path.GetFileNameWithoutExtension(imagePath);

                            Paragraph captionPara =
                                new Paragraph(
                                    new Run(
                                        new Text(caption)))
                                {
                                    ParagraphProperties =
                                        new ParagraphProperties(
                                            new Justification()
                                            {
                                                Val = JustificationValues.Center
                                            })
                                };

                            tc.Append(imagePara);
                            tc.Append(captionPara);

                            index++;
                        }

                        tc.Append(new TableCellProperties(
                            new TableCellWidth()
                            {
                                Type = TableWidthUnitValues.Dxa,
                                Width = (usableWidthTwips / cols).ToString()
                            }));

                        tr.Append(tc);
                    }

                    table.Append(tr);
                }

                body.Append(table);

                if (pageStart + maxImagesPerPage < imagePaths.Count)
                {
                    body.Append(
                        new Paragraph(
                            new Run(
                                new Break()
                                {
                                    Type = BreakValues.Page
                                })));
                }
            }
        }

        private static Drawing AddImage(
            MainDocumentPart mainPart,
            string imagePath,
            long width,
            long height,
            UInt32 imageId)
        {
            //var type = ImagePartType.Jpeg;

            //string ext = Path.GetExtension(imagePath).ToLower();
            var type = ImagePartType.Jpeg;

            string ext = Path.GetExtension(imagePath).ToLower();

            switch (ext)
            {
                case ".png":
                    type = ImagePartType.Png;
                    break;

                case ".bmp":
                    type = ImagePartType.Bmp;
                    break;

                case ".gif":
                    type = ImagePartType.Gif;
                    break;

                case ".tiff":
                case ".tif":
                    type = ImagePartType.Tiff;
                    break;

                case ".jpeg":
                case ".jpg":
                default:
                    type = ImagePartType.Jpeg;
                    break;
            }
            if (ext == ".png") type = ImagePartType.Png;
            if (ext == ".bmp") type = ImagePartType.Bmp;

            ImagePart imagePart = mainPart.AddImagePart(type);

            using (FileStream stream =
                   new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string relId = mainPart.GetIdOfPart(imagePart);

            return new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent()
                    {
                        Cx = width,
                        Cy = height
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                    {
                        Id = imageId,
                        Name = "Picture"
                    },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                    {
                                        Id = 0U,
                                        Name = "Image"
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip()
                                    {
                                        Embed = relId
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset()
                                        {
                                            X = 0,
                                            Y = 0
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Extents()
                                        {
                                            Cx = width,
                                            Cy = height
                                        }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                    {
                                        Preset =
                                        DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                                    })
                            )
                        )
                        {
                            Uri =
                            "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    )
                )
            );
        }
        //public static void InsertImagesWithCaptions(DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDoc, List<string> imagePaths)
        //{
        //    DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart = wordDoc.MainDocumentPart;

        //    DocumentFormat.OpenXml.Wordprocessing.Body body =
        //        mainPart.Document.Body;

        //    //DocumentFormat.OpenXml.Wordprocessing.SectionProperties section =
        //    //    body.Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().LastOrDefault();

        //    SectionProperties section =body.Elements<SectionProperties>().LastOrDefault()?? body.AppendChild(new SectionProperties());

        //    DocumentFormat.OpenXml.Wordprocessing.PageSize pageSize =
        //        section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();

        //    DocumentFormat.OpenXml.Wordprocessing.PageMargin margin =
        //        section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageMargin>();

        //    long usableWidthTwips =
        //        pageSize.Width - margin.Left - margin.Right;

        //    long usableHeightTwips =
        //        pageSize.Height - margin.Top - margin.Bottom;

        //    long usableWidthEMU = usableWidthTwips * 635;
        //    long usableHeightEMU = usableHeightTwips * 635;

        //    int imageCount = imagePaths.Count;

        //    int cols = (int)Math.Ceiling(Math.Sqrt(imageCount));
        //    int rows = (int)Math.Ceiling((double)imageCount / cols);

        //    long cellWidthEMU = usableWidthEMU / cols;
        //    long cellHeightEMU = usableHeightEMU / rows;

        //    long imageHeight = (long)(cellHeightEMU * 0.7);
        //    long imageWidth = (long)(cellWidthEMU * 0.9);

        //    DocumentFormat.OpenXml.Wordprocessing.Table table =
        //        new DocumentFormat.OpenXml.Wordprocessing.Table();

        //    table.AppendChild(
        //        new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
        //            new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
        //                new DocumentFormat.OpenXml.Wordprocessing.TopBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None },
        //                new DocumentFormat.OpenXml.Wordprocessing.BottomBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None },
        //                new DocumentFormat.OpenXml.Wordprocessing.LeftBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None },
        //                new DocumentFormat.OpenXml.Wordprocessing.RightBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None },
        //                new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None },
        //                new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder
        //                { Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.None }
        //            )));

        //    int index = 0;
        //    uint imageId = 1;

        //    for (int r = 0; r < rows; r++)
        //    {
        //        DocumentFormat.OpenXml.Wordprocessing.TableRow tr =
        //            new DocumentFormat.OpenXml.Wordprocessing.TableRow();

        //        for (int c = 0; c < cols; c++)
        //        {
        //            DocumentFormat.OpenXml.Wordprocessing.TableCell tc =
        //                new DocumentFormat.OpenXml.Wordprocessing.TableCell();

        //            if (index < imageCount)
        //            {
        //                string imagePath = imagePaths[index];

        //                DocumentFormat.OpenXml.Wordprocessing.Drawing drawing =
        //                    AddImage(mainPart, imagePath, imageWidth, imageHeight, imageId++);

        //                DocumentFormat.OpenXml.Wordprocessing.Paragraph imagePara =
        //                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
        //                        new DocumentFormat.OpenXml.Wordprocessing.Run(drawing))
        //                    {
        //                        ParagraphProperties =
        //                            new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
        //                                new DocumentFormat.OpenXml.Wordprocessing.Justification
        //                                {
        //                                    Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center
        //                                })
        //                    };

        //                string caption =
        //                    System.IO.Path.GetFileNameWithoutExtension(imagePath);

        //                DocumentFormat.OpenXml.Wordprocessing.Paragraph captionPara =
        //                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
        //                        new DocumentFormat.OpenXml.Wordprocessing.Run(
        //                            new DocumentFormat.OpenXml.Wordprocessing.Text(caption)))
        //                    {
        //                        ParagraphProperties =
        //                            new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
        //                                new DocumentFormat.OpenXml.Wordprocessing.Justification
        //                                {
        //                                    Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center
        //                                })
        //                    };

        //                tc.Append(imagePara);
        //                tc.Append(captionPara);

        //                index++;
        //            }

        //            tc.Append(new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
        //                new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth
        //                {
        //                    Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa,
        //                    Width = (usableWidthTwips / cols).ToString()
        //                }));

        //            tr.Append(tc);
        //        }

        //        table.Append(tr);
        //    }

        //    body.Append(table);
        //}
        //private static DocumentFormat.OpenXml.Wordprocessing.Drawing AddImage(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, string imagePath, long width, long height, UInt32 imageId)
        //{
        //    var type =
        //        DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg;

        //    string ext = System.IO.Path.GetExtension(imagePath).ToLower();

        //    if (ext == ".png")
        //        type = DocumentFormat.OpenXml.Packaging.ImagePartType.Png;

        //    if (ext == ".bmp")
        //        type = DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp;

        //    DocumentFormat.OpenXml.Packaging.ImagePart imagePart =
        //        mainPart.AddImagePart(type);

        //    using (System.IO.FileStream stream =
        //           new System.IO.FileStream(imagePath, System.IO.FileMode.Open))
        //    {
        //        imagePart.FeedData(stream);
        //    }

        //    string relId = mainPart.GetIdOfPart(imagePart);

        //    return new DocumentFormat.OpenXml.Wordprocessing.Drawing(
        //        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
        //            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
        //            {
        //                Cx = width,
        //                Cy = height
        //            },
        //            new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
        //            {
        //                Id = (UInt32Value)imageId,
        //                Name = "Picture"
        //            },
        //            new DocumentFormat.OpenXml.Drawing.Graphic(
        //                new DocumentFormat.OpenXml.Drawing.GraphicData(
        //                    new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
        //                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
        //                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
        //                            {
        //                                Id = 0U,
        //                                Name = "Image"
        //                            },
        //                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
        //                        ),
        //                        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
        //                            new DocumentFormat.OpenXml.Drawing.Blip
        //                            {
        //                                Embed = relId
        //                            },
        //                            new DocumentFormat.OpenXml.Drawing.Stretch(
        //                                new DocumentFormat.OpenXml.Drawing.FillRectangle())
        //                        ),
        //                        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
        //                            new DocumentFormat.OpenXml.Drawing.Transform2D(
        //                                new DocumentFormat.OpenXml.Drawing.Offset
        //                                {
        //                                    X = 0,
        //                                    Y = 0
        //                                },
        //                                new DocumentFormat.OpenXml.Drawing.Extents
        //                                {
        //                                    Cx = width,
        //                                    Cy = height
        //                                }),
        //                            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
        //                                new DocumentFormat.OpenXml.Drawing.AdjustValueList())
        //                            {
        //                                Preset =
        //                                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
        //                            })
        //                    )
        //                )
        //                {
        //                    Uri =
        //                    "http://schemas.openxmlformats.org/drawingml/2006/picture"
        //                }
        //            )
        //        )
        //    );
        //}

    }
}
