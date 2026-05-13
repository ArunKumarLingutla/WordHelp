using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WordHelp
{
    public class DocxConverter
    {
        public static void ConvertToHtml(string docxPath, string outputHtmlPath)
        {
            var hyperlinkTargets = LoadHyperlinkTargets(docxPath);

            using (var doc = WordprocessingDocument.Open(docxPath, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var html = new XElement("html",
                    new XElement("head",
                        new XElement("meta", new XAttribute("charset", "utf-8"))
                    ),
                    new XElement("body",
                        body.Elements()
                            .Select(e => ConvertElement(e, doc, hyperlinkTargets))
                            .Where(e => e != null)
                    )
                );

                File.WriteAllText(outputHtmlPath, html.ToString(), Encoding.UTF8);
            }
        }

        private static object ConvertElement(OpenXmlElement element, WordprocessingDocument doc, Dictionary<string, string> hyperlinkTargets)
        {
            if (element is Word.Paragraph para)
            {
                double marginTop = 0, marginBottom = 0, lineHeight = 0;

                var pPr = para.ParagraphProperties;
                if (pPr?.SpacingBetweenLines != null)
                {
                    var spacing = pPr.SpacingBetweenLines;

                    if (spacing.Before != null)
                        marginTop = ConvertTwipToPx(int.Parse(spacing.Before));

                    if (spacing.After != null)
                        marginBottom = ConvertTwipToPx(int.Parse(spacing.After));

                    if (spacing.Line != null)
                    {
                        double line = int.Parse(spacing.Line);
                        lineHeight = Math.Round(line / 240.0, 2);
                    }
                }

                string style = $"margin:{marginTop}px 0 {marginBottom}px 0;";
                if (lineHeight > 0)
                    style += $"line-height:{lineHeight}em;";

                return new XElement("p",
                    new XAttribute("style", style),
                    para.Elements().Select((OpenXmlElement e) =>
                    {
                        if (e is Word.Run r) return (object)ConvertRun(r, doc);
                        if (e is Word.Hyperlink h) return (object)ConvertHyperlink(h, doc, hyperlinkTargets);
                        return null;
                    }).Where(e => e != null)
                );
            }
            else if (element is Word.Hyperlink hyperlink)
            {
                return ConvertHyperlink(hyperlink, doc, hyperlinkTargets);
            }
            else if (element is Word.Table table)
            {
                var tblPr = table.GetFirstChild<Word.TableProperties>();
                var tblW = tblPr?.GetFirstChild<Word.TableWidth>();

                string tableStyle = "border-collapse:collapse;margin:10px 0;table-layout:fixed;";
                if (tblW != null)
                {
                    if (tblW.Type == Word.TableWidthUnitValues.Dxa)
                    {
                        double px = ConvertTwipToPx(int.Parse(tblW.Width));
                        tableStyle += $"width:{px}px;";
                    }
                    else if (tblW.Type == Word.TableWidthUnitValues.Pct)
                    {
                        double percent = int.Parse(tblW.Width) / 50.0;
                        tableStyle += $"width:{percent}%;";
                    }
                }

                return new XElement("table",
                    new XAttribute("border", "1"),
                    new XAttribute("cellpadding", "4"),
                    new XAttribute("style", tableStyle),
                    table.Elements<Word.TableRow>().Select(tr =>
                        new XElement("tr",
                            tr.Elements<Word.TableCell>().Select(tc =>
                            {
                                string style = "border:1px solid black;padding:4px;";
                                var tcPr = tc.GetFirstChild<Word.TableCellProperties>();
                                var tcW = tcPr?.GetFirstChild<Word.TableCellWidth>();

                                if (tcW != null)
                                {
                                    if (tcW.Type == Word.TableWidthUnitValues.Dxa)
                                    {
                                        double px = ConvertTwipToPx(int.Parse(tcW.Width));
                                        style += $"width:{px}px;";
                                    }
                                    else if (tcW.Type == Word.TableWidthUnitValues.Pct)
                                    {
                                        double percent = int.Parse(tcW.Width) / 50.0;
                                        style += $"width:{percent}%;";
                                    }
                                }

                                return new XElement("td",
                                    new XAttribute("style", style),
                                    (tc.TableCellProperties?.GridSpan != null) ?
                                        new XAttribute("colspan", tc.TableCellProperties.GridSpan.Val) : null,
                                    (tc.TableCellProperties?.VerticalMerge != null &&
                                     tc.TableCellProperties.VerticalMerge.Val != null &&
                                     tc.TableCellProperties.VerticalMerge.Val != Word.MergedCellValues.Restart) ?
                                        new XAttribute("rowspan", "???") : null,
                                    tc.Elements().Select(e => ConvertElement(e, doc, hyperlinkTargets)).Where(e => e != null)
                                );
                            })
                        )
                    )
                );
            }

            return null;
        }

        private static double ConvertTwipToPx(int twips)
        {
            return Math.Round(twips * 96.0 / 1440.0, 2);
        }

        private static object ConvertRun(Word.Run run, WordprocessingDocument doc)
        {
            var text = run.Elements<Word.Text>().FirstOrDefault()?.Text;
            XElement result = null;

            if (!string.IsNullOrEmpty(text))
            {
                object formatted = new XText(text);

                if (run.RunProperties != null)
                {
                    if (run.RunProperties.Bold != null)
                        formatted = new XElement("strong", formatted);

                    if (run.RunProperties.Italic != null)
                        formatted = new XElement("em", formatted);

                    if (run.RunProperties.Underline != null &&
                        run.RunProperties.Underline.Val != Word.UnderlineValues.None)
                        formatted = new XElement("u", formatted);
                }

                result = formatted as XElement ?? new XElement("span", formatted);
            }

            var drawing = run.Elements<Word.Drawing>().FirstOrDefault();
            if (drawing != null)
                return ConvertImage(drawing, doc);

            return result;
        }

        private static object ConvertImage(Word.Drawing drawing, WordprocessingDocument doc)
        {
            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (blip == null) return null;

            var embed = blip.Embed?.Value;
            if (embed == null) return null;

            var part = (ImagePart)doc.MainDocumentPart.GetPartById(embed);

            var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
            double widthPx = 0, heightPx = 0;
            if (extent != null)
            {
                widthPx = Math.Round(extent.Cx / 914400.0 * 96);
                heightPx = Math.Round(extent.Cy / 914400.0 * 96);
            }

            using (var stream = part.GetStream())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                var base64 = Convert.ToBase64String(ms.ToArray());
                var mime = part.ContentType;

                var img = new XElement("img",
                    new XAttribute("src", $"data:{mime};base64,{base64}")
                );

                if (widthPx > 0 && heightPx > 0)
                {
                    img.SetAttributeValue("width", $"{widthPx}px");
                    img.SetAttributeValue("height", $"{heightPx}px");
                    img.SetAttributeValue("style", $"width:{widthPx}px;height:{heightPx}px;");
                }

                return img;
            }
        }

        private static XElement ConvertHyperlink(Word.Hyperlink hyperlink,
            WordprocessingDocument doc, Dictionary<string, string> hyperlinkTargets)
        {
            string url = "#";

            if (hyperlink.Id != null &&
                hyperlinkTargets.TryGetValue(hyperlink.Id.Value, out var target))
            {
                if (target.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                    url = target.Replace("\\", "/");
                else if (target.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    url = target;
                else if (Path.IsPathRooted(target))
                    url = "file:///" + target.Replace("\\", "/");
                else
                    url = target;
            }
            else if (hyperlink.Anchor != null)
            {
                url = $"#{hyperlink.Anchor.Value}";
            }

            var runs = hyperlink.Elements<Word.Run>()
                .Select(r => ConvertRun(r, doc))
                .Where(r => r != null)
                .ToList();

            if (!runs.Any())
            {
                var fallbackText = hyperlink.InnerText;
                if (!string.IsNullOrEmpty(fallbackText))
                    runs.Add(new XElement("span", fallbackText));
            }

            if (!runs.Any() && url == "#")
                return null;

            return new XElement("a",
                new XAttribute("href", url),
                runs
            );
        }

        private static Dictionary<string, string> LoadHyperlinkTargets(string docxPath)
        {
            var result = new Dictionary<string, string>();

            try
            {
                using (var zip = System.IO.Compression.ZipFile.OpenRead(docxPath))
                {
                    var relsEntry = zip.GetEntry("word/_rels/document.xml.rels");
                    if (relsEntry == null)
                        return result;

                    using (var stream = relsEntry.Open())
                    using (var reader = new StreamReader(stream))
                    {
                        var xml = reader.ReadToEnd();
                        var xdoc = XDocument.Parse(xml);
                        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

                        foreach (var rel in xdoc.Descendants(ns + "Relationship"))
                        {
                            var id = rel.Attribute("Id")?.Value;
                            var target = rel.Attribute("Target")?.Value;

                            if (id != null && target != null)
                                result[id] = Uri.UnescapeDataString(target);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or rethrow as needed
                throw new InvalidOperationException($"Failed to load hyperlink targets from: {docxPath}", ex);
            }

            return result;
        }
    }
}