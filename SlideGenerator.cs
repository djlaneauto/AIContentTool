using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AIContentTool
{
    public static class SlideGenerator
    {
        /// <summary>
        /// Generates slides in the specified PowerPoint presentation based on the content and format.
        /// </summary>
        /// <param name="presentation">The PowerPoint presentation to add slides to.</param>
        /// <param name="content">The content string in XML, JSON, or Markdown format.</param>
        /// <param name="format">The format of the content ("XML", "JSON", or "MARKDOWN").</param>
        /// <returns>The total number of slides after generation.</returns>
        public static int GenerateSlides(PowerPoint.Presentation presentation, string content, string format)
        {
            try
            {
                DateTime startTime = DateTime.Now;

                // Apply the selected template if specified
                if (!string.IsNullOrEmpty(ImportHandlers.TemplateFilePath))
                {
                    try
                    {
                        presentation.ApplyTemplate(ImportHandlers.TemplateFilePath);
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.LogError($"Failed to apply template: {ex.Message}");
                        MessageBox.Show("Error applying template. Check the log for details.");
                    }
                }

                switch (format.ToUpper())
                {
                    case "XML":
                        GenerateFromXml(presentation, content);
                        break;
                    case "JSON":
                        GenerateFromJson(presentation, content);
                        break;
                    case "MARKDOWN":
                        GenerateFromMarkdown(presentation, content);
                        break;
                    default:
                        throw new ArgumentException("Unsupported content format.");
                }

                int slideCount = presentation.Slides.Count;
                KPITracker.RecordGeneration(slideCount, (DateTime.Now - startTime).TotalSeconds);
                return slideCount;
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"Slide generation failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Generates slides from XML content.
        /// </summary>
        private static void GenerateFromXml(PowerPoint.Presentation presentation, string xmlContent)
        {
            var xmlDoc = XDocument.Parse(xmlContent);
            var slides = xmlDoc.Descendants("slide");
            foreach (var slideElement in slides)
            {
                string title = slideElement.Element("title")?.Value ?? "Untitled";
                var contentElements = slideElement.Element("content")?.Elements();
                if (contentElements != null)
                {
                    AddSlideWithContentElements(presentation, title, contentElements);
                }
            }
        }

        /// <summary>
        /// Generates slides from JSON content.
        /// </summary>
        private static void GenerateFromJson(PowerPoint.Presentation presentation, string jsonContent)
        {
            try
            {
                var slides = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonContent);
                foreach (var slide in slides)
                {
                    string title = slide.ContainsKey("title") ? slide["title"].ToString() : "Untitled";
                    var contentElements = new List<XElement>();
                    if (slide.ContainsKey("content") && slide["content"] is Newtonsoft.Json.Linq.JArray contentArray)
                    {
                        foreach (var item in contentArray)
                        {
                            var contentDict = item.ToObject<Dictionary<string, string>>();
                            string type = contentDict.ContainsKey("type") ? contentDict["type"] : "p";
                            string value = contentDict["value"];
                            contentElements.Add(new XElement(type, value));
                        }
                    }
                    AddSlideWithContentElements(presentation, title, contentElements);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"JSON parsing failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Generates slides from Markdown content (basic support for titles and paragraphs).
        /// </summary>
        private static void GenerateFromMarkdown(PowerPoint.Presentation presentation, string mdContent)
        {
            var lines = mdContent.Split('\n');
            string title = null;
            var contentElements = new List<XElement>();
            foreach (var line in lines)
            {
                if (line.StartsWith("# "))
                {
                    if (title != null)
                    {
                        AddSlideWithContentElements(presentation, title, contentElements);
                        contentElements = new List<XElement>();
                    }
                    title = line.Substring(2).Trim();
                }
                else if (!string.IsNullOrWhiteSpace(line))
                {
                    contentElements.Add(new XElement("p", line.Trim()));
                }
            }
            if (title != null)
            {
                AddSlideWithContentElements(presentation, title, contentElements);
            }
        }

        /// <summary>
        /// Adds a slide with the specified title and content elements, handling text and images.
        /// </summary>
        private static void AddSlideWithContentElements(PowerPoint.Presentation presentation, string title, IEnumerable<XElement> contentElements)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            slide.Shapes.Title.TextFrame.TextRange.Text = title ?? "Untitled";
            var contentShape = slide.Shapes[2];
            var textRange = contentShape.TextFrame.TextRange;

            // Parse text-related elements (exclude images)
            var textElements = contentElements.Where(e => e.Name != "image");
            ParseContentElements(textRange, textElements, 0);

            // Add images below the content shape
            float currentTop = contentShape.Top + contentShape.Height + 10;
            foreach (var imageElement in contentElements.Where(e => e.Name == "image"))
            {
                string placeholder = imageElement.Value;
                if (ImportHandlers.PlaceholderFiles.TryGetValue(placeholder, out string filePath) && File.Exists(filePath))
                {
                    slide.Shapes.AddPicture(filePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                        contentShape.Left, currentTop, contentShape.Width, 100);
                    currentTop += 110;
                }
            }
        }

        /// <summary>
        /// Parses content elements (paragraphs and lists) into the text range.
        /// </summary>
        private static void ParseContentElements(PowerPoint.TextRange textRange, IEnumerable<XElement> elements, int indentLevel)
        {
            foreach (var element in elements)
            {
                if (element.Name == "p")
                {
                    ParseParagraph(textRange, element);
                }
                else if (element.Name == "list")
                {
                    ParseList(textRange, element, indentLevel);
                }
            }
        }

        /// <summary>
        /// Parses a paragraph element, applying inline text formatting.
        /// </summary>
        private static void ParseParagraph(PowerPoint.TextRange textRange, XElement paragraphElement)
        {
            var textElements = paragraphElement.Elements("text");
            if (textElements.Any())
            {
                foreach (var textElement in textElements)
                {
                    string text = textElement.Value;
                    string style = textElement.Attribute("style")?.Value ?? "";
                    string fontSize = textElement.Attribute("font-size")?.Value ?? "12";
                    var textRun = textRange.InsertAfter(text + " ");
                    ApplyTextFormatting(textRun, style, fontSize);
                }
                textRange.InsertAfter("\n");
            }
            else
            {
                textRange.InsertAfter(paragraphElement.Value + "\n");
            }
        }

        /// <summary>
        /// Parses a list element, handling nested lists and bullet/numbering.
        /// </summary>
        private static void ParseList(PowerPoint.TextRange textRange, XElement listElement, int indentLevel)
        {
            string listType = listElement.Attribute("type")?.Value ?? "bullet";
            foreach (var itemElement in listElement.Elements("item"))
            {
                var itemTextRange = textRange.InsertAfter("");
                itemTextRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                itemTextRange.ParagraphFormat.Bullet.Type = listType == "bullet" ?
                    PowerPoint.PpBulletType.ppBulletUnnumbered : PowerPoint.PpBulletType.ppBulletNumbered;
                itemTextRange.IndentLevel = indentLevel + 1;
                ParseContentElements(itemTextRange, itemElement.Elements(), indentLevel + 1);
                textRange.InsertAfter("\n");
            }
        }

        /// <summary>
        /// Applies text formatting (bold, underline, font size) to a text range.
        /// </summary>
        private static void ApplyTextFormatting(PowerPoint.TextRange textRun, string style, string fontSize)
        {
            if (style.Contains("bold"))
            {
                textRun.Font.Bold = Office.MsoTriState.msoTrue;
            }
            if (style.Contains("underline"))
            {
                textRun.Font.Underline = Office.MsoTriState.msoTrue;
            }
            if (int.TryParse(fontSize, out int size))
            {
                textRun.Font.Size = size;
            }
        }
    }
}