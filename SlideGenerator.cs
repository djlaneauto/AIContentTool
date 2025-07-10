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
                var elements = slideElement.Elements().Where(e => e.Name != "title");
                AddSlideWithElements(presentation, title, elements);
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
                    var elements = new List<XElement>();
                    if (slide.ContainsKey("elements") && slide["elements"] is Newtonsoft.Json.Linq.JArray elementsArray)
                    {
                        foreach (var item in elementsArray)
                        {
                            var elemDict = item.ToObject<Dictionary<string, object>>();
                            string type = elemDict.ContainsKey("type") ? elemDict["type"].ToString() : "textbox";
                            var elem = new XElement(type);
                            if (elemDict.ContainsKey("attributes"))
                            {
                                var attrs = (Dictionary<string, string>)elemDict["attributes"];
                                foreach (var attr in attrs)
                                {
                                    elem.SetAttributeValue(attr.Key, attr.Value);
                                }
                            }
                            if (elemDict.ContainsKey("content"))
                            {
                                elem.Value = elemDict["content"].ToString();
                                // For complex content like lists/tables, parse recursively if needed
                            }
                            // TODO: Enhance for nested structures if JSON provides them
                            elements.Add(elem);
                        }
                    }
                    AddSlideWithElements(presentation, title, elements);
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
            var elements = new List<XElement>();
            foreach (var line in lines)
            {
                if (line.StartsWith("# "))
                {
                    if (title != null)
                    {
                        AddSlideWithElements(presentation, title, elements);
                        elements = new List<XElement>();
                    }
                    title = line.Substring(2).Trim();
                }
                else if (!string.IsNullOrWhiteSpace(line))
                {
                    elements.Add(new XElement("textbox", new XAttribute("left", "100"), new XAttribute("top", "200"), new XAttribute("width", "500"), new XAttribute("height", "100"), line.Trim()));
                }
            }
            if (title != null)
            {
                AddSlideWithElements(presentation, title, elements);
            }
        }

        /// <summary>
        /// Adds a slide with the specified title and elements (textboxes, lists, tables, charts, images).
        /// </summary>
        private static void AddSlideWithElements(PowerPoint.Presentation presentation, string title, IEnumerable<XElement> elements)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank); // Use blank for full control
            if (!string.IsNullOrEmpty(title))
            {
                var titleShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 20, 600, 50);
                titleShape.TextFrame.TextRange.Text = title;
                titleShape.TextFrame.TextRange.Font.Size = 32;
                titleShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            }

            foreach (var element in elements)
            {
                switch (element.Name.LocalName)
                {
                    case "textbox":
                        AddTextbox(slide, element);
                        break;
                    case "list":
                        AddList(slide, element);
                        break;
                    case "table":
                        AddTable(slide, element);
                        break;
                    case "chart":
                        AddChart(slide, element);
                        break;
                    case "image":
                        AddImage(slide, element);
                        break;
                    default:
                        // Ignore unknown elements
                        break;
                }
            }
        }

        /// <summary>
        /// Adds a textbox to the slide based on the XML element.
        /// </summary>
        private static void AddTextbox(PowerPoint.Slide slide, XElement textboxElement)
        {
            float left = float.TryParse(textboxElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(textboxElement.Attribute("top")?.Value, out var t) ? t : 100;
            float width = float.TryParse(textboxElement.Attribute("width")?.Value, out var w) ? w : 400;
            float height = float.TryParse(textboxElement.Attribute("height")?.Value, out var h) ? h : 200;

            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            var textRange = shape.TextFrame.TextRange;
            ParseContentElements(textRange, textboxElement.Elements(), 0);
        }

        /// <summary>
        /// Adds a list to the slide based on the XML element (supports multilevel).
        /// </summary>
        private static void AddList(PowerPoint.Slide slide, XElement listElement)
        {
            float left = float.TryParse(listElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(listElement.Attribute("top")?.Value, out var t) ? t : 100;
            float width = float.TryParse(listElement.Attribute("width")?.Value, out var w) ? w : 400;
            float height = float.TryParse(listElement.Attribute("height")?.Value, out var h) ? h : 200;

            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            var textRange = shape.TextFrame.TextRange;
            ParseList(textRange, listElement, 0);
        }

        /// <summary>
        /// Adds a table to the slide based on the XML element.
        /// </summary>
        private static void AddTable(PowerPoint.Slide slide, XElement tableElement)
        {
            float left = float.TryParse(tableElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(tableElement.Attribute("top")?.Value, out var t) ? t : 300;
            var rows = tableElement.Elements("row").ToList();
            int rowCount = rows.Count;
            int colCount = rows.Any() ? rows[0].Elements("cell").Count() : 0;

            if (rowCount > 0 && colCount > 0)
            {
                var tableShape = slide.Shapes.AddTable(rowCount, colCount, left, top, 500, 200);
                for (int r = 0; r < rowCount; r++)
                {
                    var cells = rows[r].Elements("cell").ToList();
                    for (int c = 0; c < colCount; c++)
                    {
                        if (c < cells.Count)
                        {
                            tableShape.Table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = cells[c].Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Adds a chart to the slide based on the XML element (basic support for bar, line, pie).
        /// </summary>
        private static void AddChart(PowerPoint.Slide slide, XElement chartElement)
        {
            string type = chartElement.Attribute("type")?.Value ?? "bar";
            float left = float.TryParse(chartElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(chartElement.Attribute("top")?.Value, out var t) ? t : 300;
            float width = float.TryParse(chartElement.Attribute("width")?.Value, out var w) ? w : 500;
            float height = float.TryParse(chartElement.Attribute("height")?.Value, out var h) ? h : 300;

            Office.XlChartType chartType;
            switch (type.ToLower())
            {
                case "line":
                    chartType = Office.XlChartType.xlLine;
                    break;
                case "pie":
                    chartType = Office.XlChartType.xlPie;
                    break;
                default:
                    chartType = Office.XlChartType.xlColumnClustered; // Bar/Column
                    break;
            }

            var chartShape = slide.Shapes.AddChart(chartType, left, top, width, height);
            var chart = chartShape.Chart;

            // Parse data from <data> element, assuming CSV-like rows
            var dataElement = chartElement.Element("data");
            if (dataElement != null)
            {
                var rows = dataElement.Value.Split(';'); // Semi-colon separated rows, comma-separated values
                var workbook = chart.ChartData.Workbook;
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                for (int r = 0; r < rows.Length; r++)
                {
                    var cols = rows[r].Split(',');
                    for (int c = 0; c < cols.Length; c++)
                    {
                        worksheet.Cells[r + 1, c + 1] = cols[c].Trim();
                    }
                }
                chart.Refresh();
                workbook.Close();
            }
        }

        /// <summary>
        /// Adds an image placeholder to the slide based on the XML element.
        /// </summary>
        private static void AddImage(PowerPoint.Slide slide, XElement imageElement)
        {
            string placeholder = imageElement.Value;
            float left = float.TryParse(imageElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(imageElement.Attribute("top")?.Value, out var t) ? t : 400;
            float width = float.TryParse(imageElement.Attribute("width")?.Value, out var w) ? w : 200;
            float height = float.TryParse(imageElement.Attribute("height")?.Value, out var h) ? h : 150;

            if (ImportHandlers.PlaceholderFiles.TryGetValue(placeholder, out string filePath) && File.Exists(filePath))
            {
                slide.Shapes.AddPicture(filePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, width, height);
            }
            else
            {
                // Add a placeholder rectangle if file not found
                var placeholderShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                placeholderShape.Fill.Visible = Office.MsoTriState.msoFalse;
                placeholderShape.Line.Visible = Office.MsoTriState.msoTrue;
                placeholderShape.TextFrame.TextRange.Text = $"Placeholder: {placeholder}";
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
                textRange.InsertAfter("\r");
            }
            else
            {
                textRange.InsertAfter(paragraphElement.Value + "\r");
            }
        }

        /// <summary>
        /// Parses a list element, handling nested lists and bullet/numbering (multilevel supported via recursion).
        /// </summary>
        private static void ParseList(PowerPoint.TextRange textRange, XElement listElement, int indentLevel)
        {
            string listType = listElement.Attribute("type")?.Value ?? "bullet";
            foreach (var itemElement in listElement.Elements("item"))
            {
                var itemTextRange = textRange.InsertAfter("\r");
                itemTextRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                itemTextRange.ParagraphFormat.Bullet.Type = listType == "bullet" ?
                    PowerPoint.PpBulletType.ppBulletUnnumbered : PowerPoint.PpBulletType.ppBulletNumbered;
                itemTextRange.IndentLevel = indentLevel + 1;
                ParseContentElements(itemTextRange, itemElement.Elements(), indentLevel + 1);
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