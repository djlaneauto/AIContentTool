using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
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
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
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
                        break;
                }
            }
        }

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

        private static void AddTable(PowerPoint.Slide slide, XElement tableElement)
        {
            float left = float.TryParse(tableElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(tableElement.Attribute("top")?.Value, out var t) ? t : 300;
            float width = float.TryParse(tableElement.Attribute("width")?.Value, out var w) ? w : 500;
            float height = float.TryParse(tableElement.Attribute("height")?.Value, out var h) ? h : 200;
            var rows = tableElement.Elements("row").ToList();
            int rowCount = rows.Count;
            int colCount = rows.Any() ? rows[0].Elements("cell").Count() : 0;

            if (rowCount > 0 && colCount > 0)
            {
                var tableShape = slide.Shapes.AddTable(rowCount, colCount, left, top, width, height);
                for (int r = 0; r < rowCount; r++)
                {
                    var cells = rows[r].Elements("cell").ToList();
                    for (int c = 0; c < colCount; c++)
                    {
                        if (c < cells.Count)
                        {
                            var cell = tableShape.Table.Cell(r + 1, c + 1);
                            cell.Shape.TextFrame.TextRange.Text = cells[c].Value;
                            string style = cells[c].Attribute("style")?.Value ?? "";
                            string color = cells[c].Attribute("color")?.Value ?? "";
                            string fontSize = cells[c].Attribute("font-size")?.Value ?? "";
                            ApplyTextFormatting(cell.Shape.TextFrame.TextRange, style, fontSize, color);
                        }
                    }
                }
            }
        }

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
                    chartType = Office.XlChartType.xlColumnClustered;
                    break;
            }

            var chartShape = slide.Shapes.AddChart(chartType, left, top, width, height);
            var chart = chartShape.Chart;

            var dataElement = chartElement.Element("data");
            if (dataElement != null)
            {
                var rows = dataElement.Value.Split(';');
                int rowCount = rows.Length;
                int colCount = rowCount > 0 ? rows[0].Split(',').Length : 0;

                var workbook = chart.ChartData.Workbook;
                var worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                worksheet.UsedRange.ClearContents(); // Clear default data

                for (int r = 0; r < rowCount; r++)
                {
                    var cols = rows[r].Split(',');
                    for (int c = 0; c < cols.Length; c++)
                    {
                        worksheet.Cells[r + 1, c + 1] = cols[c].Trim();
                    }
                }

            var seriesElements = chartElement.Elements("series");
            if (seriesElements.Any())
            {
                int seriesIndex = 1;
                foreach (var seriesElem in seriesElements)
                {
                    string color = seriesElem.Attribute("color")?.Value ?? "";
                    if (!string.IsNullOrEmpty(color))
                    {
                        var series = (PowerPoint.Series)chart.SeriesCollection(seriesIndex);
                        series.Format.Fill.ForeColor.RGB = ParseColorToOle(color);
                    }
                    seriesIndex++;
                }
            }

            // Set source data range
            string lastCol = ((char)('A' + colCount - 1)).ToString();
            string rangeAddress = $"A1:{lastCol}{rowCount}";
            chart.SetSourceData($"Sheet1!{rangeAddress}", Excel.XlRowCol.xlRows);
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
                var placeholderShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                placeholderShape.Fill.Visible = Office.MsoTriState.msoFalse;
                placeholderShape.Line.Visible = Office.MsoTriState.msoTrue;
                placeholderShape.TextFrame.TextRange.Text = $"Placeholder: {placeholder}";
                string color = imageElement.Attribute("color")?.Value ?? "red"; // Default red, override via attr
                placeholderShape.TextFrame.TextRange.Font.Color.RGB = ParseColorToOle(color);
                placeholderShape.TextFrame.TextRange.Font.Size = 12;
            }
        } /// <summary>
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

        private static void ParseParagraph(PowerPoint.TextRange textRange, XElement paragraphElement)
        {
            bool hasContent = false;
            string pStyle = paragraphElement.Attribute("style")?.Value ?? "";
            string pColor = paragraphElement.Attribute("color")?.Value ?? "";
            string pFontSize = paragraphElement.Attribute("font-size")?.Value ?? "";
            foreach (var node in paragraphElement.Nodes())
            {
                if (node is XText textNode)
                {
                    string plainText = textNode.Value;
                    if (!string.IsNullOrEmpty(plainText))
                    {
                        var plainRun = textRange.InsertAfter(plainText);
                        ApplyTextFormatting(plainRun, pStyle, pFontSize, pColor); // Apply para-level if no inline
                        hasContent = true;
                    }
                }
                else if (node is XElement textElement && textElement.Name == "text")
                {
                    string text = textElement.Value;
                    if (!string.IsNullOrEmpty(text))
                    {
                        var textRun = textRange.InsertAfter(text);
                        string style = textElement.Attribute("style")?.Value ?? pStyle;
                        string fontSize = textElement.Attribute("font-size")?.Value ?? pFontSize;
                        string color = textElement.Attribute("color")?.Value ?? pColor;
                        ApplyTextFormatting(textRun, style, fontSize, color);
                        hasContent = true;
                    }
                }
            }
            if (hasContent)
            {
                textRange.InsertAfter("\r");
            }
        }
        private static void ParseList(PowerPoint.TextRange textRange, XElement listElement, int indentLevel)
        {
            string listType = listElement.Attribute("type")?.Value ?? "bullet";
            foreach (var itemElement in listElement.Elements("item"))
            {
                // Add item text if present (from value or <p>)
                if (itemElement.HasElements)
                {
                    foreach (var p in itemElement.Elements("p"))
                    {
                        ParseParagraph(textRange, p); // Adds text and \r; ParseParagraph handles formatting
                                                      // Note: If multiple <p>, it adds multiple paras per item; if you want single, we can adjust
                    }
                }
                else if (!string.IsNullOrEmpty(itemElement.Value))
                {
                    var itemRun = textRange.InsertAfter(itemElement.Value);
                    // Apply formatting for plain value (from item attributes)
                    string style = itemElement.Attribute("style")?.Value ?? "";
                    string color = itemElement.Attribute("color")?.Value ?? "";
                    string fontSize = itemElement.Attribute("font-size")?.Value ?? "";
                    ApplyTextFormatting(itemRun, style, fontSize, color);
                }

                // Always add paragraph break for this item
                textRange.InsertAfter("\r");

                // Set formatting on the new paragraph
                var itemPara = textRange.Paragraphs(textRange.Paragraphs().Count, 1);
                itemPara.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                itemPara.ParagraphFormat.Bullet.Type = listType == "bullet" ?
                    PowerPoint.PpBulletType.ppBulletUnnumbered : PowerPoint.PpBulletType.ppBulletNumbered;
                itemPara.IndentLevel = indentLevel + 1;

                // Handle nested content (e.g., sub-lists)
                if (itemElement.HasElements)
                {
                    ParseContentElements(textRange, itemElement.Elements("list"), indentLevel + 1);
                }
            }
        }
        private static void ApplyTextFormatting(PowerPoint.TextRange textRun, string style, string fontSize, string color = "")
        {
            textRun.Font.Bold = style.Contains("bold") ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            textRun.Font.Italic = style.Contains("italic") ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; // New: Italics
            textRun.Font.Underline = style.Contains("underline") ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            if (int.TryParse(fontSize, out int size))
            {
                textRun.Font.Size = size;
            }
            if (!string.IsNullOrEmpty(color))
            {
                textRun.Font.Color.RGB = ParseColorToOle(color);
            }
        }
        private static int ParseColorToOle(string colorStr)
        {
            if (string.IsNullOrEmpty(colorStr)) return 0; // Black default
            try
            {
                Color color;
                if (colorStr.StartsWith("#"))
                {
                    color = ColorTranslator.FromHtml(colorStr);
                }
                else
                {
                    color = Color.FromName(colorStr);
                }
                return ColorTranslator.ToOle(color);
            }
            catch
            {
                return 0; // Fallback black
            }
        }
    }
}