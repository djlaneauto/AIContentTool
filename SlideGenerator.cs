using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
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
        private struct FooterInfo
        {
            public string Text { get; set; }
            public float Size { get; set; }
            public int ColorRGB { get; set; }
        }

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

                // Prompt for template choice
                using (var dialog = new TemplateChoiceDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.UseCorporate)
                    {
                        string tempTemplatePath = ExtractEmbeddedTemplate();
                        if (!string.IsNullOrEmpty(tempTemplatePath))
                        {
                            try
                            {
                                presentation.ApplyTemplate(tempTemplatePath);
                                ApplyTemplateFooters(presentation, Globals.ThisAddIn.Application, tempTemplatePath);
                            }
                            catch (Exception ex)
                            {
                                ErrorHandler.LogError($"Failed to apply embedded template: {ex.Message}");
                                MessageBox.Show("Error applying template. Check the log for details.");
                            }
                            finally
                            {
                                try { File.Delete(tempTemplatePath); } catch { } // Cleanup, ignore errors if locked
                            }
                        }
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
                ErrorHandler.LogError($"Slide generation failed", ex); // Pass ex
                throw;
            }
        }

        private static string ExtractEmbeddedTemplate()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string resourceName = "AIContentTool.Resources.Template_Test.potx"; // Double-check this matches exactly
                ErrorHandler.LogInfo($"Attempting to extract resource: {resourceName}");
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        ErrorHandler.LogError("Embedded template resource not found");
                        return null;
                    }
                    string tempPath = Path.Combine(Path.GetTempPath(), "Template_Test.potx");
                    using (FileStream fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }
                    ErrorHandler.LogInfo($"Extracted template to {tempPath}");
                    return tempPath;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to extract embedded template", ex);
                return null;
            }
        }

        private class TemplateChoiceDialog : Form
        {
            public bool UseCorporate { get; private set; }

            public TemplateChoiceDialog()
            {
                this.Text = "Template Selection";
                this.Size = new System.Drawing.Size(300, 150);
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;

                Label label = new Label { Text = "Choose template:", Left = 20, Top = 20, Width = 260 };
                this.Controls.Add(label);

                Button btnCorporate = new Button { Text = "Use Corporate Template (recommended)", Left = 20, Top = 50, Width = 260 };
                btnCorporate.Click += (sender, e) => { UseCorporate = true; DialogResult = DialogResult.OK; Close(); };
                this.Controls.Add(btnCorporate);

                Button btnExisting = new Button { Text = "Use Existing Template", Left = 20, Top = 80, Width = 260 };
                btnExisting.Click += (sender, e) => { UseCorporate = false; DialogResult = DialogResult.OK; Close(); };
                this.Controls.Add(btnExisting);
            }
        }

        private static FooterInfo GetTemplateFooterInfo(string templatePath, PowerPoint.Application app)
        {
            try
            {
                var tempPres = app.Presentations.Open(templatePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                var master = tempPres.SlideMaster;
                FooterInfo info = new FooterInfo { Text = "", Size = 12, ColorRGB = 0 };

                foreach (PowerPoint.Shape shape in master.Shapes)
                {
                    if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter)
                    {
                        var textRange = shape.TextFrame.TextRange;
                        info.Text = textRange.Text ?? "";
                        info.Size = textRange.Font.Size;
                        info.ColorRGB = textRange.Font.Color.RGB;
                        break;
                    }
                }
                tempPres.Close();
                return info;
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"Failed to extract footer info from template: {ex.Message}");
                return new FooterInfo { Text = "", Size = 12, ColorRGB = 0 };
            }
        }

        private static void ApplyTemplateFooters(PowerPoint.Presentation presentation, PowerPoint.Application app, string templatePath)
        {
            FooterInfo info = GetTemplateFooterInfo(templatePath, app);

            var master = presentation.SlideMaster;
            var headersFooters = master.HeadersFooters;
            headersFooters.Footer.Visible = Office.MsoTriState.msoTrue;
            headersFooters.DateAndTime.Visible = Office.MsoTriState.msoTrue;
            headersFooters.SlideNumber.Visible = Office.MsoTriState.msoTrue;

            // Apply to footer shape
            foreach (PowerPoint.Shape shape in master.Shapes)
            {
                if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter)
                {
                    var textRange = shape.TextFrame.TextRange;
                    textRange.Text = info.Text;
                    textRange.Font.Size = info.Size;
                    textRange.Font.Color.RGB = info.ColorRGB;
                    break;
                }
            }

            // Apply to all layouts and additional masters
            foreach (PowerPoint.CustomLayout layout in master.CustomLayouts)
            {
                layout.HeadersFooters.Footer.Visible = Office.MsoTriState.msoTrue;
                foreach (PowerPoint.Shape layoutShape in layout.Shapes)
                {
                    if (layoutShape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter)
                    {
                        var layoutTextRange = layoutShape.TextFrame.TextRange;
                        layoutTextRange.Text = info.Text;
                        layoutTextRange.Font.Size = info.Size;
                        layoutTextRange.Font.Color.RGB = info.ColorRGB;
                        break;
                    }
                }
            }

            if (presentation.Designs.Count > 1)
            {
                foreach (PowerPoint.Design design in presentation.Designs)
                {
                    design.SlideMaster.HeadersFooters.Footer.Visible = Office.MsoTriState.msoTrue;
                    foreach (PowerPoint.Shape designShape in design.SlideMaster.Shapes)
                    {
                        if (designShape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter)
                        {
                            var designTextRange = designShape.TextFrame.TextRange;
                            designTextRange.Text = info.Text;
                            designTextRange.Font.Size = info.Size;
                            designTextRange.Font.Color.RGB = info.ColorRGB;
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Generates slides from XML content.
        /// </summary>
        private static void GenerateFromXml(PowerPoint.Presentation presentation, string xmlContent)
        {
            try
            {
                ErrorHandler.LogInfo("Attempting to parse XML: " + xmlContent.Substring(0, Math.Min(200, xmlContent.Length)) + "..."); // Log snippet
                var xmlDoc = XDocument.Parse(xmlContent);
                var slides = xmlDoc.Descendants("slide");
                foreach (var slideElement in slides)
                {
                    string title = slideElement.Element("title")?.Value ?? "Untitled";
                    var elements = slideElement.Elements().Where(e => e.Name != "title");
                    AddSlideWithElements(presentation, title, elements, slideElement);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("XML parsing failed", ex);
                ErrorHandler.LogInfo("Failed XML content: " + xmlContent); // Log full for review
                throw;
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
        private static void AddSlideWithElements(PowerPoint.Presentation presentation, string title, IEnumerable<XElement> elements, XElement slideElement = null)
        {
            // Lookup custom layout if specified
            PowerPoint.CustomLayout customLayout = null;
            string layoutName = slideElement?.Attribute("layout")?.Value;
            if (!string.IsNullOrEmpty(layoutName))
            {
                foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
                {
                    if (layout.Name == layoutName)
                    {
                        customLayout = layout;
                        break;
                    }
                }
            }
            if (customLayout == null)
            {
                customLayout = presentation.SlideMaster.CustomLayouts.Cast<PowerPoint.CustomLayout>().FirstOrDefault(l => l.Name == "Blank");
            }

            var slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, customLayout ?? presentation.SlideMaster.CustomLayouts[1]);
            if (customLayout == null)
            {
                slide.Layout = PowerPoint.PpSlideLayout.ppLayoutBlank; // Force blank
            }

            // Populate title placeholder
            if (!string.IsNullOrEmpty(title))
            {
                var titlePlaceholder = slide.Shapes.Placeholders.Cast<PowerPoint.Shape>().FirstOrDefault(s => s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle);
                if (titlePlaceholder != null)
                {
                    titlePlaceholder.TextFrame.TextRange.Text = title;
                    titlePlaceholder.TextFrame.TextRange.Font.Size = 32;
                    titlePlaceholder.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                }
                else
                {
                    var titleShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 20, 600, 50);
                    titleShape.TextFrame.TextRange.Text = title;
                    titleShape.TextFrame.TextRange.Font.Size = 32;
                    titleShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                }
            }

            // Elements mapping with enhanced fallback and logging
            var placeholderTypeMap = new Dictionary<string, PowerPoint.PpPlaceholderType>
    {
        { "body", PowerPoint.PpPlaceholderType.ppPlaceholderBody },
        { "subtitle", PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle },
        { "chart", PowerPoint.PpPlaceholderType.ppPlaceholderChart },
        { "table", PowerPoint.PpPlaceholderType.ppPlaceholderTable },
        { "picture", PowerPoint.PpPlaceholderType.ppPlaceholderPicture },
        { "media", PowerPoint.PpPlaceholderType.ppPlaceholderMediaClip },
        { "object", PowerPoint.PpPlaceholderType.ppPlaceholderObject } // Add more as needed
    };

            foreach (var element in elements)
            {
                string placeholderTypeStr = element.Attribute("placeholder")?.Value?.ToLower();
                if (!string.IsNullOrEmpty(placeholderTypeStr) && placeholderTypeMap.TryGetValue(placeholderTypeStr, out var ppType))
                {
                    var targetPlaceholder = slide.Shapes.Placeholders.Cast<PowerPoint.Shape>().FirstOrDefault(s => s.PlaceholderFormat.Type == ppType);
                    if (targetPlaceholder == null)
                    {
                        // Enhanced fallback to any content-like placeholder
                        targetPlaceholder = slide.Shapes.Placeholders.Cast<PowerPoint.Shape>().FirstOrDefault(s => s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderObject || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderVerticalBody);
                        if (targetPlaceholder != null)
                        {
                            ErrorHandler.LogInfo($"Fallback placeholder found for {placeholderTypeStr}: {targetPlaceholder.PlaceholderFormat.Type} at {targetPlaceholder.Left},{targetPlaceholder.Top}");
                        }
                        else
                        {
                            ErrorHandler.LogInfo("No suitable placeholder found for " + placeholderTypeStr + "; falling back to new shape");
                        }
                    }
                    else
                    {
                        ErrorHandler.LogInfo($"Found placeholder for {placeholderTypeStr}: {ppType} at {targetPlaceholder.Left},{targetPlaceholder.Top}");
                    }

                    if (targetPlaceholder != null)
                    {
                        ErrorHandler.LogInfo($"Found placeholder: {ppType} at {targetPlaceholder.Left},{targetPlaceholder.Top}");
                        if (element.Name.LocalName == "textbox" || element.Name.LocalName == "list")
                        {
                            var textRange = targetPlaceholder.TextFrame.TextRange;
                            textRange.Text = ""; // Clear default
                            ParseContentElements(textRange, element.Elements(), 0);
                        }
                        else if (element.Name.LocalName == "chart")
                        {
                            AddChart(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height);
                            targetPlaceholder.Delete();
                        }
                        else if (element.Name.LocalName == "table")
                        {
                            AddTable(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height);
                            targetPlaceholder.Delete();
                        }
                        else if (element.Name.LocalName == "image")
                        {
                            // For image, add textual description shape at placeholder position and delete the original placeholder
                            AddImage(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height);
                            targetPlaceholder.Delete();
                        }
                        continue;
                    }
                }

                // Fallback to adding new shape if no placeholder or mismatch
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
            // Original method: Parses positions from attributes
            float left = float.TryParse(tableElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(tableElement.Attribute("top")?.Value, out var t) ? t : 300;
            float width = float.TryParse(tableElement.Attribute("width")?.Value, out var w) ? w : 500;
            float height = float.TryParse(tableElement.Attribute("height")?.Value, out var h) ? h : 200;

            AddTable(slide, tableElement, left, top, width, height); // Call overload with parsed values
        }

        private static void AddTable(PowerPoint.Slide slide, XElement tableElement, float left, float top, float width, float height)
        {
            // Overload for explicit positions (e.g., from placeholder)
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
            // Original method: Parses positions from attributes
            float left = float.TryParse(chartElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(chartElement.Attribute("top")?.Value, out var t) ? t : 300;
            float width = float.TryParse(chartElement.Attribute("width")?.Value, out var w) ? w : 500;
            float height = float.TryParse(chartElement.Attribute("height")?.Value, out var h) ? h : 300;

            AddChart(slide, chartElement, left, top, width, height); // Call overload with parsed values
        }

        private static void AddChart(PowerPoint.Slide slide, XElement chartElement, float left, float top, float width, float height)
        {
            // Overload for explicit positions (e.g., from placeholder)
            string type = chartElement.Attribute("type")?.Value ?? "bar";

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
                            var series = (PowerPoint.Series)chart.SeriesCollection(seriesIndex); // Use method call with parentheses
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
            // Original method: Parses positions from attributes
            float left = float.TryParse(imageElement.Attribute("left")?.Value, out var l) ? l : 100;
            float top = float.TryParse(imageElement.Attribute("top")?.Value, out var t) ? t : 400;
            float width = float.TryParse(imageElement.Attribute("width")?.Value, out var w) ? w : 200;
            float height = float.TryParse(imageElement.Attribute("height")?.Value, out var h) ? h : 150;

            AddImage(slide, imageElement, left, top, width, height); // Call overload with parsed values
        }

        private static void AddImage(PowerPoint.Slide slide, XElement imageElement, float left, float top, float width, float height)
        {
            ErrorHandler.LogInfo($"Adding image with dimensions: left={left}, top={top}, width={width}, height={height}");
            // Overload for explicit positions (e.g., from placeholder)
            // Log dimensions for debugging
            ErrorHandler.LogInfo($"Image dimensions before validation: left={left}, top={top}, width={width}, height={height}");

            // Get slide dimensions for max clamping
            float slideWidth = slide.Parent.PageSetup.SlideWidth;
            float slideHeight = slide.Parent.PageSetup.SlideHeight;

            // Validate and clamp/fallback dimensions to valid ranges (min 1, positive, max slide size - position; default if 0/invalid)
            left = float.IsNaN(left) || left < 0 ? 100 : Math.Min(left, slideWidth - 1);
            top = float.IsNaN(top) || top < 0 ? 400 : Math.Min(top, slideHeight - 1);
            width = float.IsNaN(width) || width <= 0 ? 200 : Math.Max(1, Math.Min(width, slideWidth - left));
            height = float.IsNaN(height) || height <= 0 ? 150 : Math.Max(1, Math.Min(height, slideHeight - top));

            // Log after validation
            ErrorHandler.LogInfo($"Image dimensions after validation: left={left}, top={top}, width={width}, height={height}");

            string placeholder = imageElement.Value;

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
                        ApplyTextFormatting(plainRun, pStyle, pFontSize, pColor); // Explicit color handling
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
            textRun.Font.Italic = style.Contains("italic") ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            textRun.Font.Underline = style.Contains("underline") ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            if (int.TryParse(fontSize, out int size))
            {
                textRun.Font.Size = size;
            }
            // Explicit color set/reset to prevent inheritance
            textRun.Font.Color.RGB = !string.IsNullOrEmpty(color) ? ParseColorToOle(color) : 0; // Black if no color
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