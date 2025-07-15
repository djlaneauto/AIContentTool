using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DCharts = DocumentFormat.OpenXml.Drawing.Charts;
using DGraphic = DocumentFormat.OpenXml.Drawing;
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

        public static int GenerateSlides(PowerPoint.Presentation presentation, string content, string format)
        {
            try
            {
                DateTime startTime = DateTime.Now;

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
                                try { File.Delete(tempTemplatePath); } catch { }
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
                ErrorHandler.LogError($"Slide generation failed", ex);
                throw;
            }
        }

        private static string ExtractEmbeddedTemplate()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string resourceName = "AIContentTool.Resources.Template_Test.potx";
                ErrorHandler.LogInfo($"Attempting to extract resource: {resourceName}");
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        ErrorHandler.LogError("Embedded template resource not found");
                        return null;
                    }
                    string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Template_Test.potx");
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

        private static void GenerateFromXml(PowerPoint.Presentation presentation, string xmlContent)
        {
            try
            {
                ErrorHandler.LogInfo("Attempting to parse XML: " + xmlContent.Substring(0, Math.Min(200, xmlContent.Length)) + "...");

                string schemaString = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                        <xs:element name='presentation'>
                            <xs:complexType>
                                <xs:sequence>
                                    <xs:element name='slide' minOccurs='0' maxOccurs='unbounded'>
                                        <xs:complexType>
                                            <xs:sequence minOccurs='0' maxOccurs='unbounded'>
                                                <xs:element name='title' minOccurs='0' maxOccurs='1' type='xs:string'/>
                                                <xs:element name='textbox' minOccurs='0' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence minOccurs='0' maxOccurs='unbounded'>
                                                            <xs:element name='p' minOccurs='0' maxOccurs='unbounded'/>
                                                            <xs:element name='list' minOccurs='0' maxOccurs='unbounded'/>
                                                        </xs:sequence>
                                                        <xs:attribute name='left' type='xs:float'/>
                                                        <xs:attribute name='top' type='xs:float'/>
                                                        <xs:attribute name='width' type='xs:float'/>
                                                        <xs:attribute name='height' type='xs:float'/>
                                                        <xs:attribute name='placeholder' type='xs:string'/>
                                                    </xs:complexType>
                                                </xs:element>
                                                <xs:element name='list' minOccurs='0' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence minOccurs='0' maxOccurs='unbounded'>
                                                            <xs:element name='item' minOccurs='0' maxOccurs='unbounded'/>
                                                        </xs:sequence>
                                                        <xs:attribute name='type' type='xs:string' use='required'/>
                                                        <xs:attribute name='left' type='xs:float'/>
                                                        <xs:attribute name='top' type='xs:float'/>
                                                        <xs:attribute name='width' type='xs:float'/>
                                                        <xs:attribute name='height' type='xs:float'/>
                                                        <xs:attribute name='placeholder' type='xs:string'/>
                                                    </xs:complexType>
                                                </xs:element>
                                                <xs:element name='table' minOccurs='0' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence minOccurs='0' maxOccurs='unbounded'>
                                                            <xs:element name='row' minOccurs='0' maxOccurs='unbounded'/>
                                                        </xs:sequence>
                                                        <xs:attribute name='left' type='xs:float'/>
                                                        <xs:attribute name='top' type='xs:float'/>
                                                        <xs:attribute name='width' type='xs:float'/>
                                                        <xs:attribute name='height' type='xs:float'/>
                                                        <xs:attribute name='placeholder' type='xs:string'/>
                                                    </xs:complexType>
                                                </xs:element>
                                                <xs:element name='chart' minOccurs='0' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence minOccurs='0' maxOccurs='unbounded'>
                                                            <xs:element name='data' minOccurs='0' maxOccurs='1' type='xs:string'/>
                                                            <xs:element name='series' minOccurs='0' maxOccurs='unbounded'/>
                                                        </xs:sequence>
                                                        <xs:attribute name='type' type='xs:string' use='required'/>
                                                        <xs:attribute name='left' type='xs:float'/>
                                                        <xs:attribute name='top' type='xs:float'/>
                                                        <xs:attribute name='width' type='xs:float'/>
                                                        <xs:attribute name='height' type='xs:float'/>
                                                        <xs:attribute name='placeholder' type='xs:string'/>
                                                    </xs:complexType>
                                                </xs:element>
                                                <xs:element name='image' minOccurs='0' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:simpleContent>
                                                            <xs:extension base='xs:string'>
                                                                <xs:attribute name='left' type='xs:float'/>
                                                                <xs:attribute name='top' type='xs:float'/>
                                                                <xs:attribute name='width' type='xs:float'/>
                                                                <xs:attribute name='height' type='xs:float'/>
                                                                <xs:attribute name='color' type='xs:string'/>
                                                                <xs:attribute name='placeholder' type='xs:string'/>
                                                            </xs:extension>
                                                        </xs:simpleContent>
                                                    </xs:complexType>
                                                </xs:element>
                                            </xs:sequence>
                                            <xs:attribute name='layout' type='xs:string'/>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:sequence>
                            </xs:complexType>
                        </xs:element>
                    </xs:schema>";

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.Schemas.Add(null, XmlReader.Create(new StringReader(schemaString)));
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationEventHandler += (sender, e) => throw new XmlSchemaValidationException(e.Message);

                using (var stringReader = new StringReader(xmlContent))
                using (var xmlReader = XmlReader.Create(stringReader, settings))
                {
                    var xmlDoc = XDocument.Load(xmlReader);
                    var slides = xmlDoc.Descendants("slide");
                    foreach (var slideElement in slides)
                    {
                        string title = slideElement.Element("title")?.Value ?? "Untitled";
                        var elements = slideElement.Elements().Where(e => e.Name != "title");
                        AddSlideWithElements(presentation, title, elements, slideElement);
                    }
                }
            }
            catch (XmlSchemaValidationException valEx)
            {
                ErrorHandler.LogError("XML validation failed", valEx);
                throw new ArgumentException("Invalid XML structure per schema.", valEx);
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("XML parsing failed", ex);
                ErrorHandler.LogInfo("Failed XML content: " + xmlContent);
                throw;
            }
        }

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
                            }
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

        private static void AddSlideWithElements(PowerPoint.Presentation presentation, string title, IEnumerable<XElement> elements, XElement slideElement = null)
        {
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
                slide.Layout = PowerPoint.PpSlideLayout.ppLayoutBlank;
            }

            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;

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

            var placeholderTypeMap = new Dictionary<string, PowerPoint.PpPlaceholderType>
            {
                { "body", PowerPoint.PpPlaceholderType.ppPlaceholderBody },
                { "subtitle", PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle },
                { "chart", PowerPoint.PpPlaceholderType.ppPlaceholderChart },
                { "table", PowerPoint.PpPlaceholderType.ppPlaceholderTable },
                { "picture", PowerPoint.PpPlaceholderType.ppPlaceholderPicture },
                { "media", PowerPoint.PpPlaceholderType.ppPlaceholderMediaClip },
                { "object", PowerPoint.PpPlaceholderType.ppPlaceholderObject }
            };

            foreach (var element in elements)
            {
                string placeholderTypeStr = element.Attribute("placeholder")?.Value?.ToLower();
                if (!string.IsNullOrEmpty(placeholderTypeStr) && placeholderTypeMap.TryGetValue(placeholderTypeStr, out var ppType))
                {
                    var targetPlaceholder = slide.Shapes.Placeholders.Cast<PowerPoint.Shape>().FirstOrDefault(s => s.PlaceholderFormat.Type == ppType);
                    if (targetPlaceholder == null)
                    {
                        targetPlaceholder = slide.Shapes.Placeholders.Cast<PowerPoint.Shape>().FirstOrDefault(s => s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderObject || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderVerticalBody || s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderMixed);
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
                        if (element.Name.LocalName == "textbox")
                        {
                            var textRange = targetPlaceholder.TextFrame.TextRange;
                            textRange.Text = "";
                            ParseContentElements(textRange, element.Elements(), 0);
                        }
                        else if (element.Name.LocalName == "list")
                        {
                            var textRange = targetPlaceholder.TextFrame.TextRange;
                            textRange.Text = "";
                            ParseList(textRange, element, 0);
                        }
                        else if (element.Name.LocalName == "chart")
                        {
                            AddChart(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height, slideWidth, slideHeight);
                        }
                        else if (element.Name.LocalName == "table")
                        {
                            AddTable(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height, slideWidth, slideHeight);
                        }
                        else if (element.Name.LocalName == "image")
                        {
                            AddImage(slide, element, targetPlaceholder.Left, targetPlaceholder.Top, targetPlaceholder.Width, targetPlaceholder.Height, slideWidth, slideHeight);
                        }
                        continue;
                    }
                }

                switch (element.Name.LocalName)
                {
                    case "textbox":
                        AddTextbox(slide, element, slideWidth, slideHeight);
                        break;
                    case "list":
                        AddList(slide, element, slideWidth, slideHeight);
                        break;
                    case "table":
                        AddTable(slide, element, slideWidth, slideHeight);
                        break;
                    case "chart":
                        AddChart(slide, element, slideWidth, slideHeight);
                        break;
                    case "image":
                        AddImage(slide, element, slideWidth, slideHeight);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void AddTextbox(PowerPoint.Slide slide, XElement textboxElement, float slideWidth, float slideHeight)
        {
            float left = ParseAndValidatePosition(textboxElement.Attribute("left")?.Value, 100, 0, slideWidth - 1);
            float top = ParseAndValidatePosition(textboxElement.Attribute("top")?.Value, 100, 0, slideHeight - 1);
            float width = ParseAndValidatePosition(textboxElement.Attribute("width")?.Value, 400, 1, slideWidth - left);
            float height = ParseAndValidatePosition(textboxElement.Attribute("height")?.Value, 200, 1, slideHeight - top);

            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            var textRange = shape.TextFrame.TextRange;
            ParseContentElements(textRange, textboxElement.Elements(), 0);
        }

        private static void AddList(PowerPoint.Slide slide, XElement listElement, float slideWidth, float slideHeight)
        {
            float left = ParseAndValidatePosition(listElement.Attribute("left")?.Value, 100, 0, slideWidth - 1);
            float top = ParseAndValidatePosition(listElement.Attribute("top")?.Value, 100, 0, slideHeight - 1);
            float width = ParseAndValidatePosition(listElement.Attribute("width")?.Value, 400, 1, slideWidth - left);
            float height = ParseAndValidatePosition(listElement.Attribute("height")?.Value, 200, 1, slideHeight - top);

            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            var textRange = shape.TextFrame.TextRange;
            ParseList(textRange, listElement, 0);
        }

        private static void AddTable(PowerPoint.Slide slide, XElement tableElement, float slideWidth, float slideHeight)
        {
            float left = ParseAndValidatePosition(tableElement.Attribute("left")?.Value, 100, 0, slideWidth - 1);
            float top = ParseAndValidatePosition(tableElement.Attribute("top")?.Value, 300, 0, slideHeight - 1);
            float width = ParseAndValidatePosition(tableElement.Attribute("width")?.Value, 500, 1, slideWidth - left);
            float height = ParseAndValidatePosition(tableElement.Attribute("height")?.Value, 200, 1, slideHeight - top);

            AddTable(slide, tableElement, left, top, width, height, slideWidth, slideHeight);
        }

        private static void AddTable(PowerPoint.Slide slide, XElement tableElement, float left, float top, float width, float height, float slideWidth = 0, float slideHeight = 0)
        {
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

        private static void AddChart(PowerPoint.Slide slide, XElement chartElement, float slideWidth, float slideHeight)
        {
            float left = ParseAndValidatePosition(chartElement.Attribute("left")?.Value, 100, 0, slideWidth - 1);
            float top = ParseAndValidatePosition(chartElement.Attribute("top")?.Value, 300, 0, slideHeight - 1);
            float width = ParseAndValidatePosition(chartElement.Attribute("width")?.Value, 500, 1, slideWidth - left);
            float height = ParseAndValidatePosition(chartElement.Attribute("height")?.Value, 300, 1, slideHeight - top);

            AddChart(slide, chartElement, left, top, width, height, slideWidth, slideHeight);
        }

        private static void AddChart(PowerPoint.Slide slide, XElement chartElement, float left, float top, float width, float height, float slideWidth = 0, float slideHeight = 0)
        {
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

            try
            {
                var chartShape = slide.Shapes.AddChart(chartType, left, top, width, height);
                var chart = chartShape.Chart;

                var dataElement = chartElement.Element("data");
                if (dataElement != null)
                {
                    var rows = dataElement.Value.Split(';');
                    int rowCount = rows.Length;
                    int colCount = rowCount > 0 ? rows[0].Split(',').Length : 0;

                    if (rowCount > 100)
                    {
                        ErrorHandler.LogInfo($"Large chart data detected ({rowCount} rows); may cause timeouts—consider fallback.");
                    }

                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        workbook = chart.ChartData.Workbook;
                        workbook.Application.Visible = false;
                        worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                        worksheet.UsedRange.ClearContents();

                        if (chartType == Office.XlChartType.xlPie)
                        {
                            worksheet.Cells[1, 1] = "Categories";
                            worksheet.Cells[1, 2] = "Values";
                            for (int r = 0; r < rowCount; r++)
                            {
                                var cols = rows[r].Split(',');
                                if (cols.Length >= 2)
                                {
                                    worksheet.Cells[r + 2, 1] = cols[0].Trim();
                                    worksheet.Cells[r + 2, 2] = cols[1].Trim();
                                }
                            }
                            rowCount += 1;
                            colCount = 2;
                        }
                        else
                        {
                            for (int r = 0; r < rowCount; r++)
                            {
                                var cols = rows[r].Split(',');
                                for (int c = 0; c < cols.Length; c++)
                                {
                                    worksheet.Cells[r + 1, c + 1] = cols[c].Trim();
                                }
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

                        string lastCol = ((char)('A' + colCount - 1)).ToString();
                        string rangeAddress = $"A1:{lastCol}{rowCount}";
                        Excel.XlRowCol rowCol = (chartType == Office.XlChartType.xlPie) ? Excel.XlRowCol.xlColumns : Excel.XlRowCol.xlRows;
                        chart.SetSourceData($"Sheet1!{rangeAddress}", rowCol);
                        chart.Refresh();
                    }
                    finally
                    {
                        if (worksheet != null) Marshal.FinalReleaseComObject(worksheet);
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.FinalReleaseComObject(workbook);
                        }
                        if (chart.ChartData != null) Marshal.FinalReleaseComObject(chart.ChartData);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"Interop chart creation failed: {ex.Message}; falling back to OpenXML.");
                AddChartWithOpenXml(slide.Parent, slide, chartElement, left, top, width, height, chartType);
            }
        }

        private static void AddChartWithOpenXml(PowerPoint.Presentation presentation, PowerPoint.Slide slide, XElement chartElement, float left, float top, float width, float height, Office.XlChartType chartType)
        {
            try
            {
                string tempPptx = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "temp.pptx");
                presentation.SaveAs(tempPptx);

                using (PresentationDocument doc = PresentationDocument.Open(tempPptx, true))
                {
                    SlidePart slidePart = doc.PresentationPart.SlideParts.LastOrDefault();
                    if (slidePart == null) return;

                    ChartPart chartPart = slidePart.AddNewPart<ChartPart>();
                    chartPart.ChartSpace = new DCharts.ChartSpace();

                    DCharts.Chart chart = new DCharts.Chart();
                    switch (chartType)
                    {
                        case Office.XlChartType.xlLine:
                            chart.Append(new DCharts.LineChart());
                            break;
                        case Office.XlChartType.xlPie:
                            chart.Append(new DCharts.PieChart());
                            break;
                        default:
                            chart.Append(new DCharts.BarChart());
                            break;
                    }

                    var dataElement = chartElement.Element("data");
                    if (dataElement != null)
                    {
                        var rows = dataElement.Value.Split(';');
                        for (int seriesIndex = 0; seriesIndex < rows.Length; seriesIndex++)
                        {
                            var series = new DCharts.BarChartSeries();
                            var values = rows[seriesIndex].Split(',');
                            for (int valIndex = 0; valIndex < values.Length; valIndex++)
                            {
                                series.Append(new DCharts.NumericValue(values[valIndex].Trim()));
                            }
                            chart.FirstOrDefault()?.Append(series);
                        }
                    }

                    chartPart.ChartSpace.Append(chart);
                    chartPart.ChartSpace.Save();

                    GraphicFrame gf = new GraphicFrame();
                    gf.Graphic = new DGraphic.Graphic();
                    gf.Graphic.GraphicData = new DGraphic.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
                    gf.Graphic.GraphicData.Append(new DCharts.ChartReference() { Id = slidePart.GetIdOfPart(chartPart) });

                    slidePart.Slide.CommonSlideData.ShapeTree.Append(gf);
                    slidePart.Slide.Save();
                }

                presentation.Close();
                presentation = Globals.ThisAddIn.Application.Presentations.Open(tempPptx);
                System.IO.File.Delete(tempPptx);
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError($"OpenXML chart fallback failed: {ex.Message}");
                var placeholderShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                placeholderShape.TextFrame.TextRange.Text = "Chart Placeholder (Generation Failed)";
                placeholderShape.Fill.Visible = Office.MsoTriState.msoFalse;
                placeholderShape.Line.Visible = Office.MsoTriState.msoTrue;
            }
        }

        private static void AddImage(PowerPoint.Slide slide, XElement imageElement, float slideWidth, float slideHeight)
        {
            float left = ParseAndValidatePosition(imageElement.Attribute("left")?.Value, 100, 0, slideWidth - 1);
            float top = ParseAndValidatePosition(imageElement.Attribute("top")?.Value, 400, 0, slideHeight - 1);
            float width = ParseAndValidatePosition(imageElement.Attribute("width")?.Value, 200, 1, slideWidth - left);
            float height = ParseAndValidatePosition(imageElement.Attribute("height")?.Value, 150, 1, slideHeight - top);

            AddImage(slide, imageElement, left, top, width, height, slideWidth, slideHeight);
        }

        private static void AddImage(PowerPoint.Slide slide, XElement imageElement, float left, float top, float width, float height, float slideWidth = 0, float slideHeight = 0)
        {
            ErrorHandler.LogInfo($"Adding image with dimensions: left={left}, top={top}, width={width}, height={height}");
            string description = imageElement.Value;

            PowerPoint.Shape imageShape = null;
            if (ImportHandlers.PlaceholderFiles.TryGetValue(description, out string filePath) && File.Exists(filePath))
            {
                try
                {
                    imageShape = slide.Shapes.AddPicture(filePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, width, height);
                    imageShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                    if (imageShape.Width > width || imageShape.Height > height)
                    {
                        imageShape.ScaleWidth(1, Office.MsoTriState.msoTrue);
                        imageShape.ScaleHeight(1, Office.MsoTriState.msoTrue);
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.LogError($"Failed to add image from {filePath}: {ex.Message}");
                    imageShape = null;
                }
            }

            if (imageShape == null)
            {
                imageShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                imageShape.Fill.Visible = Office.MsoTriState.msoFalse;
                imageShape.Line.Visible = Office.MsoTriState.msoTrue;
                string color = imageElement.Attribute("color")?.Value ?? "red";
                imageShape.Line.ForeColor.RGB = ParseColorToOle(color);
            }

            if (!string.IsNullOrEmpty(description))
            {
                float textLeft = left + (width / 4);
                float textTop = top + (height / 2) - 20;
                float textWidth = width / 2;
                float textHeight = 40;
                var textShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, textLeft, textTop, textWidth, textHeight);
                textShape.TextFrame.TextRange.Text = $"Add {description}";
                textShape.TextFrame.TextRange.Font.Size = 12;
                textShape.TextFrame.TextRange.Font.Color.RGB = ParseColorToOle("black");
                textShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                textShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                textShape.Fill.Visible = Office.MsoTriState.msoFalse;
                textShape.Line.Visible = Office.MsoTriState.msoFalse;
            }

            var captionElement = imageElement.Element("caption");
            if (captionElement != null)
            {
                try
                {
                    string captionText = captionElement.Value;
                    if (!string.IsNullOrEmpty(captionText))
                    {
                        float captionLeft = 50;
                        float captionTop = top;
                        float captionWidth = left - 60;
                        float captionHeight = height;
                        var captionShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, captionLeft, captionTop, captionWidth, captionHeight);
                        captionShape.TextFrame.TextRange.Text = captionText;
                        captionShape.TextFrame.TextRange.Font.Size = 14;
                        captionShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        captionShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                        captionShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        captionShape.Fill.Visible = Office.MsoTriState.msoFalse;
                        captionShape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.LogError($"Failed to add caption: {ex.Message}");
                }
            }
        }

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
                        ApplyTextFormatting(plainRun, pStyle, pFontSize, pColor);
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
                if (itemElement.HasElements)
                {
                    foreach (var p in itemElement.Elements("p"))
                    {
                        ParseParagraph(textRange, p);
                    }
                }
                else if (!string.IsNullOrEmpty(itemElement.Value))
                {
                    var itemRun = textRange.InsertAfter(itemElement.Value);
                    string style = itemElement.Attribute("style")?.Value ?? "";
                    string color = itemElement.Attribute("color")?.Value ?? "";
                    string fontSize = itemElement.Attribute("font-size")?.Value ?? "";
                    ApplyTextFormatting(itemRun, style, fontSize, color);
                }

                textRange.InsertAfter("\r");

                var itemPara = textRange.Paragraphs(textRange.Paragraphs().Count, 1);
                itemPara.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                itemPara.ParagraphFormat.Bullet.Type = listType == "bullet" ?
                    PowerPoint.PpBulletType.ppBulletUnnumbered : PowerPoint.PpBulletType.ppBulletNumbered;
                itemPara.IndentLevel = indentLevel + 1;

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
            textRun.Font.Color.RGB = !string.IsNullOrEmpty(color) ? ParseColorToOle(color) : 0;
        }

        private static int ParseColorToOle(string colorStr)
        {
            if (string.IsNullOrEmpty(colorStr)) return 0;
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
                return 0;
            }
        }

        private static float ParseAndValidatePosition(string value, float defaultValue, float minValue, float maxValue)
        {
            if (!float.TryParse(value, out float parsed) || float.IsNaN(parsed) || float.IsInfinity(parsed))
            {
                ErrorHandler.LogInfo($"Invalid position '{value}'; using default {defaultValue}");
                parsed = defaultValue;
            }
            float clamped = Math.Max(minValue, Math.Min(parsed, maxValue));
            if (clamped != parsed)
            {
                ErrorHandler.LogInfo($"Clamped position from {parsed} to {clamped} (bounds: {minValue}-{maxValue})");
            }
            return clamped;
        }
    }
}