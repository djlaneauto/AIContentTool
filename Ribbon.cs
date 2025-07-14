using Microsoft.Office.Tools.Ribbon;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace AIContentTool
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OnImportXmlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ImportHandlers.ImportXml();
                MessageBox.Show("XML imported successfully!");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to import XML: " + ex.Message);
                MessageBox.Show("Error importing XML. Check the log for details.");
            }
        }

        // Import JSON button click handler
        private void OnImportJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ImportHandlers.ImportJson();
                MessageBox.Show("JSON imported successfully!");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to import JSON: " + ex.Message);
                MessageBox.Show("Error importing JSON. Check the log for details.");
            }
        }

        // Import Markdown button click handler
        private void OnImportMarkdownClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ImportHandlers.ImportMarkdown();
                MessageBox.Show("Markdown imported successfully!");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to import Markdown: " + ex.Message);
                MessageBox.Show("Error importing Markdown. Check the log for details.");
            }
        }

        // Import from Clipboard button click handler
        private void OnCopyFromClipboardClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ImportHandlers.ImportClipboard();
                MessageBox.Show("Clipboard content imported successfully!");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to import from clipboard: " + ex.Message);
                MessageBox.Show("Error importing from clipboard. Check the log for details.");
            }
        }

        // Generate Slides button click handler
        private void OnGenerateSlidesClick(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.PpApp == null)
            {
                MessageBox.Show("PowerPoint is still initializing. Please try again in a moment.");
                return;
            }

            try
            {
                // UNcomment this block if you want to prompt for a template selection dialog
                //// Show template selection dialog
                //using (var dialog = new TemplateSelectionDialog())
                //{
                //    if (dialog.ShowDialog() != DialogResult.OK)
                //    {
                //        // User canceled, abort slide generation
                //        return;
                //    }
                //}

                // Get the current presentation
                PowerPoint.Presentation presentation = ThisAddIn.CurrentPresentation;
                if (presentation == null)
                {
                    // No active presentation, create a new one
                    presentation = ThisAddIn.PpApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                    ThisAddIn.CurrentPresentation = presentation;
                }

                // Get imported content and format
                string content = ImportHandlers.ImportedContent;
                string format = ImportHandlers.ContentFormat;
                if (string.IsNullOrEmpty(content))
                {
                    MessageBox.Show("No content imported.");
                    return;
                }

                // Generate slides
                int slideCount = SlideGenerator.GenerateSlides(presentation, content, format);
                MessageBox.Show($"Generated {slideCount} slides successfully!");
            }
            catch (Exception ex)
            {
                ErrorHandler.LogError("Failed to generate slides: " + ex.Message);
                MessageBox.Show("Error generating slides. Check the log for details.");
            }
        }
    }
}
