using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace AIContentTool
{
    public static class ImportHandlers
    {
        public static string ImportedContent { get; set; }
        public static string ContentFormat { get; set; }
        public static Dictionary<string, string> PlaceholderFiles { get; set; } = new Dictionary<string, string>();
        public static string TemplateFilePath { get; set; } // Stores the selected template file path

        public static void ImportXml()
        {
            OpenFileDialog dialog = new OpenFileDialog { Filter = "XML Files (*.xml)|*.xml" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ImportedContent = File.ReadAllText(dialog.FileName);
                ContentFormat = "XML";
                DetectPlaceholders(); // Detect placeholders immediately after import
            }
        }

        public static void ImportJson()
        {
            OpenFileDialog dialog = new OpenFileDialog { Filter = "JSON Files (*.json)|*.json" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ImportedContent = File.ReadAllText(dialog.FileName);
                ContentFormat = "JSON";
                DetectPlaceholders(); // Detect placeholders immediately after import
            }
        }

        public static void ImportMarkdown()
        {
            OpenFileDialog dialog = new OpenFileDialog { Filter = "Markdown Files (*.md)|*.md" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ImportedContent = File.ReadAllText(dialog.FileName);
                ContentFormat = "MARKDOWN";
                DetectPlaceholders(); // Detect placeholders immediately after import
            }
        }

        public static void ImportClipboard()
        {
            if (Clipboard.ContainsText())
            {
                string clipboardText = Clipboard.GetText().Trim(); // Trim to remove leading/trailing whitespace
                if (string.IsNullOrEmpty(clipboardText))
                {
                    MessageBox.Show("Clipboard does not contain text.");
                    return;
                }

                // Try to detect and parse as XML
                try
                {
                    XDocument.Parse(clipboardText);
                    ImportedContent = clipboardText;
                    ContentFormat = "XML";
                    DetectPlaceholders();
                    return;
                }
                catch (Exception) { /* Not XML, continue */ }

                // Try to detect and parse as JSON
                try
                {
                    JsonConvert.DeserializeObject(clipboardText);
                    ImportedContent = clipboardText;
                    ContentFormat = "JSON";
                    DetectPlaceholders();
                    return;
                }
                catch (Exception) { /* Not JSON, fallback to Markdown */ }

                // Fallback to Markdown
                ImportedContent = clipboardText;
                ContentFormat = "MARKDOWN";
                DetectPlaceholders();
            }
            else
            {
                MessageBox.Show("Clipboard does not contain text.");
            }
        }

        // Optional: Uncomment if you want to allow template selection via a dialog
        //public static bool PromptForTemplate()
        //{
        //    var dialog = new TemplateSelectionDialog();
        //    if (dialog.ShowDialog() == DialogResult.OK)
        //    {
        //        return true; // Template set in dialog
        //    }
        //    return false; // Cancelled
        //}

        //public static void SelectTemplate()
        //{
        //    OpenFileDialog dialog = new OpenFileDialog { Filter = "PowerPoint Files (*.pptx)|*.pptx" };
        //    if (dialog.ShowDialog() == DialogResult.OK)
        //    {
        //        TemplateFilePath = dialog.FileName;
        //        MessageBox.Show($"Template selected: {Path.GetFileName(TemplateFilePath)}");
        //    }
        //}

        // Optional: Keep ProcessImport for backward compatibility or other uses, but redefine it
        public static void ProcessImport()
        {
            DetectPlaceholders(); // Only handle placeholders and KPIs, no slide generation
        }

        private static void DetectPlaceholders()
        {
            PlaceholderFiles.Clear();
            var placeholders = PlaceholderManager.DetectPlaceholders(ImportedContent);
            if (placeholders.Count > 0)
            {
                using (var dialog = new PlaceholderDialog(placeholders))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        PlaceholderFiles = dialog.PlaceholderFiles;
                    }
                }
            }
            KPITracker.RecordImport(ContentFormat, placeholders.Count);
        }
    }
}