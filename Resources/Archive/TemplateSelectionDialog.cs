using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AIContentTool
{
    public partial class TemplateSelectionDialog : Form
    {
        private string selectedTemplatePath;

        public TemplateSelectionDialog()
        {
            InitializeComponent();
            UpdateTemplateDisplay();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "PowerPoint Templates (*.potx;*.pptx)|*.potx;*.pptx",
                Title = "Select a PowerPoint Template"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                selectedTemplatePath = dialog.FileName;
                UpdateTemplateDisplay();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Pass the selected path to your add-in logic (e.g., a static property)
            ImportHandlers.TemplateFilePath = selectedTemplatePath;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void UpdateTemplateDisplay()
        {
            lblSelectedTemplate.Text = string.IsNullOrEmpty(selectedTemplatePath)
                ? "No template selected"
                : selectedTemplatePath;
        }
    }
}