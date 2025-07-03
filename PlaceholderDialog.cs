using System;
using System.IO;
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
    public partial class PlaceholderDialog : Form
    {
        private readonly List<string> _placeholders;
        public Dictionary<string, string> PlaceholderFiles { get; private set; } = new Dictionary<string, string>();

        public PlaceholderDialog(List<string> placeholders)
        {
            _placeholders = placeholders;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Map Placeholders to Files";
            this.Width = 400;
            this.Height = 300;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            foreach (var placeholder in _placeholders)
            {
                var label = new Label { Text = placeholder, Width = 200, Margin = new Padding(5) };
                var textBox = new TextBox { Width = 150, Margin = new Padding(5), TabIndex = panel.Controls.Count };
                var button = new Button
                {
                    Text = "Browse",
                    Tag = textBox,
                    Margin = new Padding(5),
                    TabIndex = panel.Controls.Count + 1
                };
                button.Click += BrowseButton_Click;

                panel.Controls.Add(label);
                panel.Controls.Add(textBox);
                panel.Controls.Add(button);
            }

            var sharePointNote = new Label
            {
                Text = "SharePoint integration coming soon.",
                Width = 300,
                Margin = new Padding(5)
            };
            panel.Controls.Add(sharePointNote);

            var okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Margin = new Padding(5),
                TabIndex = panel.Controls.Count + 2
            };
            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(5),
                TabIndex = panel.Controls.Count + 3
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40
            };
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);

            this.Controls.Add(panel);
            this.Controls.Add(buttonPanel);

            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
            this.Load += (s, e) =>
            {
                foreach (Control ctrl in panel.Controls)
                {
                    if (ctrl is Button btn && btn.Text == "Browse")
                    {
                        var tb = btn.Tag as TextBox;
                        if (!string.IsNullOrEmpty(tb.Text) && File.Exists(tb.Text))
                        {
                            PlaceholderFiles[_placeholders[panel.Controls.IndexOf(btn) / 3]] = tb.Text;
                        }
                    }
                }
            };
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            var textBox = button.Tag as TextBox;
            OpenFileDialog dialog = new OpenFileDialog { Filter = "All Files (*.*)|*.*" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = dialog.FileName;
                int index = this.Controls[0].Controls.IndexOf(button) / 3;
                PlaceholderFiles[_placeholders[index]] = dialog.FileName;
            }
        }
    }
}