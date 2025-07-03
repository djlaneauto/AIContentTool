namespace AIContentTool
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnImportXml = this.Factory.CreateRibbonButton();
            this.btnImportJson = this.Factory.CreateRibbonButton();
            this.btnImportMarkdown = this.Factory.CreateRibbonButton();
            this.btnImportClipboard = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnGenerateSlides = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "AI Slide Generator";
            this.tab2.Name = "tab2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnImportXml);
            this.group1.Items.Add(this.btnImportJson);
            this.group1.Items.Add(this.btnImportMarkdown);
            this.group1.Items.Add(this.btnImportClipboard);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnImportXml
            // 
            this.btnImportXml.Label = "Import XML";
            this.btnImportXml.Name = "btnImportXml";
            this.btnImportXml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportXmlClick);
            // 
            // btnImportJson
            // 
            this.btnImportJson.Label = "Import JSON";
            this.btnImportJson.Name = "btnImportJson";
            this.btnImportJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportJsonClick);
            // 
            // btnImportMarkdown
            // 
            this.btnImportMarkdown.Label = "Import Markdown";
            this.btnImportMarkdown.Name = "btnImportMarkdown";
            this.btnImportMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImportMarkdownClick);
            // 
            // btnImportClipboard
            // 
            this.btnImportClipboard.Label = "Import from Clipboard";
            this.btnImportClipboard.Name = "btnImportClipboard";
            this.btnImportClipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnCopyFromClipboardClick);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnGenerateSlides);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btnGenerateSlides
            // 
            this.btnGenerateSlides.Label = "Generate Slides";
            this.btnGenerateSlides.Name = "btnGenerateSlides";
            this.btnGenerateSlides.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnGenerateSlidesClick);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportXml;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportMarkdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerateSlides;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
