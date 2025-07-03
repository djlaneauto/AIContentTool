using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AIContentTool
{
    public partial class ThisAddIn
    {
        public static PowerPoint.Application PpApp { get; private set; }
        public static PowerPoint.Presentation CurrentPresentation { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                MessageBox.Show("Add-in starting up");
                PpApp = this.Application;
                MessageBox.Show("PpApp set to " + (PpApp != null ? "not null" : "null"));
                PpApp.PresentationOpen += (PowerPoint.Presentation pres) => CurrentPresentation = pres;
                PpApp.AfterNewPresentation += (PowerPoint.Presentation pres) => CurrentPresentation = pres;
                PpApp.WindowActivate += (PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn) => CurrentPresentation = pres;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in startup: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
