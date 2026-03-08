using System;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace SDTD
{
    public partial class ThisAddIn
    {
        internal Word.Application WordApp => Application;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SdtdRibbon(this);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e) { }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
