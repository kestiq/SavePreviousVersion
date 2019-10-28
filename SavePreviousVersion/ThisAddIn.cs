using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;

namespace SavePreviousVersion
{
    public partial class ThisAddIn
    {
        private const string DateFormat = "dd.MM.yyyy-HH.mm.ss";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            Application.StatusBar = "[SavePreviousVersion] Backup enabled.";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.StatusBar = "[SavePreviousVersion] Backup disabled.";
        }

        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Application.StatusBar = "[SavePreviousVersion] Creating a backup before saving...";

            string basePath;
            string fileName;
            string fileExt;
            string fileBasePath;

            try
            {
                basePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.BasePathExp);
                fileName = Path.GetFileNameWithoutExtension(Wb.FullName);
                fileExt = Path.GetExtension(Wb.FullName);
                fileBasePath = Path.Combine(basePath, fileName);

                if (!Directory.Exists(basePath))
                    Directory.CreateDirectory(basePath);

                if (!Directory.Exists(fileBasePath))
                    Directory.CreateDirectory(fileBasePath);

                Wb.SaveCopyAs(Path.Combine(fileBasePath, $"{fileName}({DateTime.Now.ToString(DateFormat)}){fileExt}"));

                Application.StatusBar = $"[SavePreviousVersion] Backup created at: {fileBasePath}";
            }
            catch (Exception e)
            {
                Application.StatusBar = $"[SavePreviousVersion] Unable to create backup: {e.Message}";
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbon();
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
