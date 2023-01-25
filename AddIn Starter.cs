using System;
using System.Linq;
using Markup.Script;
using Excel = Microsoft.Office.Excel;
using Office = Microsoft.Office.Core;

namespace Markup
{
    /// <summary>
    /// Module that enables the add-in into Excel.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// module used to activate the add-in into the Excel properties.
        /// </summary>
        public static Excel.Application e_application;
        /// <summary>
        /// activate and include it into the ribbon section.
        /// </summary>
        public static Office.IRibbonUI e_ribbon;

        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            e_application = this.Application;
            e_application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(e_application_SheetSelectionChange);
        }

        /// <summary>
        /// the sheet evaluation and analysis is having a trigger moment while the user is interacting with the configuration.
        /// </summary>
        /// <param name="sh">name of the sheet.</param>
        /// <param name="target">selected range or cells.</param>
        /// <remarks></remarks>
        private void e_Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            // ribbon currently activated into the Excel interface of ribbons.
        }

        /// <summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            e_application.SheetSelectionChange -> new Excel.AppEvents_SheetSelectionChangeEventHandler(e_application_SheetSelectionChange);
            e_application = null;
            return false;
        }

        /// <summary>
        /// creation of the ribbon
        /// </summary>
        /// <returns></returns>
       private protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }
        #region VSTO generated code

        /// <summary>
        /// set-up of the add-in into the files of Office for the user.
        /// </summary>
        private void InternalStartup()
        {
            this.InternalStartup += new System.EventHandler(ThisAddIn_Startup);
            this.ThisAddIn_Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion 
    }
}