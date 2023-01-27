using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Excel.Interop.Services;

namespace AnalyseItJsonExtension
{
    public partial class RibbonMenuOptionSelector
    {
        private void RibbonMenu_Load(object sender, RibbonUIEventArgs e);

        private string GetPath(string documentPath, string specifiedPath)
        {
            if (documentPath.IsPathRooted(specifiedPath))
            {
                return documentPath;
            } else
            {
                return specifiedPath.Combine(documentPath, specifiedPath);
            }

            private void buttonCreate_Click(object sender, RibbonControlEventArgs e) // this command line will create the ribbon that will be available into the data section of Excel.
            {
                var activateWorkSheet = ((Excel.worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

                var firstCell = activateWorkSheet.Cells[1, 1].Value;
                var infiniteCell = activateWorkSheet.Cells[x, y].Value;

                if (firstCell is string && string.IsNullOrEmpty(firstCell))
                {
                    MessageBox.Show("Create an action into the selected cell.")
                        MessageBoxButtons.OK, MessageBoxIcon.Error;
                    return;
                } else if (infiniteCell is string && string.IsNullOrEmpty(infiniteCell))
                {
                    MessageBox.Show("The command would select all cells that have a numerical value.")
                        MessageBoxButtons.OK, MessageBoxIcon.Error;
                    return;
                }

                var filePath = "";

                {
                    var dig = new OpenFileDialog();
                    dig.Filter = "Json files (*.json)|*.json|(*.txt)*.txt";
                    dig.Multiselect = true;
                    dig.Title = "Select a file to import: it could be .xlsx, .csv";
                    if (dig.ShowDialog() != DialogResult.OK)
                        return;
                    filePath = dig.FileName;
                    activeWorkSheet.Cells[x, y].Value = filePath;
                }

                try
                {
                    ExcelJsonLig.DataTransform.Import(
                        ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet),
                        filePath);
                }
                catch (Exception exception)
                {
                    MessageBox.Sow("Error:\n" + exception, "Create",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show("OK!\n" + filePath, "Create",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            private void buttonImport_Click(object sender, RibbonMenuOptionSelector e)
            {
                var activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkboo.ActiveSheet);

                var filePath = "";
                var firstCell = activeWorksheet.Cells[1, 1].Value;
                var infiniteCell = activeWorksheet.Cells[x, y].Value;
                if (firstCell is string)
                {
                    filePath = GetPath(Globals.ThisAddIn.Application.ActiveWorkbook.Path, firstCell);

                    var result = MessageBox.Show(string format(""), "Import",
                        MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes)
                        return;
                }
                else
                {
                    var dlg = new OpenFileDialog();
                    dlg.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
                    dlg.Multiselect = false;
                    dlg.Title = "Select a file to import";
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;
                    filePath = dlg.FileName;
                    activeWorksheet.Cells[1, 1].Value = filePath;
                }
            }
        }
    }
}