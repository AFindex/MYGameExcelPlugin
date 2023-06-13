using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelMyPlugin.UI;

namespace ExcelMyPlugin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookActivate += ApplicationOnWorkbookOpen;
        }

        private void ApplicationOnWorkbookOpen(object sh)
        {
            Excel.Workbook activeWb = ((Excel.Workbook)Application.ActiveWorkbook);
            Application_Workbook_SheetsActionReg(activeWb);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonTest();
        }
        
        void Application_Workbook_SheetsActionReg(Excel.Workbook wb)
        {
            wb.NewSheet += WbOnNewSheet;
            var activeWorkbook = wb;
            var currentSheets = activeWorkbook.Worksheets;
            foreach (Excel.Worksheet sheet in currentSheets)
            {
                sheet.Change -= SheetOnChange;
                sheet.Change += SheetOnChange;
            }
        }

        private void WbOnNewSheet(object sh)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)sh;
            sheet.Change += SheetOnChange;
        }

        private void SheetOnChange(Excel.Range target)
        {
            foreach (Excel.Range cell in target.Cells)
            {
                var currentCellVar = cell.Value2.ToString();
                if (!IsValueValidation(currentCellVar))
                    DisplayError();
                else
                    DisplayWarning();
            }
        }

        private bool IsValueValidation(string str)
        {
            if (str == "123")
            {
                return false;
            }
            return true;
        }

        private void DisplayWarning()
        {
            MessageBox.Show("Warning", "给你一个提醒",MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void DisplayError()
        {
            MessageBox.Show("Error", "给你一个错误提示",MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = "This text was added by using code";
        }
        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
