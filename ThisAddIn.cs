using System;
using System.Collections.Generic;
using System.Drawing;
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
            //this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            //this.Application.WorkbookActivate += ApplicationOnWorkbookOpen;
            this.Application.SheetSelectionChange += ApplicationOnSheetSelectionChange ;
            this.Application.SheetChange += ApplicationOnSheetChange;
        }

        private void ApplicationOnSheetSelectionChange(object sh, Excel.Range target)
        {
            var currentSheet = (Excel.Worksheet)sh;
            var sheetName = currentSheet.Name.ToString();
            if (Utility.IsContainChinese(sheetName))
            {
                WidgetPresets.CloseInfoInTargetRange();
            }
            else
            {
                WidgetPresets.ShowInfoInTargetRange(target,"测试提示标题","测试提示123");
            }
            
        }


        private void ApplicationOnSheetChange(object sh, Excel.Range target)
        {
            bool shouldWraning = false;
            bool shouldError = false;
            var currentSheet = (Excel.Worksheet)sh;
            var sheetName = currentSheet.Name.ToString();
            if (Utility.IsContainChinese(sheetName))
            {
                return;
            }
            
            foreach (Excel.Range cell in target.Cells)
            {
                var currentCellVar = cell.Value2.ToString();
                if (!IsValueValidation(currentCellVar))
                {
                    //cell.Value2 = "不准涩涩!";
                    shouldError = true;
                }
                else
                    shouldWraning = true;
            }

            if (shouldError)
            {
                WidgetPresets.DisplayError("错误","123了");
            }

            if (shouldWraning)
            {
                WidgetPresets.DisplayWarning("警告","测试了");
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonTest();
        }

        private bool IsValueValidation(string str)
        {
            if (str == "123")
            {
                return false;
            }
            return true;
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
