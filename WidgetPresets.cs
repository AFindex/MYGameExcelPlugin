using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelMyPlugin.UI;

namespace ExcelMyPlugin
{
    public class WidgetPresets
    {
        static Form popup = new Form();

        public static void CloseInfoInTargetRange()
        {
            if(popup != null)
                popup.Close();
        }
        
        public static void ShowInfoInTargetRange(Excel.Range target,string title ,string infoContent)
        {
            var range = target;
            
            int zoom = Convert.ToInt32(range.Application.ActiveWindow.Zoom);
            int relativeLeft = 0;
            int relativeTop = 0;
            for (int index = 1; index < range.Column; index++)
            {
                relativeLeft += Convert.ToInt32(Math.Round(
                    (double)((Excel.Range)range.Worksheet.Cells[1, index]).Width
                    * 4 * zoom / 300,
                    MidpointRounding.AwayFromZero));
            }
            int left = range.Application.ActiveWindow.PointsToScreenPixelsX(relativeLeft);
            // The formula for row height is similar to the one for column width.
            for (int index = 1; index <= range.Row; index++)
            {
                relativeTop += Convert.ToInt32(Math.Round(
                    (double)((Excel.Range)range.Worksheet.Cells[index, 1]).Height
                    * 4 * zoom / 300,
                    MidpointRounding.AwayFromZero));
            }
            int Top = range.Application.ActiveWindow.PointsToScreenPixelsY(relativeTop);
            
            var LabelTitle = new Label();
            LabelTitle.Text = title;
            LabelTitle.Top = 0;
            LabelTitle.AutoSize = true;
            
            var LabelLine = new Label();
            LabelLine.Text = "";
            LabelLine.AutoSize = true;
            LabelLine.BorderStyle = BorderStyle.Fixed3D;
            LabelLine.ForeColor = Color.Brown;
            LabelLine.BackColor = Color.Brown;
            LabelLine.Height = 4;
            LabelLine.Top = 12;
            
            var LabelInfo = new Label();
            LabelInfo.Text = infoContent;
            LabelInfo.AutoSize = true;
            LabelInfo.Top = 24;

            popup.Close();
            popup = new Form();
            popup.StartPosition = FormStartPosition.Manual;
            popup.ControlBox = false;
            popup.TopMost = true;
            popup.FormBorderStyle = FormBorderStyle.None;
            popup.BackColor = Color.Cornsilk;
            popup.Left = left;
            popup.Top = Top;
            popup.Width = 50;
            popup.Height = 70;
            
            popup.Controls.Add(LabelTitle);
            popup.Controls.Add(LabelLine);
            popup.Controls.Add(LabelInfo);
            
            popup.Show();
        }
        
        
        public static void DisplayWarning(string title, string content)
        {
            MessageBox.Show(title, content,MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void DisplayError(string title, string content)
        {
            MessageBox.Show(title, content,MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        
    }
}
