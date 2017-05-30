using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace EAI_SplitToRow
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void SplitToRow()
        {
            Excel.Application ExcelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet ActiveWSD = ExcelApp.ActiveSheet;
            Excel.Range SelectedRng = ExcelApp.Selection;

            //Select Impact Range, cannot multi area
            Excel.Range ImpactRng;
            RetryImpactRng:
            try
            {
                ImpactRng = ExcelApp.InputBox("Please select range of your data", "Split To Row", Type: 8, Default: SelectedRng.Address[0,0]);
            }
            catch (Exception)
            {
                return;
            }
            if (ImpactRng.Areas.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("You can Only select 1 Area");
                goto RetryImpactRng;
            }

            //Select Split Column
            Excel.Range SplitRng;
            RetrySplitRng:
            try
            {
                SplitRng = ExcelApp.InputBox("Please select cells you want to split to Row", "Split To Row", Type: 8, Default: SelectedRng.Address[0, 0]);
            }
            catch (Exception)
            {
                return;
            }
            if (SplitRng.Areas.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("You can Only select 1 Area");
                goto RetrySplitRng;
            }

            //Get Start Row & End Row
            //sRow = splitColumn.Find("*", SearchOrder:= xlByRows, SearchDirection:= xlNext).Row - 1
            //eRow = splitColumn.Find("*", SearchOrder:= xlByRows, SearchDirection:= xlPrevious).Row
            int sRow = SplitRng.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext).Row - 1;
            int eRow = SplitRng.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row;
            int nCol = SplitRng.Column;
            SplitRng = ExcelApp.Range[ExcelApp.Cells[sRow, nCol], ExcelApp.Cells[eRow, nCol]];


            //Get Splitter

            //Define Split Point

            //Loop Split Column
            //Split Text in Split Point, r = count of split
            //Copy & Insert Range


            ActiveWSD.Range["A1"].Value = ImpactRng.Address;
            ActiveWSD.Range["A2"].Value = SplitRng.Address;
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
