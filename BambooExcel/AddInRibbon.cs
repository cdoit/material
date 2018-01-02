using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using BambooExcel.Forms;
using Microsoft.Office.Interop.Excel;
using BambooExcel.Helpers;



namespace BambooExcel
{
    public partial class AddInRibbon
    {

        //TODO:  Need to unload all COM

        private void AddInRibbon_Load(object sender, RibbonUIEventArgs e)
        {
           
        }

        void ws_SelectionChange(Range Target)
        {
            try
            {
                if (Target.Text == "")
                    return;
                MessageBox.Show(Target.Text);
                Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
                Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
                Worksheet ws = null;
                if (wb == null)
                    return;
                Object temp;

                foreach (Worksheet wstemp in wb.Worksheets)
                {
                    ws = wstemp;
                    break;
                }
                if (ws == null)
                    return;
                int rowcount = ws.UsedRange.Rows.Count;
                int colcount = ws.UsedRange.Columns.Count;
                for (int i = 1; i <= rowcount; i++)
                {
                    for (int j = 1; j <= colcount; j++)
                    {
                        ((Range)ws.Cells[i, j]).Font.ColorIndex = 1; 
                    }
                }
                for (int i = 1; i <= rowcount; i++)
                {
                    for (int j = 1; j <= colcount; j++)
                    {
                        if (ws.Cells[i, j].Text == Target.Text)
                        {
                            ((Range)ws.Cells[i, j]).Font.ColorIndex = 3; 
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void btnImportFileList_Click(object sender, RibbonControlEventArgs e)
        {
            //frmFileListOptions frm = new frmFileListOptions();
            //frm.Show();

        }


        void ws_BeforeDoubleClick(Range Target, ref bool Cancel)
        {
            MessageBox.Show("d");
        }

        private void btnReplaceTextInFiles_Click(object sender, RibbonControlEventArgs e)
        {
            //Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            //Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
            //foreach(Worksheet ws in wb.Worksheets)
            //{
            //    ws.BeforeDoubleClick += ws_BeforeDoubleClick;
            //    ws.SelectionChange+=ws_SelectionChange;
            //}
            //MessageBox.Show("初始化成功");
        }

        private void btnUnprotectWB_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("12");
            Helpers.SecurityHelper.UnprotectWorkbook();
        }

        private void btnImportDirList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
                Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
                Worksheet ws = wb.Worksheets[2];
                if (wb == null)
                    return;
                Object temp;

                foreach (Worksheet wstemp in wb.Worksheets)
                {
                    ws = wstemp;
                    break;
                }
                if (ws == null)
                    return;
                int rowcount = ws.UsedRange.Rows.Count;
                int colcount = ws.UsedRange.Columns.Count;
                for (int i = 1; i <= rowcount; i++)
                {
                    for (int j = 1; j <= colcount; j++)
                    {
                        ((Range)ws.Cells[i, j]).Font.ColorIndex = 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDocExplorerPane_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {

        }


    }
}
