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
using MySql.Data.MySqlClient;
using BambooExcel.Mod;



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
            OpenFileDialog dlg = new OpenFileDialog();
            if(dlg.ShowDialog()==DialogResult.OK)
            {
                string file = dlg.FileName;
               
            }
        }

        

        //导入设计包
        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
                Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
                Worksheet ws = null;
                if (wb == null)
                    return;
                Object temp;
                //获取第一个工作簿
                ws = wb.Worksheets[2];


                //foreach (Worksheet wstemp in wb.Worksheets)
                //{
                //    ws = wstemp;
                //    break;
                //}


                if (ws == null)
                    return;

                //定义集合(工作簿1数据)
                Dictionary<int, List<String>> dic = new Dictionary<int, List<String>>();
                int rowcount = ws.UsedRange.CurrentRegion.Rows.Count;
                MessageBox.Show("列：" + rowcount);
                int colcount = ws.UsedRange.CurrentRegion.Rows.Count;
                MessageBox.Show("行：" + colcount);

                for (int i = 3; i <= colcount; i++)
                {
                    List<String> strList = new List<String>();
                    houseplan houseplan = new houseplan();
                    houseplan.Id = i + 14;
                    for (int j = 2; j <= rowcount; j++)
                    {
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text) && j != 4)
                        {
                            houseplan.Planname = "壹号院方案";
                            //strList.Add(ws.Cells[i, j].Text);
                        }
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text) && j == 2)
                        {
                            houseplan.Room = ws.Cells[i, j].Text;
                        }
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text) && j == 3)
                        {
                            houseplan.Part = ws.Cells[i, j].Text;
                        }
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text) && j == 5)
                        {
                            houseplan.Materialpackge1 = ws.Cells[i, j].Text;
                        }

                    }
                    //插数据
                    insertinto(houseplan);
                    dic.Add(i, strList);
                }

                //MessageBox.Show(list.Count + "");


                //工作簿3的数据
                //Worksheet ws3 = wb.Worksheets[2];
                //int ws3count = ws3.UsedRange.CurrentRegion.Rows.Count;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }





        //建立mysql数据库链接
        public static MySqlConnection getMySqlCon()
        {
            MySqlConnection mysql = null;
            try
            {
                String mysqlStr = "Database=cdomaterial;Data Source=127.0.0.1;User Id=root;Password=root;pooling=false;CharSet=utf8;port=3306";
                // String mySqlCon = ConfigurationManager.ConnectionStrings["MySqlCon"].ConnectionString;
                mysql = new MySqlConnection(mysqlStr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return mysql;
        }

        //建立执行命令语句对象
        public static MySqlCommand getSqlCommand(String sql, MySqlConnection mysql)
        {
            MySqlCommand mySqlCommand = null;
            try
            {
                mySqlCommand = new MySqlCommand(sql, mysql);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //  MySqlCommand mySqlCommand = new MySqlCommand(sql);
            // mySqlCommand.Connection = mysql;
            return mySqlCommand;
        }


        //添加数据
        public static void getInsert(MySqlCommand mySqlCommand)
        {
            try
            {
                mySqlCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                String message = ex.Message;
                Console.WriteLine("插入数据失败了！" + message);
                MessageBox.Show("插入数据失败了！" + message);
            }

        }



    }
}
