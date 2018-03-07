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
using BambooExcel.matrail;
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
            Formlogin dlg = new Formlogin();
            dlg.ShowDialog();
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
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                if(dlg.ShowDialog()==DialogResult.OK)
                {
                    string file = dlg.FileName;
                    BaseMatrail matrail = new BaseMatrail();
                    matrail.import(file);
                    MessageBox.Show(file+"导入完成");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        //导入设计数据
        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            insertForm dlg = new insertForm();
            dlg.ShowDialog();
            
        }

        //导入提量数据
        private void toggleButton2_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("导入提量数据，这里默认是三级");
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
            Worksheet ws = null;
            if (wb == null)
                return;
            //获取第三个工作簿
            ws = wb.Worksheets[1];
            if (ws == null)
                return;

            //定义集合(工作簿1数据  用于存数据的坐标--数据信息)
            Dictionary<String, bamBean> dic = new Dictionary<String, bamBean>();
            int rowcount = ws.UsedRange.CurrentRegion.Rows.Count;
            //MessageBox.Show("行：" + rowcount);
            int colcount = ws.UsedRange.CurrentRegion.Columns.Count;
            //MessageBox.Show("列：" + colcount);

            //先取出表格第二行数据
            string projectname = ws.Cells[2, 2].Text;
            string projectcode = ws.Cells[2, 9].Text;
            string address = ws.Cells[2, 5].Text;
            string area = ws.Cells[2, 7].Text;

            project project = new project();
            string projectId = System.Guid.NewGuid().ToString("N");
            project.Id = projectId;
            project.Name = projectname;
            project.Code = projectcode;
            project.Address = address;
            project.Area = area;
            insertinto(project);

            //批量插入提量数据
            for (int i = 4; i <= rowcount; i++)
            {
                for (int j = 2; j <= colcount;j++ )
                {
                    //判断是否为构件
                    if(j == 2)
                    {
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text))
                        {
                            component component = new component();
                            String componentId = System.Guid.NewGuid().ToString("N");
                            component.Id = componentId;
                            component.Projectid = projectId;
                            component.Name = ws.Cells[i, j].Text;
                            saveComponent(component);
                            

                            //记录数据
                            bamBean bean = new bamBean();
                            bean.Id = componentId;
                            dic.Add(i + "-" + j, bean);
                        }
                        else
                        {
                            //记录数据
                            bamBean bean = new bamBean();
                            bean.Id = dic[(i - 1) + "-" + j].Id;
                            dic.Add(i + "-" + j, bean);
                        }
                    }
                    //第三列
                    else if (j == 3)
                    {
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text))
                        {
                            meterage meterage = new meterage();
                            String meterageId = System.Guid.NewGuid().ToString("N");
                            meterage.Id = meterageId;
                            // 从dic中取j-1中的数据
                            meterage.Componentid = dic[i + "-" + (j-1)].Id;
                            meterage.Parentid = "-1";
                            meterage.Materielcode = "code";
                            meterage.Materielname = ws.Cells[i, j].Text;
                            meterage.Specifications = ws.Cells[i, (j+2)].Text;
                            meterage.Count = ws.Cells[i, (j + 3)].Text;
                            meterage.Unit = ws.Cells[i, (j + 4)].Text;
                            meterage.Rule = ws.Cells[i, (j + 5)].Text;
                            meterage.Lossrate = "";
                            meterage.Memo = ws.Cells[i, (j + 6)].Text;
                            saveMeterage(meterage);

                            //记录数据
                            bamBean bean = new bamBean();
                            bean.Id = meterageId;
                            dic.Add(i + "-" + j, bean);

                        }
                        else
                        {
                            if(string.IsNullOrEmpty(ws.Cells[i, (j-1)].Text))
                            {
                                //记录数据
                                bamBean bean = new bamBean();
                                bean.Id = dic[(i - 1) + "-" + j].Id;
                                dic.Add(i + "-" + j, bean);
                            }
                            
                        }
                    }
                    else if(j == 4)
                    {
                        if (!string.IsNullOrEmpty(ws.Cells[i, j].Text))
                        {
                            meterage meterage = new meterage();
                            meterage.Id = System.Guid.NewGuid().ToString("N");
                            meterage.Componentid = dic[i + "-" + (j - 2)].Id;
                            meterage.Parentid = dic[i + "-" + (j - 1)].Id;
                            meterage.Materielcode = "code";
                            meterage.Materielname = ws.Cells[i, j].Text;
                            meterage.Specifications = ws.Cells[i, (j + 1)].Text;
                            meterage.Count = ws.Cells[i, (j + 2)].Text;
                            meterage.Unit = ws.Cells[i, (j + 3)].Text;
                            meterage.Rule = ws.Cells[i, (j + 4)].Text;
                            meterage.Lossrate = "";
                            meterage.Memo = ws.Cells[i, (j + 5)].Text;
                            saveMeterage(meterage);
                        }
                    }
                }
            }

            MessageBox.Show("导入成功！");

        }

        public void insertinto(BambooExcel.Mod.project project)
        {
            try
            {
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into project values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',NULL)", project.Id, project.Name, project.Code, project.Address, project.Area, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void saveComponent(BambooExcel.Mod.component component)
        {
            try
            {
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into materielcomponent values ('{0}','{1}','{2}','{3}','{4}',NULL)", component.Id, component.Projectid, component.Name, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void saveMeterage(BambooExcel.Mod.meterage meterage)
        {
            try
            {
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into meteragebill values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}',NULL)", meterage.Id, meterage.Componentid, meterage.Parentid, meterage.Materielcode, meterage.Materielname, meterage.Specifications, meterage.Count, meterage.Unit, meterage.Rule, meterage.Lossrate, meterage.Memo, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    }
}
