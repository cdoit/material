using BambooExcel.Helpers;
using BambooExcel.Mod;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BambooExcel.Forms
{
    public partial class insertForm : Form
    {
        public insertForm()
        {
            InitializeComponent();
        }

        //导入设计数据
        private void btok_Click(object sender, EventArgs e)
        {
            try
            {
                //处理对话框中的数据
                String textProject = this.textProject.Text;
                String txt1 = this.txt1.Text;
                String text2 = this.text2.Text;
                String text3 = this.text3.Text;
                String textPlan = this.textPlan.Text;
                //始末行号
                String textStart = this.textStart.Text;
                String textEnd = this.textEnd.Text;

                String houseId = findHouseId(textProject);
                if(string.IsNullOrEmpty(houseId)){
                    //保存house数据
                    house house = new house();
                    houseId = System.Guid.NewGuid().ToString("N");
                    house.Id = houseId;
                    house.Housename = textProject;
                    saveHouse(house);
                }

                //保存houseplan数据
                houseplan houseplan = new houseplan();
                String houseplanid = System.Guid.NewGuid().ToString("N");
                houseplan.Id = houseplanid;
                houseplan.Houseid = houseId;
                houseplan.Planname = textPlan;
                houseplan.Room = txt1;
                houseplan.Part = text2;
                houseplan.Materialpackge = text3;
                saveHousePlan(houseplan);

                Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
                Workbook wb = ExcelHelper.GetActiveWorkbook(true, excelApp);
                Worksheet ws = null;
                if (wb == null)
                    return;
                //获取第三个工作簿
                ws = wb.Worksheets[3];

                if (ws == null)
                    return;

                //定义集合(工作簿1数据  用于存数据的坐标--数据信息)
                Dictionary<String, bamBean> dic = new Dictionary<String, bamBean>();
                int rowcount = ws.UsedRange.CurrentRegion.Rows.Count;
                //MessageBox.Show("行：" + rowcount);
                int colcount = ws.UsedRange.CurrentRegion.Columns.Count;
                //MessageBox.Show("列：" + colcount);

                //处理行数
                int start = String.IsNullOrEmpty(textStart) ? 5 : int.Parse(textStart);
                int end = String.IsNullOrEmpty(textEnd) ? rowcount : int.Parse(textEnd);

                for (int i = start; i <= end; i++)
                {
                    //获取当前行的表达式，坐标为：（i，G）
                    String express = "0";
                    if (!string.IsNullOrEmpty(ws.Cells[i, 7].Text))
                    {
                        express = "G" + i + "=" + ws.Cells[i, 7].Text;
                    }
                    String unit = ws.Cells[i, 8].Text;
                    //获取公式
                    String gongshi = ws.Cells[i, 13].Formula;
                    for (int j = 2; j <= colcount; j++)
                    {
                        // 列7 ：表达式     列8：单位
                        if (j == 2 || j == 3 || j == 4 || j == 5 || j == 9 || j == 10)
                        {
                            String cName = ws.Cells[i, j].Text;
                            String id = System.Guid.NewGuid().ToString("N");

                            if (i == start)
                            {
                                //第一行一定是有值的，所以不用判断是否为空
                                designpackge designpackge = new designpackge();
                                //designpackge2Id = System.Guid.NewGuid().ToString("N");
                                designpackge.Id = id;
                                designpackge.Houseplanid = houseplanid;
                                designpackge.Nametype = "room";
                                designpackge.Cname = cName;
                                
                                designpackge.Unit = unit;
                                //designpackge.Parent_id = "-1";

                                //填入公式express
                                if (j == 3 || j == 4 || j == 5 || j == 9)
                                {
                                    designpackge.Expression = express;
                                }
                                else if(j == 10)
                                {
                                    if (string.IsNullOrEmpty(gongshi))
                                    {
                                        designpackge.Expression = "0";
                                    }
                                    else
                                    {
                                        designpackge.Expression = "H" + i + gongshi;
                                    }
                                    //填入物料编码
                                    designpackge.Materialid = ws.Cells[i, 11].Text;
                                }
                                else
                                {
                                    designpackge.Expression = "0";
                                }


                                String Parent_id = "";
                                if (j == 3 || j == 4 || j == 5 || j == 10)
                                {
                                    Parent_id = dic[i + "-" + (j - 1)].Id;
                                }
                                else if (j == 9)
                                {
                                    Parent_id = dic[i + "-" + 5].Id;
                                }
                                else if (j == 2)
                                {
                                    Parent_id = "-1";
                                }
                                designpackge.Parent_id = Parent_id;
                                insertinto(designpackge);


                                //记录数据
                                bamBean bean = new bamBean();
                                bean.Id = id;
                                bean.Name = cName;
                                bean.Parentid = Parent_id;
                                dic.Add(i + "-" + j, bean);


                            }
                            else if (i > start)
                            {
                                //大于第一层
                                if (!string.IsNullOrEmpty(cName))
                                {
                                    designpackge designpackge = new designpackge();
                                    //designpackge2Id = System.Guid.NewGuid().ToString("N");
                                    designpackge.Id = id;
                                    designpackge.Houseplanid = houseplanid;
                                    designpackge.Nametype = "room";
                                    designpackge.Cname = cName;
                                    //designpackge.Expression = express;
                                    designpackge.Unit = unit;
                                    //designpackge.Parent_id = "-1";

                                    //填入公式express
                                    if (j == 5)
                                    {
                                        designpackge.Expression = express;
                                    }
                                    else if (j == 10)
                                    {
                                        if (string.IsNullOrEmpty(gongshi))
                                        {
                                            designpackge.Expression = "0";
                                        }
                                        else 
                                        {
                                            designpackge.Expression = "H" + i + gongshi;
                                        }
                                        //填入物料编码
                                        designpackge.Materialid = ws.Cells[i, 11].Text;
                                    }
                                    else
                                    {
                                        designpackge.Expression = "0";
                                    }

                                    String Parent_id = "";
                                    if (j == 3 || j == 4 || j == 5 || j == 10)
                                    {
                                        if (String.IsNullOrEmpty(ws.Cells[i, j - 1].Text))
                                        {
                                            Parent_id = dic[(i - 1) + "-" + j].Parentid;
                                        }
                                        else
                                        {
                                            Parent_id = dic[i + "-" + (j - 1)].Id;
                                        }

                                    }
                                    else if (j == 9)
                                    {
                                        if (String.IsNullOrEmpty(ws.Cells[i, j - 1].Text))
                                        {
                                            Parent_id = dic[(i - 1) + "-" + j].Parentid;
                                        }
                                        else
                                        {
                                            Parent_id = dic[i + "-" + 5].Id;
                                        }

                                    }
                                    else if (j == 2)
                                    {
                                        Parent_id = "-1";
                                    }



                                    designpackge.Parent_id = Parent_id;
                                    insertinto(designpackge);


                                    //记录数据
                                    bamBean bean = new bamBean();
                                    bean.Id = id;
                                    bean.Name = cName;
                                    bean.Parentid = Parent_id;
                                    dic.Add(i + "-" + j, bean);
                                }
                                else
                                {
                                    // 说明这里只需要存dictionary就行了  
                                    //读取自己的name以及父级节点
                                    bamBean bean = new bamBean();
                                    bean.Id = id;
                                    //bean.Name = ws.Cells[i - 1, j].Text;
                                    bean.Name = dic[(i - 1) + "-" + j].Name;

                                    if (j == 3 || j == 4 || j == 5 || j == 9 || j == 10)
                                    {
                                        bean.Parentid = dic[(i - 1) + "-" + j].Parentid;
                                    }
                                    else if (j == 2)
                                    {
                                        bean.Parentid = "-1";
                                    }
                                    dic.Add(i + "-" + j, bean);
                                }
                            }
                        }
                    }
                }
                //关闭窗口
                this.Close();
                MessageBox.Show("导入成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btcancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }





        public void insertinto(BambooExcel.Mod.designpackge designpackge)
        {
            try
            {
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into designpackge values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", designpackge.Id, designpackge.Houseplanid, designpackge.Nametype, designpackge.Cname, designpackge.Expression, designpackge.Unit, designpackge.Parent_id, designpackge.Materialid);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }





        public void saveHouse(BambooExcel.Mod.house house)
        {
            try
            {
                //MySqlConnection mysql = getMySqlCon();
                ////插入sql
                //string sqlInsert = string.Format("insert into house values ('{0}','{1}')", house.Id, house.Housename);
                //MySqlCommand mySqlCommand = getSqlCommand(sqlInsert, mysql);
                //mysql.Open();
                //getInsert(mySqlCommand);
                ////记得关闭
                //mysql.Close();

                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into house values ('{0}','{1}')", house.Id, house.Housename);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void saveHousePlan(BambooExcel.Mod.houseplan houseplan)
        {
            try
            {
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlInsert = string.Format("insert into houseplan values ('{0}','{1}','{2}','{3}','{4}','{5}')", houseplan.Id, houseplan.Houseid, houseplan.Planname, houseplan.Room, houseplan.Part, houseplan.Materialpackge);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        public String findHouseId(String houseName)
        {
            String houseId = null;
            try
            {
                //MySqlConnection mysql = getMySqlCon();
                ////插入sql
                //string sqlselect = string.Format("select * from house where housename ='{0}'", houseName);
                //Console.WriteLine(sqlselect);
                ////四种语句对象
                //MySqlCommand mySqlCommand = getSqlCommand(sqlselect, mysql);
                //mysql.Open();
                //MySqlDataReader reader = mySqlCommand.ExecuteReader();


                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                string sqlselect = string.Format("select * from house where housename ='{0}'", houseName);
                cmd.CommandText = sqlselect;
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        houseId = reader.GetString(0);
                        break;
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return houseId;
        }


        public MySqlDataReader select(String cName)
        {
            MySqlDataReader reader = null;
            try
            {   
                MySqlConnection mysql = getMySqlCon();
                //插入sql
                string sqlselect = string.Format("select * from designpackge where cname ='{0}'", cName);
                //String sqlInsert = "insert into houseplan values (12,1,'planname','room','part','materialpackge')";
                //打印SQL语句
                Console.WriteLine(sqlselect);
                //四种语句对象
                MySqlCommand mySqlCommand = getSqlCommand(sqlselect, mysql);

                mysql.Open();
                reader = mySqlCommand.ExecuteReader();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return reader;
        }


        //建立mysql数据库链接
        public static MySqlConnection getMySqlCon()
        {
            MySqlConnection mysql = null;
            try
            {
                String mysqlStr = "Database=cdomaterial;Data Source=192.168.31.242;User Id=root;Password=root;pooling=false;CharSet=utf8;port=3306";
                //String mySqlCon = ConfigurationManager.ConnectionStrings["MySqlCon"].ConnectionString;
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
            return mySqlCommand;
        }


        //  添加数据
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


        //查询并获得结果集并遍历
        public static void getResultset(MySqlCommand mySqlCommand)
        {
            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        Console.WriteLine("编号:" + reader.GetInt32(0) + "|姓名:" + reader.GetString(1) + "|年龄:" + reader.GetInt32(2) + "|学历:" + reader.GetString(3));
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("查询失败了！");
            }
            finally
            {
                reader.Close();
            }
        }



        public static string ToName(int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }

    }
}
