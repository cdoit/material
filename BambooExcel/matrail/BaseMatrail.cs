using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace BambooExcel.matrail
{
    class BaseMatrail
    {
        public bool import(string file)
        {
            bool bok = false;
           
            XmlDocument doc = new System.Xml.XmlDocument();
            doc.Load(file);
            XmlNode root=findroot(doc.FirstChild);
            if (MessageBox.Show("清空物料基础数据，是否继续？", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
                return false;
            MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
            cmd.CommandText = "delete  from material";
            cmd.ExecuteNonQuery();
            cmd = Application.instance().myConnection.CreateCommand();
            cmd.CommandText = "delete  from materialcategory";
            cmd.ExecuteNonQuery();

            foreach (XmlNode child in root.ChildNodes)
            {
                string id = child.Attributes["ID"].Value;
                string nodetext = child.Attributes["TEXT"].Value;
                string[] arry=nodetext.Split('【');
                if (arry.Length == 2)
                {
                    string name = arry[0];
                    string num = arry[1];
                    num = num.Replace("】","");
                    loadlevel("-1","", child);
                }
                
            }
             
            return bok;
        }

        private XmlNode findroot(XmlNode node)
        {
            XmlNode textnode=node.Attributes.GetNamedItem("TEXT");
            if (textnode !=null&& node.Attributes["TEXT"].Value == "中心主题")
            {
                return node;
            }
            else
            {
                foreach (XmlNode child in node.ChildNodes)
                    return findroot(child);
            }
            return null;
        }

      
        private void loadlevel(string parentid,string parentnum,XmlNode node)
        {
            //string id = node.Attributes["ID"].Value;
            string nodetext= node.Attributes["TEXT"].Value;
            nodetext=nodetext.Replace(" ", "");
            string[] arry=nodetext.Split('【');
            if (arry.Length == 2)
            {
                string name = arry[0];
                string num = arry[1];
                num = num.Replace("】", "");
                parentnum = parentnum + num;
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                cmd.CommandText = "INSERT  INTO materialcategory(id,name,parentid,codelength)VALUES ('" + parentnum + "','" + name + "','" + parentid + "',1)";
                cmd.ExecuteNonQuery();

                if (node.HasChildNodes)
                {
                    string childnametemp = node.ChildNodes[0].Attributes["TEXT"].Value;
                    string[] childarry = childnametemp.Split('【');
                    if (childarry.Length != 2)
                    {
                        List<string[]> res = new List<string[]>();
                        foreach (XmlNode rootchild in node.ChildNodes)
                        {
                            childnametemp = rootchild.Attributes["TEXT"].Value;
                            childarry = childnametemp.Split('【');

                            if (childarry.Length != 2)
                            {
                                List<string> nodestr = new List<string>();
                                foreach (XmlNode basechild in rootchild.ChildNodes)
                                {
                                    string childnodetext = basechild.Attributes["TEXT"].Value;
                                    childnodetext = childnodetext.Replace(" ", "");
                                    string[] childarrynext = childnodetext.Split('【');
                                    string selfnum = childarrynext[1].Replace("】", "");
                                    if (childarrynext.Length == 2)
                                    {
                                        nodestr.Add(childnametemp.Trim() + "/" + childarrynext[0].Trim() + "/" + selfnum.Trim() + "/");
                                    }

                                }
                                res.Add(nodestr.ToArray());

                            }
                        }
                        return;
                    }
                }
            }
           
            if(node.HasChildNodes == false)
                return;
            foreach(XmlNode child in node.ChildNodes)
            {

                loadlevel(parentnum, parentnum, child);
            }
        }
    }
}
