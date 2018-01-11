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
            XmlNode root = findroot(doc.FirstChild);
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
                string[] arry = nodetext.Split('【');
                if (arry.Length == 2)
                {
                    string name = arry[0];
                    string num = arry[1];
                    num = num.Replace("】", "");
                    loadlevel("-1", "", child);
                }

            }

            return bok;
        }

        private XmlNode findroot(XmlNode node)
        {
            XmlNode textnode = node.Attributes.GetNamedItem("TEXT");
            if (textnode != null && node.Attributes["TEXT"].Value == "中心主题")
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

        private string[] bianli(List<string[]> al)
        {
            if (al.Count == 0)
                return null;
            int size = 1;
            for (int i = 0; i < al.Count; i++)
            {
                size = size * al[i].Length;
            }
            string[] str = new string[size];
            for (int j = 0; j < size; j++)
            {
                for (int m = 0; m < al.Count; m++)
                {
                    str[j] = str[j] + al[m][(j * jisuan(al, m) / size) % al[m].Length] + " ";
                }
                str[j] = str[j].Trim(' ');
            }
            return str;
        }
        private int jisuan(List<string[]> al, int m)
        {
            int result = 1;
            for (int i = 0; i < al.Count; i++)
            {
                if (i <= m)
                {
                    result = result * al[i].Length;
                }
                else
                {
                    break;
                }
            }
            return result;
        }
        private void loadlevel(string parentid, string parentnum, XmlNode node)
        {
            //string id = node.Attributes["ID"].Value;
            string nodetext = node.Attributes["TEXT"].Value;
            nodetext = nodetext.Replace(" ", "");
            string[] arry = nodetext.Split('【');
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
                        string[] bianlianarry = bianli(res);
                        foreach (string onecell in bianlianarry)
                        {
                            string selfid = parentnum;
                            string[] cellarry = onecell.Split('/');
                            string code = "";
                            string attribute = "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><row>";
                            for (int i = 0; i < cellarry.Length - 2; i++)
                            {
                                attribute += "<field name=" + cellarry[i].Trim() + ">" + cellarry[i + 1].Trim() + "</field>";
                                selfid += cellarry[i + 2].Trim();
                                i = i + 2;
                            }
                            attribute += "</row>";
                            string codetemp = selfid.Substring(selfid.IndexOf(parentnum) + parentnum.Length);
                            selfid = selfid.PadRight(15, '0');
                            for (int j = 0; j < codetemp.Length; j++)
                            {
                                code += codetemp[j] + "_";
                            }
                            cmd = Application.instance().myConnection.CreateCommand();
                            cmd.CommandText = "INSERT  INTO material(id,categoryid,name,specifications,attributeinfo)VALUES ('" + selfid + "','" + parentnum + "','" + "" + "','" + code + "','" + attribute + "')";
                            cmd.ExecuteNonQuery();
                        }
                        return;
                    }
                }
            }

            if (node.HasChildNodes == false)
                return;
            foreach (XmlNode child in node.ChildNodes)
            {

                loadlevel(parentnum, parentnum, child);
            }
        }
    }
}
