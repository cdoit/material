﻿using System;
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
            string id = doc.ChildNodes[1].FirstChild.Attributes["ID"].Value;
            loadlevel("-1", doc.ChildNodes[1].FirstChild);
            Application.instance().myConnection.Close();
            return bok;
        }

        private void loadlevel(string parentid,XmlNode node)
        {
            string id = node.Attributes["ID"].Value;
            string nodetext= node.Attributes["TEXT"].Value;
            string[] arry=nodetext.Split('【');
            if (arry.Length==2)
            {
                string name=arry[0];
                string num = arry[1];
                MySqlCommand cmd = Application.instance().myConnection.CreateCommand();
                cmd.CommandText = "INSERT  INTO materialcategory(id,name,parentid,codelength)VALUES ('" + id + "','" + name + "','"+parentid+"',1)";
                cmd.ExecuteNonQuery();
            }
            if(node.HasChildNodes == false)
                return;
            foreach(XmlNode child in node.ChildNodes)
            {

                loadlevel(id,child);
            }
        }
    }
}
