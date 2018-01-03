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
    public partial class Formlogin : Form
    {
        public Formlogin()
        {
            InitializeComponent();
        }


        private void btok_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection myConnection = new MySqlConnection("server=" + this.txthost.Text + ";uid = " + this.txtusr.Text + ";pwd = " + this.txtpwd.Text + ";database=cdomaterial");
                myConnection.Open();
                if (myConnection.State == ConnectionState.Open)
                {
                    Application.instance().myConnection = myConnection;
                    MessageBox.Show("连接成功");
                    this.Close();
                }
                else
                {
                    MessageBox.Show("连接不成功");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btcancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
