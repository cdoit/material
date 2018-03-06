namespace BambooExcel.Forms
{
    partial class Formlogin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txthost = new System.Windows.Forms.TextBox();
            this.txtusr = new System.Windows.Forms.TextBox();
            this.txtpwd = new System.Windows.Forms.TextBox();
            this.btok = new System.Windows.Forms.Button();
            this.btcancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.databaseText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txthost
            // 
            this.txthost.Location = new System.Drawing.Point(116, 44);
            this.txthost.Name = "txthost";
            this.txthost.Size = new System.Drawing.Size(136, 21);
            this.txthost.TabIndex = 0;
            this.txthost.Text = "localhost";
            // 
            // txtusr
            // 
            this.txtusr.Location = new System.Drawing.Point(116, 106);
            this.txtusr.Name = "txtusr";
            this.txtusr.Size = new System.Drawing.Size(136, 21);
            this.txtusr.TabIndex = 1;
            this.txtusr.Text = "test";
            // 
            // txtpwd
            // 
            this.txtpwd.Location = new System.Drawing.Point(116, 171);
            this.txtpwd.Name = "txtpwd";
            this.txtpwd.Size = new System.Drawing.Size(136, 21);
            this.txtpwd.TabIndex = 2;
            this.txtpwd.Text = "test";
            // 
            // btok
            // 
            this.btok.Location = new System.Drawing.Point(35, 270);
            this.btok.Name = "btok";
            this.btok.Size = new System.Drawing.Size(75, 23);
            this.btok.TabIndex = 3;
            this.btok.Text = "连接";
            this.btok.UseVisualStyleBackColor = true;
            this.btok.Click += new System.EventHandler(this.btok_Click);
            // 
            // btcancel
            // 
            this.btcancel.Location = new System.Drawing.Point(177, 270);
            this.btcancel.Name = "btcancel";
            this.btcancel.Size = new System.Drawing.Size(75, 23);
            this.btcancel.TabIndex = 4;
            this.btcancel.Text = "取消";
            this.btcancel.UseVisualStyleBackColor = true;
            this.btcancel.Click += new System.EventHandler(this.btcancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "数据库地址：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "用户名：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(33, 174);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "密码：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 225);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 9;
            this.label4.Text = "数据库：";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // databaseText
            // 
            this.databaseText.Location = new System.Drawing.Point(116, 222);
            this.databaseText.Name = "databaseText";
            this.databaseText.Size = new System.Drawing.Size(136, 21);
            this.databaseText.TabIndex = 8;
            this.databaseText.Text = "cdomaterial";
            this.databaseText.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // Formlogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 333);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.databaseText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btcancel);
            this.Controls.Add(this.btok);
            this.Controls.Add(this.txtpwd);
            this.Controls.Add(this.txtusr);
            this.Controls.Add(this.txthost);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Formlogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txthost;
        private System.Windows.Forms.TextBox txtusr;
        private System.Windows.Forms.TextBox txtpwd;
        private System.Windows.Forms.Button btok;
        private System.Windows.Forms.Button btcancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox databaseText;
    }
}