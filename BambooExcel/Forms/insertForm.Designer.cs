namespace BambooExcel.Forms
{
    partial class insertForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.txt1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.text2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.text3 = new System.Windows.Forms.TextBox();
            this.textPlan = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textProject = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textStart = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textEnd = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btcancel = new System.Windows.Forms.Button();
            this.btok = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(72, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "预置信息：";
            // 
            // txt1
            // 
            this.txt1.Location = new System.Drawing.Point(142, 71);
            this.txt1.Name = "txt1";
            this.txt1.Size = new System.Drawing.Size(136, 21);
            this.txt1.TabIndex = 7;
            this.txt1.Text = "厨房";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(72, 127);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "继承信息：";
            // 
            // text2
            // 
            this.text2.Location = new System.Drawing.Point(142, 124);
            this.text2.Name = "text2";
            this.text2.Size = new System.Drawing.Size(136, 21);
            this.text2.TabIndex = 9;
            this.text2.Text = "天花";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(72, 181);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "继承信息2：";
            // 
            // text3
            // 
            this.text3.Location = new System.Drawing.Point(142, 178);
            this.text3.Name = "text3";
            this.text3.Size = new System.Drawing.Size(136, 21);
            this.text3.TabIndex = 11;
            this.text3.Text = "扣板吊顶";
            // 
            // textPlan
            // 
            this.textPlan.Location = new System.Drawing.Point(142, 226);
            this.textPlan.Name = "textPlan";
            this.textPlan.Size = new System.Drawing.Size(136, 21);
            this.textPlan.TabIndex = 13;
            this.textPlan.Text = "壹号院方案";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(72, 229);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "方案名称：";
            // 
            // textProject
            // 
            this.textProject.Location = new System.Drawing.Point(142, 29);
            this.textProject.Name = "textProject";
            this.textProject.Size = new System.Drawing.Size(136, 21);
            this.textProject.TabIndex = 15;
            this.textProject.Text = "民乐村";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(72, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "项目名称：";
            // 
            // textStart
            // 
            this.textStart.Location = new System.Drawing.Point(142, 273);
            this.textStart.Name = "textStart";
            this.textStart.Size = new System.Drawing.Size(53, 21);
            this.textStart.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(72, 276);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 16;
            this.label6.Text = "开始行号：";
            // 
            // textEnd
            // 
            this.textEnd.Location = new System.Drawing.Point(142, 317);
            this.textEnd.Name = "textEnd";
            this.textEnd.Size = new System.Drawing.Size(53, 21);
            this.textEnd.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(72, 320);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 18;
            this.label7.Text = "结束行号：";
            // 
            // btcancel
            // 
            this.btcancel.Location = new System.Drawing.Point(203, 376);
            this.btcancel.Name = "btcancel";
            this.btcancel.Size = new System.Drawing.Size(75, 23);
            this.btcancel.TabIndex = 21;
            this.btcancel.Text = "取消";
            this.btcancel.UseVisualStyleBackColor = true;
            this.btcancel.Click += new System.EventHandler(this.btcancel_Click);
            // 
            // btok
            // 
            this.btok.Location = new System.Drawing.Point(74, 376);
            this.btok.Name = "btok";
            this.btok.Size = new System.Drawing.Size(75, 23);
            this.btok.TabIndex = 20;
            this.btok.Text = "导入";
            this.btok.UseVisualStyleBackColor = true;
            this.btok.Click += new System.EventHandler(this.btok_Click);
            // 
            // insertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 437);
            this.Controls.Add(this.btcancel);
            this.Controls.Add(this.btok);
            this.Controls.Add(this.textEnd);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textStart);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textProject);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textPlan);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.text3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.text2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txt1);
            this.Controls.Add(this.label1);
            this.Name = "insertForm";
            this.Text = "insertForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox text2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox text3;
        private System.Windows.Forms.TextBox textPlan;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textProject;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textStart;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textEnd;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btcancel;
        private System.Windows.Forms.Button btok;
    }
}