namespace EquReportToWord
{
    partial class FrmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_Start = new System.Windows.Forms.Button();
            this.btn_Close = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textPath = new System.Windows.Forms.TextBox();
            this.txt_select = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.labE6 = new System.Windows.Forms.Label();
            this.labF6 = new System.Windows.Forms.Label();
            this.labM6 = new System.Windows.Forms.Label();
            this.labE5 = new System.Windows.Forms.Label();
            this.labF5 = new System.Windows.Forms.Label();
            this.labM5 = new System.Windows.Forms.Label();
            this.labE4 = new System.Windows.Forms.Label();
            this.labF4 = new System.Windows.Forms.Label();
            this.labE3 = new System.Windows.Forms.Label();
            this.labF3 = new System.Windows.Forms.Label();
            this.labE2 = new System.Windows.Forms.Label();
            this.labF2 = new System.Windows.Forms.Label();
            this.labE1 = new System.Windows.Forms.Label();
            this.labF1 = new System.Windows.Forms.Label();
            this.labM4 = new System.Windows.Forms.Label();
            this.labM3 = new System.Windows.Forms.Label();
            this.labM2 = new System.Windows.Forms.Label();
            this.labM1 = new System.Windows.Forms.Label();
            this.labPercent = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(168, 444);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(87, 27);
            this.btn_Start.TabIndex = 0;
            this.btn_Start.Text = "开始生成";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // btn_Close
            // 
            this.btn_Close.Location = new System.Drawing.Point(417, 444);
            this.btn_Close.Name = "btn_Close";
            this.btn_Close.Size = new System.Drawing.Size(87, 27);
            this.btn_Close.TabIndex = 1;
            this.btn_Close.Text = "关闭";
            this.btn_Close.UseVisualStyleBackColor = true;
            this.btn_Close.Click += new System.EventHandler(this.btn_Close_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.textPath);
            this.groupBox2.Controls.Add(this.txt_select);
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(710, 75);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "报告生成路径";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 33);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(105, 14);
            this.label4.TabIndex = 11;
            this.label4.Text = "文档保存位置：";
            // 
            // textPath
            // 
            this.textPath.Location = new System.Drawing.Point(141, 30);
            this.textPath.Name = "textPath";
            this.textPath.ReadOnly = true;
            this.textPath.Size = new System.Drawing.Size(432, 23);
            this.textPath.TabIndex = 5;
            // 
            // txt_select
            // 
            this.txt_select.Location = new System.Drawing.Point(583, 30);
            this.txt_select.Name = "txt_select";
            this.txt_select.Size = new System.Drawing.Size(79, 26);
            this.txt_select.TabIndex = 7;
            this.txt_select.Text = "浏览";
            this.txt_select.UseVisualStyleBackColor = true;
            this.txt_select.Click += new System.EventHandler(this.txt_select_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.labE6);
            this.groupBox3.Controls.Add(this.labF6);
            this.groupBox3.Controls.Add(this.labM6);
            this.groupBox3.Controls.Add(this.labE5);
            this.groupBox3.Controls.Add(this.labF5);
            this.groupBox3.Controls.Add(this.labM5);
            this.groupBox3.Controls.Add(this.labE4);
            this.groupBox3.Controls.Add(this.labF4);
            this.groupBox3.Controls.Add(this.labE3);
            this.groupBox3.Controls.Add(this.labF3);
            this.groupBox3.Controls.Add(this.labE2);
            this.groupBox3.Controls.Add(this.labF2);
            this.groupBox3.Controls.Add(this.labE1);
            this.groupBox3.Controls.Add(this.labF1);
            this.groupBox3.Controls.Add(this.labM4);
            this.groupBox3.Controls.Add(this.labM3);
            this.groupBox3.Controls.Add(this.labM2);
            this.groupBox3.Controls.Add(this.labM1);
            this.groupBox3.Controls.Add(this.labPercent);
            this.groupBox3.Controls.Add(this.progressBar1);
            this.groupBox3.Location = new System.Drawing.Point(6, 100);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(710, 328);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "报告生成进度";
            // 
            // labE6
            // 
            this.labE6.AutoSize = true;
            this.labE6.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE6.ForeColor = System.Drawing.Color.Red;
            this.labE6.Location = new System.Drawing.Point(426, 293);
            this.labE6.Name = "labE6";
            this.labE6.Size = new System.Drawing.Size(84, 14);
            this.labE6.TabIndex = 16;
            this.labE6.Text = "正在进行...";
            this.labE6.Visible = false;
            // 
            // labF6
            // 
            this.labF6.AutoSize = true;
            this.labF6.Enabled = false;
            this.labF6.Location = new System.Drawing.Point(137, 293);
            this.labF6.Name = "labF6";
            this.labF6.Size = new System.Drawing.Size(21, 14);
            this.labF6.TabIndex = 17;
            this.labF6.Text = "6.";
            // 
            // labM6
            // 
            this.labM6.AutoSize = true;
            this.labM6.Enabled = false;
            this.labM6.Location = new System.Drawing.Point(177, 293);
            this.labM6.Name = "labM6";
            this.labM6.Size = new System.Drawing.Size(91, 14);
            this.labM6.TabIndex = 15;
            this.labM6.Text = "写入演练总结";
            // 
            // labE5
            // 
            this.labE5.AutoSize = true;
            this.labE5.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE5.ForeColor = System.Drawing.Color.Red;
            this.labE5.Location = new System.Drawing.Point(426, 253);
            this.labE5.Name = "labE5";
            this.labE5.Size = new System.Drawing.Size(84, 14);
            this.labE5.TabIndex = 13;
            this.labE5.Text = "正在进行...";
            this.labE5.Visible = false;
            // 
            // labF5
            // 
            this.labF5.AutoSize = true;
            this.labF5.Enabled = false;
            this.labF5.Location = new System.Drawing.Point(137, 253);
            this.labF5.Name = "labF5";
            this.labF5.Size = new System.Drawing.Size(21, 14);
            this.labF5.TabIndex = 14;
            this.labF5.Text = "5.";
            // 
            // labM5
            // 
            this.labM5.AutoSize = true;
            this.labM5.Enabled = false;
            this.labM5.Location = new System.Drawing.Point(177, 253);
            this.labM5.Name = "labM5";
            this.labM5.Size = new System.Drawing.Size(133, 14);
            this.labM5.TabIndex = 12;
            this.labM5.Text = "写入各单位成绩明细";
            // 
            // labE4
            // 
            this.labE4.AutoSize = true;
            this.labE4.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE4.ForeColor = System.Drawing.Color.Red;
            this.labE4.Location = new System.Drawing.Point(426, 211);
            this.labE4.Name = "labE4";
            this.labE4.Size = new System.Drawing.Size(84, 14);
            this.labE4.TabIndex = 11;
            this.labE4.Text = "正在进行...";
            this.labE4.Visible = false;
            // 
            // labF4
            // 
            this.labF4.AutoSize = true;
            this.labF4.Enabled = false;
            this.labF4.Location = new System.Drawing.Point(137, 211);
            this.labF4.Name = "labF4";
            this.labF4.Size = new System.Drawing.Size(21, 14);
            this.labF4.TabIndex = 11;
            this.labF4.Text = "4.";
            // 
            // labE3
            // 
            this.labE3.AutoSize = true;
            this.labE3.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE3.ForeColor = System.Drawing.Color.Red;
            this.labE3.Location = new System.Drawing.Point(426, 168);
            this.labE3.Name = "labE3";
            this.labE3.Size = new System.Drawing.Size(84, 14);
            this.labE3.TabIndex = 11;
            this.labE3.Text = "正在进行...";
            this.labE3.Visible = false;
            // 
            // labF3
            // 
            this.labF3.AutoSize = true;
            this.labF3.Enabled = false;
            this.labF3.Location = new System.Drawing.Point(137, 168);
            this.labF3.Name = "labF3";
            this.labF3.Size = new System.Drawing.Size(21, 14);
            this.labF3.TabIndex = 11;
            this.labF3.Text = "3.";
            // 
            // labE2
            // 
            this.labE2.AutoSize = true;
            this.labE2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE2.ForeColor = System.Drawing.Color.Red;
            this.labE2.Location = new System.Drawing.Point(426, 125);
            this.labE2.Name = "labE2";
            this.labE2.Size = new System.Drawing.Size(84, 14);
            this.labE2.TabIndex = 11;
            this.labE2.Text = "正在进行...";
            this.labE2.Visible = false;
            // 
            // labF2
            // 
            this.labF2.AutoSize = true;
            this.labF2.Enabled = false;
            this.labF2.Location = new System.Drawing.Point(137, 125);
            this.labF2.Name = "labF2";
            this.labF2.Size = new System.Drawing.Size(21, 14);
            this.labF2.TabIndex = 11;
            this.labF2.Text = "2.";
            // 
            // labE1
            // 
            this.labE1.AutoSize = true;
            this.labE1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labE1.ForeColor = System.Drawing.Color.Red;
            this.labE1.Location = new System.Drawing.Point(426, 85);
            this.labE1.Name = "labE1";
            this.labE1.Size = new System.Drawing.Size(84, 14);
            this.labE1.TabIndex = 11;
            this.labE1.Text = "正在进行...";
            this.labE1.Visible = false;
            // 
            // labF1
            // 
            this.labF1.AutoSize = true;
            this.labF1.Enabled = false;
            this.labF1.Location = new System.Drawing.Point(137, 85);
            this.labF1.Name = "labF1";
            this.labF1.Size = new System.Drawing.Size(21, 14);
            this.labF1.TabIndex = 11;
            this.labF1.Text = "1.";
            // 
            // labM4
            // 
            this.labM4.AutoSize = true;
            this.labM4.Enabled = false;
            this.labM4.Location = new System.Drawing.Point(177, 211);
            this.labM4.Name = "labM4";
            this.labM4.Size = new System.Drawing.Size(161, 14);
            this.labM4.TabIndex = 10;
            this.labM4.Text = "写入各单位演练成绩汇总";
            // 
            // labM3
            // 
            this.labM3.AutoSize = true;
            this.labM3.Enabled = false;
            this.labM3.Location = new System.Drawing.Point(177, 168);
            this.labM3.Name = "labM3";
            this.labM3.Size = new System.Drawing.Size(119, 14);
            this.labM3.TabIndex = 10;
            this.labM3.Text = "写入事故处置详情";
            // 
            // labM2
            // 
            this.labM2.AutoSize = true;
            this.labM2.Enabled = false;
            this.labM2.Location = new System.Drawing.Point(177, 125);
            this.labM2.Name = "labM2";
            this.labM2.Size = new System.Drawing.Size(119, 14);
            this.labM2.TabIndex = 10;
            this.labM2.Text = "写入演练基本情况";
            // 
            // labM1
            // 
            this.labM1.AutoSize = true;
            this.labM1.Enabled = false;
            this.labM1.Location = new System.Drawing.Point(177, 85);
            this.labM1.Name = "labM1";
            this.labM1.Size = new System.Drawing.Size(77, 14);
            this.labM1.TabIndex = 10;
            this.labM1.Text = "文档初始化";
            // 
            // labPercent
            // 
            this.labPercent.AutoSize = true;
            this.labPercent.Location = new System.Drawing.Point(622, 43);
            this.labPercent.Name = "labPercent";
            this.labPercent.Size = new System.Drawing.Size(21, 14);
            this.labPercent.TabIndex = 9;
            this.labPercent.Text = "0%";
            this.labPercent.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(69, 34);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(535, 37);
            this.progressBar1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.btn_Start);
            this.groupBox1.Controls.Add(this.btn_Close);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Location = new System.Drawing.Point(12, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(719, 486);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(747, 499);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成演练结果报告";
            this.Load += new System.EventHandler(this.FrmMian_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.Button btn_Close;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textPath;
        private System.Windows.Forms.Button txt_select;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label labE5;
        private System.Windows.Forms.Label labF5;
        private System.Windows.Forms.Label labM5;
        private System.Windows.Forms.Label labE4;
        private System.Windows.Forms.Label labF4;
        private System.Windows.Forms.Label labE3;
        private System.Windows.Forms.Label labF3;
        private System.Windows.Forms.Label labE2;
        private System.Windows.Forms.Label labF2;
        private System.Windows.Forms.Label labE1;
        private System.Windows.Forms.Label labF1;
        private System.Windows.Forms.Label labM4;
        private System.Windows.Forms.Label labM3;
        private System.Windows.Forms.Label labM2;
        private System.Windows.Forms.Label labM1;
        private System.Windows.Forms.Label labPercent;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label labE6;
        private System.Windows.Forms.Label labF6;
        private System.Windows.Forms.Label labM6;
    }
}

