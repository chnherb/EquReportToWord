using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSAutoWord;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace EquReportToWord
{
    public partial class FrmMain : Form
    {

        private string m_sWordPath = Application.StartupPath + "\\Word\\";
        CSWord word = new CSWord();
        private string m_sFileName = ""; 


        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMian_Load(object sender, EventArgs e)
        {
            ExpressProgross.labF.Add(labF1);
            ExpressProgross.labF.Add(labF2);
            ExpressProgross.labF.Add(labF3);
            ExpressProgross.labF.Add(labF4);
            ExpressProgross.labF.Add(labF5);
            ExpressProgross.labF.Add(labF6);
            ExpressProgross.labM.Add(labM1);
            ExpressProgross.labM.Add(labM2);
            ExpressProgross.labM.Add(labM3);
            ExpressProgross.labM.Add(labM4);
            ExpressProgross.labM.Add(labM5);
            ExpressProgross.labM.Add(labM6);
            ExpressProgross.labE.Add(labE1);
            ExpressProgross.labE.Add(labE2);
            ExpressProgross.labE.Add(labE3);
            ExpressProgross.labE.Add(labE4);
            ExpressProgross.labE.Add(labE5);
            ExpressProgross.labE.Add(labE6);
            ExpressProgross.m_labCur = labPercent;
            ExpressProgross.m_proBar = progressBar1;
            textPath.Text = m_sWordPath;
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            if (btn_Start.Text.Equals("开始生成"))
            {
                btn_Start.Enabled = false;
                ExpressProgross.RefreshCurrent(0);
                ExpressProgross.SetProgress(0);
                BaseLogic baseLogic = new BaseLogic(m_sWordPath);
                m_sFileName = baseLogic.CreateReport();
            }
            if (btn_Start.Text.Equals("开始生成") && m_sFileName != "")
            {
                    MessageBox.Show("文档生成成功！");
                    btn_Start.Text = "查看文档";
                    btn_Start.Enabled = true;
            }
            else
            {
                if (!word.DisplayWordFile((object)(m_sWordPath + "/" + m_sFileName)))
                    MessageBox.Show("文档不存在或被占用");
            }
        }

        private void btn_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txt_select_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog G_FolderBrowserDialog = null;//定义浏览文件夹字段
            G_FolderBrowserDialog = //创建浏览文件夹对象
               new FolderBrowserDialog();
            DialogResult P_DialogResult =//浏览文件夹
                G_FolderBrowserDialog.ShowDialog();

            if (P_DialogResult == DialogResult.OK)//确认已经选择文件夹
            {

                if (textPath.Text != G_FolderBrowserDialog.SelectedPath)
                {
                    textPath.Text = //显示选择路径
                        G_FolderBrowserDialog.SelectedPath;
                    //m_sWordPath = string.Format(//计算文件保存路径
                    //   @"{0}", G_FolderBrowserDialog.SelectedPath + "\\");
                    m_sWordPath = G_FolderBrowserDialog.SelectedPath + "\\";
                    textPath.Text = m_sWordPath;
                }
                Initial();
            }
        }

        private void Initial()
        {
            for (int i = 0; i < ExpressProgross.labE.Count; i++)
            {
                ExpressProgross.labE[i].Visible = false;
                ExpressProgross.labE[i].Text = "正在进行...";
            }
            btn_Start.Text = "开始生成";
            m_sFileName = "";
            ExpressProgross.m_proBar.Value = ExpressProgross.m_proBar.Minimum;
            ExpressProgross.m_labCur.Visible = false;
            ExpressProgross.m_labCur.Text = "0%";
        }

       
    }
}
