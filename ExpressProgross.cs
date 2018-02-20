using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EquReportToWord
{
    class ExpressProgross
    {

        public static ProgressBar m_proBar;
        public static Label m_labCur;
        public static List<Label> labF = new List<Label>();
        public static List<Label> labM = new List<Label>();
        public static List<Label> labE = new List<Label>();

        public static void InitAll()
        {
            m_proBar.Value = m_proBar.Minimum;
            m_labCur.Visible = true;
            m_labCur.Text = "0%";
        }

       
     

        public static void SetProgress(int pro)
        {
            if (pro == 0)
                InitAll();
            m_proBar.Value = pro + m_proBar.Minimum;
            m_labCur.Text = pro.ToString() + "%";
        }

        public static void RefreshCurrent(int i)
        {
            if (i > 0)
                RefreshLast(i - 1);
            labF[i].Text = "--->";
            labF[i].Enabled = true;
            labM[i].Enabled = true;
            labE[i].Visible = true;
        }

        public static void RefreshLast(int i)
        {
            labF[i].Text = (i + 1) + ".";
            labF[i].Enabled = false;
            labM[i].Enabled = false;
            labE[i].Text = "完成";
        }

    }
}
