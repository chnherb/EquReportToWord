using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EquReportToWord
{
    class BaseInfo
    {
        private const int m_ChapCount = 5;//章节数
        private static int[] m_index;//每个章节string字符串的下标
        private static List<string>[] m_list;

        public BaseInfo()
        {
            m_index = new int[m_ChapCount];
            for (int i = 0; i < m_ChapCount; i++)
            {
                m_index[i] = -1;
            }
            m_list = new List<string>[m_ChapCount];
            for (int i = 0; i < m_ChapCount; i++)
            {
                m_list[i] = new List<string>();
            }
            m_list[0].Add("第一章 演练基本情况");
            m_list[0].Add("1.1 演练预案说明");

            m_list[0].Add("1.2 演练的时间");
            m_list[0].Add("演练开始时间:");
            m_list[0].Add("演练结束时间:");

            m_list[0].Add("1.3 参演人员参与情况");
            m_list[0].Add("表1-1详细说明参演人员具体参与情况。");
            m_list[0].Add("演练开始时间:");
            m_list[0].Add("演练开始时间:");
            m_list[0].Add("演练开始时间:");
            m_list[0].Add("表1-3 参演人员参与情况表");


            m_list[1].Add("第二章  事故处置详情");
            m_list[1].Add("2.1 预案基本流程");

            m_list[1].Add("2.2 预案处置基本情况");
            m_list[1].Add("2.3 事故全程处置记录");
            m_list[1].Add("2.4 任务评审情况");

            m_list[2].Add("第三章 各单位演练成绩汇总");
            m_list[2].Add("3.1 演练答题成绩总体情况");
            m_list[2].Add("表3-1 演练答题成绩汇总表");

            m_list[2].Add("3.2 演练任务完成总体情况");
            m_list[2].Add("表3-2 演练任务完成情况");

            m_list[3].Add("第四章 各单位成绩明细");
            m_list[3].Add("4.1 用户答题得分明细");
            //m_list[3].Add("4.1.1 部门一答题情况");
            //m_list[3].Add("4.1.2 部门二答题情况");
            m_list[3].Add("4.2 用户任务完成明细");
            m_list[3].Add("4.3 用户总分情况");
            m_list[4].Add("第五章 演练总结");
            m_list[4].Add("5.1 领导总结");
            m_list[4].Add("5.2 各个单位负责人感言");
        }

        public string GetInfo(int i)
        {
            if (i < 0 || i >= m_ChapCount)
                return "";
            m_index[i]++;
            if (m_index[i] >= m_list[i].Count)
                return "";
            return m_list[i][m_index[i]];
        }


    }
}
