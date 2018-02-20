using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSAutoWord;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Data;

namespace EquReportToWord
{
    class BaseLogic
    {

        private string m_path;
        private string file_name;
        private int lineSpacing = 18;//行间距
        private string m_sDatabaseFilePath = Directory.GetCurrentDirectory() + "/DB_CSMT.mdb";
        private LogicDB m_logicDB;

        CSWord word = new CSWord();
        BaseInfo baseInfo = new BaseInfo();

        public BaseLogic(string path)
        {
            m_path = path;
            m_logicDB = new LogicDB(m_sDatabaseFilePath);
        }


        public string CreateReport()
        {
            try
            {
                //ExpressProgross.ExpressProgressIndex(0);
                
                CreateCover();
                ExpressProgross.SetProgress(10);
                word.InsertText("目录", 16, 1, Word.WdParagraphAlignment.wdAlignParagraphCenter, true);
                word.InsertNewPage();
                ExpressProgross.SetProgress(20);

                ExpressProgross.RefreshCurrent(1);
                ChapterOne();
                
                ExpressProgross.SetProgress(40);
                ExpressProgross.RefreshCurrent(2);
                ChapterTwo();
                
                ExpressProgross.SetProgress(60);
                ExpressProgross.RefreshCurrent(3);
                ChapterThree();
                
                ExpressProgross.SetProgress(70);
                ExpressProgross.RefreshCurrent(4);
                ChapterFour();
                
                ExpressProgross.SetProgress(80);
                ExpressProgross.RefreshCurrent(5);
                ChapterFive();
                
                word.CreateContents();
                //ExpressProgross.ExpressProgressIndex(6);
                ExpressProgross.SetProgress(100);

                if (!Directory.Exists(m_path))
                    Directory.CreateDirectory(m_path);
                file_name = DateTime.Now.ToString("yyyy年M月d日H时m分") + "-演练结果报告";
                word.SaveWordDocument(m_path, file_name);
                ExpressProgross.RefreshLast(5);
                return file_name;
            }
            catch (Exception e)
            {
                return "";
            }
        }

        #region 创建封面
        private void CreateCover()
        {
            string[] title = new string[] { "神华集团煤制油分公司", "应急与培训系统", "演练结果报告" };
            for (int i = 0; i < title.Length; i++ )
                word.InsertText(title[i], 36, 3, Word.WdParagraphAlignment.wdAlignParagraphCenter);
         

            for(int i=0; i<5;i++)
                word.InsertLine();

            string logoPath = Directory.GetCurrentDirectory() + "/ShenhuaLogo.jpg";

            word.InsertPicture(logoPath, 180, 150, Word.WdParagraphAlignment.wdAlignParagraphCenter);

            for (int i = 0; i < 15; i++)
                word.InsertLine();

            string unit = "编制单位：中国神华煤制油化工有限公司";
            word.InsertText(unit, 15, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            string time = DateTime.Now.ToString("文档生成时间：yyyy年M月d日");
            word.InsertText(time, 15, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter);

        }
        #endregion

        #region 第一章
        public void ChapterOne()
        {
            int iChapter = 0;
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING1, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//1.1
            List<string> listText = m_logicDB.GetText("1.1.txt");
            for (int i = 0; i < listText.Count; i++)
            {
                word.InsertText(listText[i], 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            }

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//1.2 
            word.InsertText(baseInfo.GetInfo(iChapter) + m_logicDB.GetRunTime(), 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
          
            word.InsertText(baseInfo.GetInfo(iChapter) + m_logicDB.GetExitTime(), 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);///////////1.3

            string str = "此次参演的部门有";
            List<string> listRoleUser = m_logicDB.GetRoleUser();
            List<string> listRoleName = m_logicDB.GetRoleName(listRoleUser);
            for (int i = 0; i < listRoleName.Count; i++)
            {
                str += listRoleName[i] + "、";
            }
            str = str.Substring(0, str.Length - 1);
            str += "，其中";
            for (int i = 0; i < listRoleName.Count; i++)
            {
                str += listRoleName[i] + m_logicDB.GetUserCount(listRoleUser, listRoleName[i]) + "人、";
            }
            str = str.Substring(0, str.Length - 1);
            str += "。";

            //string str2 = "此次参演的部门有部门1、部门1、部门1、部门1、部门1、部门1、部门1、部门1（导演组除外），其中部门1 18人，。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。";
            word.InsertText("\t" + str, 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);//////1.3正文
            str = "表1-3详细说明参演人员具体参与情况。";
            word.InsertText(str, 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);

            word.InsertLine();
            word.InsertText(baseInfo.GetInfo(iChapter), (float)10.5, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, lineSpacing);
            Word.Table tableInfo = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, listRoleUser.Count/2 + 1, 4, false);
            tableInfo.Cell(1, 1).Range.Text = "角色";
            tableInfo.Cell(1, 1).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdGray25;
            tableInfo.Cell(1, 2).Range.Text = "用户名";
            tableInfo.Cell(1, 2).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdGray25;
            tableInfo.Cell(1, 3).Range.Text = "是否参与";
            tableInfo.Cell(1, 3).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdGray25;
            tableInfo.Cell(1, 4).Range.Text = "参与时限";
            tableInfo.Cell(1, 4).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdGray25;

            List<string> list_IsPart = m_logicDB.GetIsPart(listRoleUser);
            List<string> list_LogInOutTime = m_logicDB.GetLogInOutTime(listRoleUser);
            for (int i = 0; i < listRoleUser.Count; i+=2)
            {
                tableInfo.Cell(i / 2 + 2, 1).Range.Text = listRoleUser[i + 1];
                tableInfo.Cell(i / 2 + 2, 2).Range.Text = listRoleUser[i];
                tableInfo.Cell(i / 2 + 2, 3).Range.Text = list_IsPart[i / 2];
                tableInfo.Cell(i / 2 + 2, 4).Range.Text = list_LogInOutTime[i / 2];
            }


        }
        #endregion

        #region 第二章
        public void ChapterTwo()
        {
            int iChapter = 1;
            word.InsertNewPage();
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING1, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//2.1
            DataTable dt_ProcessManager = m_logicDB.GetProcessManager();
            for (int i = 0; i < dt_ProcessManager.Rows.Count; i++ )
                word.InsertText(dt_ProcessManager.Rows[i]["Process"].ToString(), 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//2.2
            List<string> listText = m_logicDB.GetText("2.2.txt");
            for (int i = 0; i < listText.Count; i++)
            {
                word.InsertText(listText[i], 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            }

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//2.3
            DataTable dt_CtrlList = m_logicDB.GetCtrlList();
            for(int i=0; i<dt_CtrlList.Rows.Count; i++)
                word.InsertText(dt_CtrlList.Rows[i]["Time"].ToString() + " " + dt_CtrlList.Rows[i]["Process"].ToString(), 
                    12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//2.4
            DataTable dt_TaskRate = m_logicDB.GetDataTableTaskRate();
            DataTable dt_TaskConfig = m_logicDB.GetDataTableTaskConfig();
            
            List<string> listTableHeader = new List<string>();
            listTableHeader.Add("任务名");
            listTableHeader.Add("参与部门");
            listTableHeader.Add("各评审分数");
            listTableHeader.Add("平均分");
            Word.Table tableInfo = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, dt_TaskConfig.Rows.Count + 1, listTableHeader.Count, false);
            SetTableHeader(tableInfo, listTableHeader);
            for (int i = 0; i < dt_TaskConfig.Rows.Count; i++)
            {
                int index = 1;
                tableInfo.Cell(i + 2, index++).Range.Text = dt_TaskConfig.Rows[i]["TaskName"].ToString();
                tableInfo.Cell(i + 2, index++).Range.Text = dt_TaskConfig.Rows[i]["RolePart"].ToString();
                DataView dv = new DataView(dt_TaskRate);
                dv.RowFilter = "TaskID = '" + dt_TaskConfig.Rows[i]["TaskID"].ToString() + "'";
                DataTable dt = dv.ToTable();
                if (dt.Rows.Count != 0)
                {
                    tableInfo.Cell(i + 2, index++).Range.Text = GetTaskRate(dt.Rows[0]["TaskRate"].ToString());
                    tableInfo.Cell(i + 2, index++).Range.Text = ((int)m_logicDB.GetAvgTaskRate(dt_TaskConfig.Rows[i]["TaskID"].ToString())).ToString();
                }
            }
        }

        private string GetTaskRate(string sTaskRates)
        {
            string[] str = sTaskRates.Split(',');
            string sRate = "";
            for(int i=0; i<str.Length; i++)
            {
                sRate += "评审" + (i + 1).ToString() + ":  " + str[i] + "分\n";
            }
            sRate = sRate.Substring(0, sRate.Length - 1);
            return sRate;
        }
        #endregion

        #region 第三章
        public void ChapterThree()
        {
            int iChapter = 2;
            word.InsertNewPage();
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING1, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//3.1
            DataTable dt_UserResult = m_logicDB.GetUsersResult();
            List<string> listResutOfAnswer = m_logicDB.GetResultOfAnswer();
            int index3_1 = 0;
            string str3_1 = "在本次演练中，答题参与人员是" + listResutOfAnswer[index3_1++] + "人，答题及格人数是" + listResutOfAnswer[index3_1++] +
                "人，不及格人数" + listResutOfAnswer[index3_1++] + "人，";
            if (index3_1 < listResutOfAnswer.Count)
                str3_1 += "其中";
            while (index3_1 < listResutOfAnswer.Count)
            {
                str3_1 += listResutOfAnswer[index3_1++] + "，";
            }
            str3_1 = str3_1.Substring(0, str3_1.Length - 1);
            str3_1 += "。表3-1详细说明此次演练的详细情况。";
            word.InsertText("\t" + str3_1, (float)12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            word.InsertLine();
            word.InsertText(baseInfo.GetInfo(iChapter), (float)10.5, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, lineSpacing);//表格
            
             List<string> listTableHeader = new List<string>();
            listTableHeader.Add("部门");
            listTableHeader.Add("职员");
            listTableHeader.Add("单选对比例");
            listTableHeader.Add("多选对比例");
            listTableHeader.Add("问答对比例");
            listTableHeader.Add("用户成绩");
            Word.Table tableInfo = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, dt_UserResult.Rows.Count + 1, listTableHeader.Count, false);
            SetTableHeader(tableInfo, listTableHeader);
            for (int i = 0; i < dt_UserResult.Rows.Count; i++)
            {
                tableInfo.Cell(i + 2, 1).Range.Text = dt_UserResult.Rows[i]["RoleName"].ToString();
                tableInfo.Cell(i + 2, 2).Range.Text = dt_UserResult.Rows[i]["UserName"].ToString();
                tableInfo.Cell(i + 2, 3).Range.Text = dt_UserResult.Rows[i]["SingleCount"].ToString();
                tableInfo.Cell(i + 2, 4).Range.Text = dt_UserResult.Rows[i]["MutiCount"].ToString();
                tableInfo.Cell(i + 2, 5).Range.Text = dt_UserResult.Rows[i]["TextCount"].ToString();
                tableInfo.Cell(i + 2, 6).Range.Text = dt_UserResult.Rows[i]["AllScore"].ToString();
            }
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//3.2
            int iTotalCount = 0;//任务的总数
            int iTotalFinishCount = 0; //完成任务的数量
            List<string> listRoleName;//参加的部门
            int[] iFinishCount;//各部门完成任务个数
            int[] iCount;//各部门总任务个数
            int[] iScoreRole;// 各组任务的总得分
            int[] iCountUserFinish;//得分贡献人数
            List<string> listResultOfTask = m_logicDB.GetResultOfTask(out iTotalCount,out iTotalFinishCount, out listRoleName, out iFinishCount, out iCount, out iScoreRole, out iCountUserFinish);
            int index = 0;
            string str = "在本次演练中，总共参加的部门有" + listResultOfTask[index++] + "个，" +
                "其中完成预定任务的部门有" + listResultOfTask[index++] + "个，任务的总数为" + iTotalCount +
                "个，完成的个数" + iTotalFinishCount + "个，任务的总完成率是" + listResultOfTask[index] + "。";
            word.InsertText("\t" + str, 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            word.InsertLine();

            word.InsertText(baseInfo.GetInfo(iChapter), (float)10.5, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, lineSpacing);//表格
            List<string> listTableHeader2 = new List<string>();
            listTableHeader2.Add("部门");
            listTableHeader2.Add("人数");
            listTableHeader2.Add("完成任务比例");
            listTableHeader2.Add("任务得分");
            listTableHeader2.Add("得分贡献人数");
            Word.Table tableInfo2 = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, listRoleName.Count + 1, listTableHeader2.Count, false);
            SetTableHeader(tableInfo2, listTableHeader2);
            for (int i = 0; i < listRoleName.Count; i++)
            {
                tableInfo2.Cell(i + 2, 1).Range.Text = listRoleName[i];//部门
                tableInfo2.Cell(i + 2, 2).Range.Text = m_logicDB.GetUserCount(listRoleName[i]).ToString();//各部门人数
                tableInfo2.Cell(i + 2, 3).Range.Text = iFinishCount[i] + "/" + iCount[i];//完成任务比例
                tableInfo2.Cell(i + 2, 4).Range.Text = iScoreRole[i].ToString();//任务得分
                tableInfo2.Cell(i + 2, 5).Range.Text = iCountUserFinish[i].ToString();//得分贡献人数
            }
        }
        #endregion

        #region 第四章
        public void ChapterFour()
        {
            int iChapter = 3;
            word.InsertNewPage();
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING1, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//4.1
            List<string> listRoleResult = new List<string>();
            word.InsertText("\t" + m_logicDB.GetResultDetail(out listRoleResult), (float)12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            int iIndex4_1 = 1;

            for (int i = 0; i < listRoleResult.Count; i++)
            {

                word.InsertText(m_logicDB.GetTitleOf4_1(iIndex4_1++, listRoleResult[i]), CSWord.WORD_HEADING3);
                List<string> listExitUserResultDetail;
                word.InsertText("\t" + m_logicDB.GetDetail(listRoleResult[i], out listExitUserResultDetail), (float)12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
                for (int j = 0; j < listExitUserResultDetail.Count; j++)
                {
                    word.InsertText(m_logicDB.Get4_1TableName(listExitUserResultDetail[j]), (float)10.5, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, lineSpacing);
                    DataTable dt_UserDetail = m_logicDB.GetUserTable(listExitUserResultDetail[j]);
                    List<string> listTableHeader4_1 = new List<string>();
                    listTableHeader4_1.Add("题型");
                    listTableHeader4_1.Add("题目内容");
                    listTableHeader4_1.Add("正确答案");
                    listTableHeader4_1.Add("用户答案");
                    listTableHeader4_1.Add("耗时");
                    listTableHeader4_1.Add("题目分数");
                    listTableHeader4_1.Add("得分");

                    Word.Table tableInfo4_1 = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, dt_UserDetail.Rows.Count + 1, listTableHeader4_1.Count, false);
                    SetTableHeader(tableInfo4_1, listTableHeader4_1);
                    for (int k = 0; k < dt_UserDetail.Rows.Count; k++)
                    {
                        int index = 1;
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["QuesSort"].ToString();
                        if (dt_UserDetail.Rows[k]["QuesContext"].ToString().Length > 10 )
                            tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["QuesContext"].ToString().Substring(0, 10) + "...";
                        else
                            tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["QuesContext"].ToString();
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["CorrAns"].ToString();
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["UserAns"].ToString();
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["SpendTime"].ToString();
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["Score"].ToString();
                        tableInfo4_1.Cell(k + 2, index++).Range.Text = dt_UserDetail.Rows[k]["UserScore"].ToString();
                       
                    }
                }
            }
            ///////////////////////4.2
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//4.2
            List<string> listRoleName;//参加的部门
            int[] iFinishCount;//各部门完成任务个数
            int[] iCount;//各部门总任务个数
            int[] iScoreRole;// 各组任务的总得分
            int[] iCountUserFinish;//得分贡献人数
            int iTotalCount = 0;//总任务数
            int iTotalFinishCount = 0;//总的完成的任务数
            m_logicDB.GetResultOfTask(out iTotalCount,out iTotalFinishCount, out listRoleName, out iFinishCount, out iCount, out iScoreRole, out iCountUserFinish);
           
            for (int i = 0; i < iCount.Length; i++)
            {
                iTotalCount += iCount[i];
                iTotalFinishCount += iFinishCount[i];
            }
            string str = "此次演练总共有任务" + iTotalCount + "个，" + "完成的个数为" + iTotalFinishCount + "个。";
            if (iTotalCount > 0)
                str += "其中";
            for (int i = 0; i < iCount.Length; i++)
            {
                if (iCount[i] > 0)
                    str += listRoleName[i] + "有" + iCount[i] + "个任务点，完成了" + iFinishCount[i] + "个任务点；";
            }
            if (iTotalCount > 0)
            {
                str = str.Substring(0, str.Length - 1);
                str += "。";
            }
            word.InsertText("\t" + str, (float)12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            int iIndex4_2 = 0;
            for (int i = 0; i < listRoleName.Count; i++)
            {
                if (iCount[i] > 0)
                {
                    iIndex4_2++;
                    word.InsertText(m_logicDB.GetTitleOf4_2(iIndex4_2, listRoleName[i]), CSWord.WORD_HEADING3);
                    string str4_2 = "";
                    str4_2 = listRoleName[i] + "有职工" + m_logicDB.GetCountUserDoTask(listRoleName[i]) + "人参与，，总共有任务" +
                        iCount[i] + "个，完成任务" + iFinishCount[i] + "个，得分贡献人数为" + iCountUserFinish[i] + "人。表4_2_" +
                        iIndex4_2 + "说明部门的任务完成详情。";
                    word.InsertText("\t" + str4_2, (float)12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
                    string strTableTitle = "表4_2_" + iIndex4_2 + " " + listRoleName[i] + "的任务完成详情";
                    word.InsertText(strTableTitle, (float)10.5, 0, Word.WdParagraphAlignment.wdAlignParagraphCenter, lineSpacing);

                    List<string> listTableHeader4_2 = new List<string>();
                    listTableHeader4_2.Add("任务类型");
                    listTableHeader4_2.Add("任务内容");
                    listTableHeader4_2.Add("时间限制");
                    listTableHeader4_2.Add("分数");
                    listTableHeader4_2.Add("执行者");
                    listTableHeader4_2.Add("耗时");
                    listTableHeader4_2.Add("得分");
                    DataTable dt_UserDoTask = m_logicDB.GetDataTableUserDoTask(listRoleName[i]);
                    Word.Table tableInfo4_2 = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, dt_UserDoTask.Rows.Count + 1, listTableHeader4_2.Count, false);
                    SetTableHeader(tableInfo4_2, listTableHeader4_2);
                    for (int j = 0; j < dt_UserDoTask.Rows.Count; j++)
                    {
                        DataTable dt_TaksConfig = m_logicDB.GetTaskConfig(dt_UserDoTask.Rows[j]["TaskID"].ToString());
                        int index = 1;
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_TaksConfig.Rows[0]["Type"].ToString();
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_TaksConfig.Rows[0]["Content"].ToString();
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_TaksConfig.Rows[0]["TimeLimit"].ToString() + "分钟";
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_TaksConfig.Rows[0]["Score"].ToString();
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_UserDoTask.Rows[j]["Completer"].ToString();
                        tableInfo4_2.Cell(j + 2, index++).Range.Text = int.Parse(dt_UserDoTask.Rows[j]["SpendTime"].ToString()) / 60 + "分钟" +
                            int.Parse(dt_UserDoTask.Rows[j]["SpendTime"].ToString()) % 60 + "秒";
                        if (dt_UserDoTask.Rows[j]["IsComplete"].ToString() == "True")
                            tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_UserDoTask.Rows[j]["Score"].ToString();
                        else
                            tableInfo4_2.Cell(j + 2, index++).Range.Text = dt_UserDoTask.Rows[j]["Score"].ToString() + "(" +
                                dt_UserDoTask.Rows[j]["Remark"].ToString() + ")";
                    }
                }
            }

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//4.3
           
            List<string> listTableHeader4_3 = new List<string>();
            listTableHeader4_3.Add("用户名");
            listTableHeader4_3.Add("角色名");
            listTableHeader4_3.Add("总分");
            DataTable dt_UserResult = m_logicDB.GetUsersResult();
            Word.Table tableInfo4_3 = word.InsertTable(Word.WdParagraphAlignment.wdAlignParagraphLeft, dt_UserResult.Rows.Count + 1, listTableHeader4_3.Count, false);
            SetTableHeader(tableInfo4_3, listTableHeader4_3);
            for (int i = 0; i < dt_UserResult.Rows.Count; i++)
            {
                int index = 1;
                tableInfo4_3.Cell(i + 2, index++).Range.Text = dt_UserResult.Rows[i]["UserName"].ToString();
                tableInfo4_3.Cell(i + 2, index++).Range.Text = dt_UserResult.Rows[i]["RoleName"].ToString();
                tableInfo4_3.Cell(i + 2, index++).Range.Text = m_logicDB.GetTotalScore(dt_UserResult.Rows[i]["UserName"].ToString(),
                    int.Parse(dt_UserResult.Rows[i]["AllScore"].ToString())).ToString();
            }

        }
        #endregion

        #region 第五章
        public void ChapterFive()
        {
            int iChapter = 4;
            word.InsertNewPage();
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING1, Word.WdParagraphAlignment.wdAlignParagraphCenter);
            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//5.1
            List<string> listText5_1 = m_logicDB.GetText("5.1.txt");
            for (int i = 0; i < listText5_1.Count; i++)
            {
                word.InsertText(listText5_1[i], 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            }

            word.InsertText(baseInfo.GetInfo(iChapter), CSWord.WORD_HEADING2);//5.2
            List<string> listText5_2 = m_logicDB.GetText("5.2.txt");
            for (int i = 0; i < listText5_2.Count; i++)
            {
                word.InsertText(listText5_2[i], 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft, lineSpacing);
            }
        }
        #endregion

        public void SetTableHeader(Word.Table tableInfo, List<string> listTableHeader)
        {
            for (int i = 0; i < listTableHeader.Count; i++)
            {
                tableInfo.Cell(1, i+1).Range.Text = listTableHeader[i];
                tableInfo.Cell(1, i+1).Shading.BackgroundPatternColorIndex = Word.WdColorIndex.wdGray25;
            }
        }

    }
}
