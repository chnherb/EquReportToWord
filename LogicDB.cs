using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SetQues.DB;

namespace EquReportToWord
{
    class LogicDB
    {
        private string m_sFilePath;
        private BaseDBData m_db;
        private DataTable m_dt_TaskConfig;
        private DataTable m_dt_DoTask;
        private DataTable m_dt_RoleType;
        private DataTable m_dt_UsersResult;
        private DataTable m_dt_UserResultDetai;
        private DataTable m_dt_TaskRate;
        public LogicDB(string sFilePath)
        {
            m_sFilePath = sFilePath;
            m_db = new BaseDBData(sFilePath);

            string sql = "select * from tblTaskConfig";
            m_dt_TaskConfig = m_db.GetTableInfo(sql);

            sql = "select * from tblDoTask";
            m_dt_DoTask = m_db.GetTableInfo(sql);

            sql = "select * from tblRoleType";
            m_dt_RoleType = m_db.GetTableInfo(sql);

            sql = "select * from tblUsersResult";
            m_dt_UsersResult = m_db.GetTableInfo(sql);

            sql = "select * from tblUserResultDetail";
            m_dt_UserResultDetai = m_db.GetTableInfo(sql);

            sql = "select * from tblTaskRate";
            m_dt_TaskRate = m_db.GetTableInfo(sql);
            
        }

        public DataTable GetDataTableTaskRate()
        {
            return m_dt_TaskRate;
        }
        public DataTable GetDataTableTaskConfig()
        {
            return m_dt_TaskConfig;
        }

        //返回用户名、角色名，用list保存在一起
        public List<string> GetRoleUser()
        {
            List<string> list_RoleUser = new List<string>();

            //string sql = "select RoleID,RoleName,RolePoint from tblRoleType ";
            //DataTable dt_all = m_db.GetTableInfo(sql);

            DataView dv_all = new DataView(m_dt_RoleType);
            dv_all.RowFilter = "RolePoint = 'User'";
            DataTable dt_User = dv_all.ToTable();

            for (int i = 0; i < dt_User.Rows.Count; i++)
            {
                if (dt_User.Rows[i]["RoleName"].ToString().Substring(0, 1) != "#" &&
                    dt_User.Rows[i]["RoleName"].ToString().Substring(dt_User.Rows[i]["RoleName"].ToString().Length - 1, 1) != "$")
                {

                    list_RoleUser.Add(dt_User.Rows[i]["RoleName"].ToString());//添加用户名

                    DataView dv_Role = new DataView(m_dt_RoleType);
                    dv_Role.RowFilter = "RoleID = '" + dt_User.Rows[i]["RoleID"].ToString().Substring(0, 7) + "'";
                    DataTable dt_Role = dv_Role.ToTable();
                    list_RoleUser.Add(dt_Role.Rows[0]["RoleName"].ToString());//添加角色名

                    //删除导演组及其角色
                    if (list_RoleUser[list_RoleUser.Count - 1] == "导演组")
                    {
                        list_RoleUser.RemoveAt(list_RoleUser.Count - 1);
                        list_RoleUser.RemoveAt(list_RoleUser.Count - 1);
                    }
                }
            }
            return list_RoleUser;
        }

        public List<string> GetRoleName(List<string> list_RoleUser)
        {
            List<string> list_RoleName = new List<string>();
            for (int i = 1; i < list_RoleUser.Count; i+=2)
            {
                bool flag = false;
                for (int j = 0; j < list_RoleName.Count; j++)
                {
                    if (list_RoleUser[i] == list_RoleName[j])
                    {
                        flag = true;//表示找到同名组
                        break;
                    }
                }
                if (false == flag)
                    list_RoleName.Add(list_RoleUser[i]);
            }
            return list_RoleName;
        }

        //获取不同角色的用户人数，形参为GetRoleUser()的返回值和角色名，避免多次读数据库，提高效率
        public int GetUserCount(List<string> list_RoleUser,string Role)
        {
            int iUserCount = 0;
            for (int i = 1; i < list_RoleUser.Count; i+=2)
            {
                if (Role == list_RoleUser[i])
                    iUserCount++;
            }
            return iUserCount;
        }

        //判断用户是否参与 根据tblDoTask查询是否有记录而判断
        //public List<string> GetIsPart(List<string> list_RoleUser)
        //{
        //    List<string> list_IsPart = new List<string>();
        //    string sql = "select TaskID,Completer from tblDoTask";
        //    DataTable dt = m_db.GetTableInfo(sql);
        //    for (int i = 0; i < list_RoleUser.Count; i += 2)
        //    {
        //        DataView dv = new DataView(dt);
        //        dv.RowFilter = "Completer = '" + list_RoleUser[i] + "'";
        //        DataTable dt_UserPart = dv.ToTable();
        //        if (dt_UserPart.Rows.Count != 0)
        //            list_IsPart.Add("是");
        //        else
        //            list_IsPart.Add("否");
        //    }
        //    return list_IsPart;
        //}

        //判断用户是否参与 根据tblDoTask查询是否有记录而判断 根据tblUsersResult中的UserState确定

        public List<string> GetIsPart(List<string> list_RoleUser)
        {
            List<string> list_IsPart = new List<string>();
            //string sql = "select UserName,RoleName,UserState from tblUsersResult";
            //DataTable dt = m_db.GetTableInfo(sql);
            for (int i = 0; i < list_RoleUser.Count; i += 2)
            {
                DataView dv = new DataView(m_dt_UsersResult);
                dv.RowFilter = "UserName = '" + list_RoleUser[i] + "'";
                DataTable dt_UserPart = dv.ToTable();
                if (dt_UserPart.Rows[0]["UserState"].ToString() != "未登录")
                    list_IsPart.Add("是");
                else
                    list_IsPart.Add("否");
            }
            return list_IsPart;
        }

        public List<string> GetLogInOutTime(List<string> list_RoleUser)
        {
            List<string> list_LogInOutTime = new List<string>();
            string sql = "select * from tblUsersResult";
            DataTable dt = m_db.GetTableInfo(sql);
            for (int i = 0; i < list_RoleUser.Count; i += 2)
            {
                DataView dv = new DataView(dt);
                dv.RowFilter = "UserName = '" + list_RoleUser[i] + "'";
                DataTable dt_UserPart = dv.ToTable();
                list_LogInOutTime.Add(dt_UserPart.Rows[0]["LogInTime"] + " - " + dt_UserPart.Rows[0]["LogOutTime"]);
            }
            return list_LogInOutTime;
        }

        //第二章 返回基本流程
        public DataTable GetProcessManager()
        {
            string sql = "select * from tblProcessManager";
            DataTable dt = m_db.GetTableInfo(sql);
            return dt;
        }
        public DataTable GetCtrlList()
        {
            string sql = "select * from tblCtrlList";
            DataTable dt = m_db.GetTableInfo(sql);
            return dt;
        }


        //第三章 返回用户答题情况字符串  去除带# $
        public DataTable GetUsersResult()
        {
            //string sql = "select * from tblUsersResult where RoleName != '导演组' and UserName not like '#%' and UserName not like '%$'";//出错？？

            //string sql = "select * from tblUsersResult where UserName not like '#%' and UserName not like '%$'";
            //DataTable dt_UserQues = m_db.GetTableInfo(sql);

            DataView dv = new DataView(m_dt_UsersResult);
            dv.RowFilter = "UserName not like '#%' and UserName not like '%$'";
            DataTable dt_UserQues = dv.ToTable();
            return dt_UserQues;
        }

        public List<string> GetRoleOfResult(DataTable dt_UserQues)
        {
            List<string> listRoleOfResult = new List<string>();
            for (int i = 0; i < dt_UserQues.Rows.Count; i++)
            {
                if(!listRoleOfResult.Contains(dt_UserQues.Rows[i]["RoleName"].ToString()))
                {
                    listRoleOfResult.Add(dt_UserQues.Rows[i]["RoleName"].ToString());
                }
            }
            return listRoleOfResult;
        }

        //返回3.1 演练答题成绩总体情况
        public List<string> GetResultOfAnswer()
        {
            List<string> listResutOfAnswer = new List<string>();
            DataTable dt_UserQues = GetUsersResult();
            
            int[] iCount = new int[dt_UserQues.Rows.Count];//总的答题数
            int[] iCorrCount = new int[dt_UserQues.Rows.Count];//完成的答题数
            int iPassCount = 0;//及格人数
            List<string> listRoleName = GetRoleOfResult(dt_UserQues);//各个分组
            int[] iNoPassCount = new int[listRoleName.Count];//各组不及格人数
            
            for (int i = 0; i < dt_UserQues.Rows.Count; i++)
            {
                if (IsPass(dt_UserQues.Rows[i]["UserName"].ToString()))//该用户是否及格
                    iPassCount++;//及格人数++
                else
                {
                    int index = listRoleName.IndexOf(dt_UserQues.Rows[i]["RoleName"].ToString());
                    iNoPassCount[index]++;
                }
            }
            listResutOfAnswer.Add(dt_UserQues.Rows.Count.ToString());//参与人员总共人数
            listResutOfAnswer.Add(iPassCount.ToString());//添加及格人数
            listResutOfAnswer.Add((dt_UserQues.Rows.Count - iPassCount).ToString());//添加不及格人数
            for (int i = 0; i < listRoleName.Count; i++)
            {
                if (iNoPassCount[i] != 0)
                    listResutOfAnswer.Add(listRoleName[i] + iNoPassCount[i] + "人");
            }
            return listResutOfAnswer;
        }

        //判断该用户是否及格，从tblUsersResult里面总的答题数量来判断，取60%
        public bool IsPass(string sUserName)
        {
            DataView dv = new DataView(GetUsersResult());
            dv.RowFilter = "UserName = '" + sUserName + "'";
            DataTable dt_User = dv.ToTable();

            string[] sSingleCount = dt_User.Rows[0]["SingleCount"].ToString().Split('/');
            string[] sMutiCount = dt_User.Rows[0]["MutiCount"].ToString().Split('/');
            string[] sTextCount = dt_User.Rows[0]["TextCount"].ToString().Split('/');
            int iCorrCount = int.Parse(sSingleCount[0]) + int.Parse(sMutiCount[0]) + int.Parse(sTextCount[0]);
            int iCount = int.Parse(sSingleCount[1]) + int.Parse(sMutiCount[1]) + int.Parse(sTextCount[1]);
            if ((float)iCorrCount * 5 / 3 < iCount)//没有及格
            {
                return false;
            }
            return true;
        }

        //参加的部门，各部门完成任务个数、各部门总任务个数、 各组任务的总得分、得分贡献人数
        public List<string> GetResultOfTask(out int iTotalCount,out int iTotalFinishCount,out List<string> listRoleName, out int[] iFinishCount, out int[] iCount, out int[] iScoreRole, out int[] iCountUserFinish)
        {
            List<string> listResultOfTask = new List<string>();

            listRoleName = new List<string>();//参加的部门
            for (int i = 0; i < m_dt_TaskConfig.Rows.Count; i++)
            {
                string[] str = m_dt_TaskConfig.Rows[i]["RolePart"].ToString().Split(',');
                for (int j = 0; j < str.Length; j++)
                {
                    if (!listRoleName.Contains(str[j]) && str[j] != "导演组")
                        listRoleName.Add(str[j]);
                }
            }
            listResultOfTask.Add(listRoleName.Count.ToString());//添加总共参加部门的个数

            iFinishCount = new int[int.Parse(listRoleName.Count.ToString())];//各部门完成任务个数
            iCount = new int[int.Parse(listRoleName.Count.ToString())];//各部门总任务个数
            iScoreRole = new int[int.Parse(listRoleName.Count.ToString())];//各组任务的总得分
            iCountUserFinish = new int[int.Parse(listRoleName.Count.ToString())];//得分贡献人数
            List<string>[] listUserFinish = new List<string>[listRoleName.Count];//得分贡献人数，任务完成，则将该用户加入
            List<string> listTaskFinish = new List<string>();//若该任务完成，则添加
            //int iFinishTaskCount = 0;//完成任务的个数
            for (int j = 0; j < listUserFinish.Length; j++)
            {
                listUserFinish[j] = new List<string>();
            }
            for (int i = 0; i < m_dt_TaskConfig.Rows.Count; i++)
            {
                string[] str = m_dt_TaskConfig.Rows[i]["RolePart"].ToString().Split(',');
                for (int j = 0; j < str.Length; j++)
                {
                    if (str[j] == "导演组")
                        break;
                    int index = listRoleName.IndexOf(str[j]);//找到list中该组的下标
                    iCount[index]++;
                    DataView dv = new DataView(m_dt_DoTask);
                    dv.RowFilter = "Completer in " + GetUserNames(str[j]) + " and IsComplete = true and TaskID = '" + 
                        m_dt_TaskConfig.Rows[i]["TaskID"].ToString() + "'";
                    DataTable dt_UserPart = dv.ToTable();
                    if (dt_UserPart.Rows.Count != 0)    //该任务完成
                    {
                        iFinishCount[index]++;
                        if (!listTaskFinish.Contains(m_dt_TaskConfig.Rows[i]["TaskID"].ToString()))
                            listTaskFinish.Add(m_dt_TaskConfig.Rows[i]["TaskID"].ToString());//若任务完成，则加入完成的list中
                        iScoreRole[index] += int.Parse(dt_UserPart.Rows[0]["Score"].ToString());
                    }
                    for (int k = 0; k < dt_UserPart.Rows.Count; k++)
                    {
                        if (!listUserFinish[index].Contains(dt_UserPart.Rows[k]["Completer"].ToString()))
                            listUserFinish[index].Add(dt_UserPart.Rows[k]["Completer"].ToString());//将完成了任务的用户加入
                    }
                }
                
            }
            for (int i = 0; i < listUserFinish.Length; i++)
            {
                iCountUserFinish[i] = listUserFinish[i].Count;
            }
            int iFinishRoleCount = 0;//完成任务的部门数
           
            for (int i = 0; i < iCount.Length; i++)
            {
                //iFinishTaskCount += iFinishCount[i];
                if (iFinishCount[i] == iCount[i])
                    iFinishRoleCount++;
            }
            listResultOfTask.Add(iFinishRoleCount.ToString());//完成任务的部门数
            iTotalCount = m_dt_TaskConfig.Rows.Count;                       //任务的总数
            //listResultOfTask.Add(m_dt_TaskConfig.Rows.Count.ToString());    //任务的总数
            //listResultOfTask.Add(listTaskFinish.Count.ToString());  //完成任务的个数
            iTotalFinishCount = listTaskFinish.Count;           //完成任务的个数
            double iRateFinish = Math.Round((double)listTaskFinish.Count / m_dt_TaskConfig.Rows.Count, 2);//任务完成率
            iRateFinish *= 100;
            listResultOfTask.Add(iRateFinish.ToString() + "%");
            return listResultOfTask;
        }
       
        public string GetUserNames(string sRoleName)//以（‘’，‘’）的形式返回该组的用户
        {
            string sUserNames = "(";
            DataView dv_RoleID = new DataView(m_dt_RoleType);
            dv_RoleID.RowFilter = "RoleName = '" + sRoleName + "'";
            DataTable dt_RoleID = dv_RoleID.ToTable();
            DataView dv_UserName = new DataView(m_dt_RoleType);
            dv_UserName.RowFilter = "RoleID like '" + dt_RoleID.Rows[0]["RoleID"] + "%'";
            DataTable dt_UserName = dv_UserName.ToTable();
            for (int i = 0; i < dt_UserName.Rows.Count; i++)
            {
                if (dt_UserName.Rows[i]["RoleID"].ToString().Length == 10)
                {
                    sUserNames += "'" + dt_UserName.Rows[i]["RoleName"].ToString() + "',";
                }
            }
            sUserNames = sUserNames.Substring(0, sUserNames.Length - 1);
            sUserNames += ")";
            return sUserNames;
        }

        public int GetUserCount(string sRoleName)//返回该组用户的个数
        {
            int iUserCount = 0;
            DataView dv_RoleID = new DataView(m_dt_RoleType);
            dv_RoleID.RowFilter = "RoleName = '" + sRoleName + "'";
            DataTable dt_RoleID = dv_RoleID.ToTable();
            DataView dv_UserName = new DataView(m_dt_RoleType);
            dv_UserName.RowFilter = "RoleID like '" + dt_RoleID.Rows[0]["RoleID"] + "%'";
            DataTable dt_UserName = dv_UserName.ToTable();
            for (int i = 0; i < dt_UserName.Rows.Count; i++)
            {
                if (dt_UserName.Rows[i]["RoleID"].ToString().Length == 10)
                {
                    iUserCount++;
                }
            }
            return iUserCount;
        }


        //第四章
        public string GetResultDetail(out List<string> listRoleResult)
        {
            string sResultDetail = "";
            listRoleResult = GetRoleOfResult(GetUsersResult());
            int[] sQuesCount = new int[listRoleResult.Count];//每个组的题数
            float[] sAverage = new float[listRoleResult.Count]; //每个组的平均分
            for (int i = 0; i < listRoleResult.Count; i++)
            {
                DataView dv_UserResultDetai = new DataView(m_dt_UserResultDetai);
                dv_UserResultDetai.RowFilter = "UserName in " + GetUserNames(listRoleResult[i]);
                DataTable dt = dv_UserResultDetai.ToTable();
                sQuesCount[i] = dt.Rows.Count;

                DataView dv_UsersResult = new DataView(m_dt_UsersResult);
                dv_UsersResult.RowFilter = "RoleName = '" + listRoleResult[i] + "'";
                DataTable dt_UsersResult = dv_UsersResult.ToTable();
                int iTotal = 0;
                for (int j = 0; j < dt_UsersResult.Rows.Count; j++)
                {
                    iTotal += int.Parse(dt_UsersResult.Rows[j]["AllScore"].ToString());
                }
                if (dt_UsersResult.Rows.Count != 0)
                    sAverage[i] = (float)iTotal / dt_UsersResult.Rows.Count;
                else
                    sAverage[i] = 0;
            }
            int iTotalCount = 0;
            for (int i = 0; i < sQuesCount.Length; i++)
            {
                iTotalCount += sQuesCount[i];
            }
            sResultDetail = "此次演练答题，总共有" + iTotalCount.ToString() + "道题。";
            if (iTotalCount>0)
                sResultDetail += "其中"; 
            for (int i = 0; i < sQuesCount.Length; i++)
            {
                sResultDetail += listRoleResult[i] + "答了" + sQuesCount[i] + "道题，平均分为" + sAverage[i] + "分；";
            }
            sResultDetail = sResultDetail.Substring(0, sResultDetail.Length - 1);
            sResultDetail += "。";
            return sResultDetail;
        }
        //获取4.1三级标题文本
        public string GetTitleOf4_1(int index, string sRoleName)
        {
            string str = "4.1." + index + " " + sRoleName + "答题情况";
            return str;
        }

        //4.1.分析情况
        //out tblUserResultDetail中有记录的用户
        public string GetDetail(string sRoleName, out List<string> listExitUserResultDetail)
        {

            listExitUserResultDetail = new List<string>();
            DataView dv_UserResult = new DataView(GetUsersResult());
            dv_UserResult.RowFilter = "RoleName = '" + sRoleName + "'";
            DataTable dt_Users = dv_UserResult.ToTable();//该组用户名的集合
         
            int iUserCount = dt_Users.Rows.Count;//该组人数
            int iPassCount = 0;//该组及格人数
            for (int i = 0; i < dt_Users.Rows.Count; i++ )
            {
                if (IsPass(dt_Users.Rows[i]["UserName"].ToString()))
                    iPassCount++;

                DataView dv_UserExist = new DataView(m_dt_UserResultDetai);
                dv_UserExist.RowFilter = "UserName ='" + dt_Users.Rows[i]["UserName"].ToString() + "'";
                DataTable dt_UserExist = dv_UserExist.ToTable();
           
                if (dt_UserExist.Rows.Count != 0)
                    listExitUserResultDetail.Add(dt_Users.Rows[i]["UserName"].ToString());//该用户有明细表数据，添加
            }
            string str = sRoleName + "有职工" + iUserCount + "人参与，" + iPassCount + "人及格，" + (iUserCount - iPassCount) + "人不及格。";
            if (listExitUserResultDetail.Count != 0)
                str += "下面是职工";
            for (int i = 0; i < listExitUserResultDetail.Count; i++)
            {
                str += listExitUserResultDetail[i] + "、";
            }
            if (listExitUserResultDetail.Count != 0)
            {
                str = str.Substring(0, str.Length - 1);
                str += "的答题成绩单。";
            }
            return str;
        }

        public string Get4_1TableName(string sUserName)
        {
            string str = sUserName + "的成绩表格";
            return str;
        }

        public DataTable GetUserTable(string sUserName)
        {
            DataView dv = new DataView(m_dt_UserResultDetai);
            dv.RowFilter = "UserName ='" + sUserName + "'";
            DataTable dt = dv.ToTable();
            return dt;
        }

       
        //获取4.2三级标题文本
        public string GetTitleOf4_2(int index, string sRoleName)
        {
            string str = "4.2." + index + " " + sRoleName + "任务完成情况";
            return str;
        }

        public int GetCountUserDoTask(string sRoleName)
        {
            DataView dv_DoTask = new DataView(m_dt_DoTask);
            dv_DoTask.RowFilter = "Completer in " + GetUserNames(sRoleName);
            DataTable dt_DoTask = dv_DoTask.ToTable();
            return dt_DoTask.Rows.Count;
        }

        public DataTable GetDataTableUserDoTask(string sRoleName)
        {
            DataView dv_UserDoTask = new DataView(m_dt_DoTask);
            dv_UserDoTask.RowFilter = "Completer in" + GetUserNames(sRoleName);
            DataTable dt_UserDoTask = dv_UserDoTask.ToTable();
            return dt_UserDoTask;
        }

        public DataTable GetTaskConfig(string sTaskID)
        {
            DataView dv_TaskConfig = new DataView(m_dt_TaskConfig);
            dv_TaskConfig.RowFilter = "TaskID = '" + sTaskID + "'";
            DataTable dt_TaskConfig = dv_TaskConfig.ToTable();
            return dt_TaskConfig;
        }

        //iScore 为用户答题和完成任务的分数
        //总分为答题分数+任务分数+每个任务的评价分数
        public int GetTotalScore(string sUsers,int iScore)
        {
            DataView dv_UserDoTask = new DataView(m_dt_DoTask);
            dv_UserDoTask.RowFilter = "Completer = '" + sUsers + "'";
            DataTable dt_UserDoTask = dv_UserDoTask.ToTable();
            double iTaskRate = 0;
            for (int i = 0; i < dt_UserDoTask.Rows.Count; i++)
            {
                iTaskRate += GetAvgTaskRate(dt_UserDoTask.Rows[i]["TaskID"].ToString());
            }
            return iScore + (int)iTaskRate;
        }

        public double GetAvgTaskRate(string sTaskID)
        {
            double iTaskRte = 0;
            DataView dv_TaskRate = new DataView(m_dt_TaskRate);
            dv_TaskRate.RowFilter = "TaskID = '" + sTaskID + "'";
            DataTable dt_TaskRate = dv_TaskRate.ToTable();
            if (dt_TaskRate.Rows.Count != 0)
            {
                string[] str = dt_TaskRate.Rows[0]["TaskRate"].ToString().Split(',');
                for (int j = 0; j < str.Length; j++)
                {
                    iTaskRte += int.Parse(str[j]);
                }
                iTaskRte = Math.Round(iTaskRte / str.Length, 2);
            }
            return iTaskRte;
        }

        public List<string> GetText(string sFileName)
        {
            string path = System.IO.Path.GetDirectoryName(m_sFilePath);
            path += "/" + sFileName;
            List<string> listText = new List<string>();
            System.IO.StreamReader reader = null;
            try
            {
                reader = new System.IO.StreamReader(path, System.Text.Encoding.Default);
                reader.BaseStream.Seek(0, System.IO.SeekOrigin.Begin);
                string line = reader.ReadLine();
                while (line != null)
                {
                    listText.Add(line);
                    line = reader.ReadLine();
                }
            }
            catch (Exception e)
            {
            }
            finally
            {
                if(reader != null)
                reader.Close();
            }
            return listText;
        }

        public string GetRunTime()
        {
            string sql = "select * from tblParaControl where ParaName = 'RunTime'";
            DataTable dt = m_db.GetTableInfo(sql);
            return dt.Rows[0]["ParaValue"].ToString();
        }

        public string GetExitTime()
        {
            string sql = "select * from tblParaControl where ParaName = 'ExitTime'";
            DataTable dt = m_db.GetTableInfo(sql);
            return dt.Rows[0]["ParaValue"].ToString();
        }
    }
}
