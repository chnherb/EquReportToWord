using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace SetQues.DB
{
    class BaseDBData
    {

        public string strConnect;
        public OleDbConnection con;


        public BaseDBData(string sDatabaseFilePath)
        {
            strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sDatabaseFilePath + @";Persist Security Info=False;Mode=Share Deny None";
            con = new OleDbConnection(strConnect);
        }

        public string GetTableInfo()
        {
            StreamReader f = new StreamReader("DBConfig.txt");
            string str = f.ReadLine();
            f.Close();
            return str;
        }

        //获取查询结果的DataTable
        public DataTable GetTableInfo(string sql)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
                con.Open();
            }
            DataTable dt = new DataTable();
            OleDbDataAdapter ad = new OleDbDataAdapter(sql, con);
            ad.Fill(dt);
            con.Close();
            return dt;
        }

        //执行操作语句
        public string ExcuteSql(string sql)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
                con.Open();
            }
            try
            {
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //执行多条操作语句
        public string ExcuteSql(List<string> Listsql)
        {

            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
                con.Open();
            }

            try
            {
                for (int i = 0; i < Listsql.Count; i++)
                {
                    OleDbCommand cmd = new OleDbCommand(Listsql[i], con);
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //获取数据库中查询到的数量
        public int GetAddQuesCount(string tbName)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }

           // string sql = "select count(*)  from " + tbName + " where LEN(QID) = 6 ";
            string sql = "select count(*)  from " + tbName;
            OleDbCommand cmd = new OleDbCommand(sql, con);
            OleDbDataReader dr = cmd.ExecuteReader();

            using (dr)
            {
                if (dr.Read())
                {
                    string strCount = dr[0].ToString();
                    con.Close();
                    return int.Parse(strCount);
                }
            }
            con.Close();
            return -1;
        }

        public void GetStuAge(int stuAge)
        {
            string sql = "select * from tblStudent where StuAge=" + stuAge.ToString();
            GetTableInfo(sql);
        }

        public int GetStuAge(string stuNo)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
                con.Open();
            }

            string sql = "select StuAge from tblStudent where StuNo='" + stuNo + "'";
            OleDbCommand cmd = new OleDbCommand(sql, con);
            OleDbDataReader dr = cmd.ExecuteReader();

            string StuAge = "";

            if (dr.Read())
            {
                StuAge = dr[0].ToString();
                dr.Close();
            }

            con.Close();

            return int.Parse(StuAge);
        }

        //判断是否存在满足某个条件的信息
        public bool IsExistsQus(string tbName, string qusID)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }
            string sql = "select QID  from " + tbName + " where QID = '" + qusID + "'";
            OleDbCommand cmd = new OleDbCommand(sql, con);
            OleDbDataReader dr = cmd.ExecuteReader();

            using (dr)
            {
                if (dr.Read())
                {
                    dr.Close();
                    con.Close();
                    return true;
                }
            }
            con.Close();
            return false;
        }
		
		public bool IsExistsQus(string tableName, string fieldName, string proValue)
        {
            
           
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }
            string sql = "select " + fieldName + " from " + tableName + " where " + fieldName + " = '" + proValue + "'";
            OleDbCommand cmd = new OleDbCommand(sql, con);
            OleDbDataReader dr = cmd.ExecuteReader();

            using (dr)
            {
                if (dr.Read())
                {
                    dr.Close();
                    con.Close();
                    return true;
                }
            }
            con.Close();
            return false;
        }

        public DataTable GetAllData(string tableName, string KnowType)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }


            DataTable dt = new DataTable();
            string sql = "";

            switch (tableName)
            {
                case "tbItemSelect":
                    sql = "select false as 选择,QID as 题目编号,Title as 标题,QusOption as 选项,Answer as 答案,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemSelect ";
                    break;

                case "tbItemLine":
                    sql = "select false as 选择,QID as 题目编号,ImageId as 图片编号,Title as 标题,nHot as 热区个数,Answer as 答案,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemLine ";
                    break;

                case "tbItemBlankTwo":
                    sql = "select false as 选择,QID as 题目编号,ImageId as 图片编号,Title as 标题,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemBlankTwo ";
                    break;

                case "tbItemBlankFour":
                    sql = "select false as 选择,QID as 题目编号,ImageId as 图片编号,Title as 标题,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemBlankFour ";
                    break;
                case "tbItemBlankSeveral":
                    sql = "select false as 选择,QID as 题目编号,ImageId as 图片编号,Title as 标题,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemBlankSeveral ";
                    break;
                case "tbItemSteelTemp":
                    sql = "select false as 选择,QID as 题目编号,ImageId as 图片编号,Title as 标题,DifDegree as 难度系数,KnowledgeType as 知识点,Score as 分数,ExamTime as 答题时间 from tbItemSteelTemp ";
                    break;
            }

            if (KnowType != "" && KnowType != "Know")
                sql += " where KnowledgeType like '" + KnowType + "%'";
            OleDbDataAdapter ada = new OleDbDataAdapter(sql, con);

            ada.Fill(dt);
            con.Close();

            return dt;
        }

        //单选题操作
        public void InsertSingleData(string QID, string Title, string Option, int OpCount, string Answer, string DifDegree, string KnowledgeType, string Score)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }

            DataTable dt = new DataTable();
            string sql = "insert into tbItemSelect(QID,Title,QusOption,OptiCount,Answer,DifDegree,KnowledgeType,Score) values('" + QID + "','" + Title + "','" + Option + "'," + OpCount.ToString() + ",'" + Answer + "','" + DifDegree + "'," + KnowledgeType + ",'" + Score + "')";
            OleDbCommand cmd = new OleDbCommand(sql, con);

            cmd.ExecuteNonQuery();
            con.Close();
        }

        //获取题目编号的最大值
        public int GetQuesMaxId(string tbName)
        {

            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
            }

            //string sql = "select  MAX(ID)  from " + tbName + " where Title is not null and Title <> ''";
            string sql = "select  MAX(QID)  from " + tbName;
            OleDbCommand cmd = new OleDbCommand(sql, con);
            OleDbDataReader dr = cmd.ExecuteReader();

            using (dr)
            {
                if (dr.Read())
                {
                    string strCount = dr[0].ToString().Substring(2, 4);
                    con.Close();
                    return int.Parse(strCount);
                }
            }
            con.Close();
            return -1;
        }
    }
}
