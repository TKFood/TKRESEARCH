﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using FastReport;
using FastReport.Data;
using System.Net.Mail;//<-基本上發mail就用這個class
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Diagnostics;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using TKITDLL;
using System.Net.Http;
using System.Net;

namespace TKRESEARCH
{
    public partial class FrmUOFMESSAGES : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;

        int result;

        string SUBJECTS = null;

        public FrmUOFMESSAGES()
        {
            InitializeComponent();

            SETETXT();
        }
        private void FrmUOFMESSAGES_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 20;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);


            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 8;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 1);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView3.Controls.Add(cbHeader);
        }
        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView3.EndEdit();

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView3.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        #region FUNCTION

        public void SETETXT()
        {
            textBox1.Text = DateTime.Now.Year.ToString();
        }
        public void SEARCH_TB_EIP_SCH_DEVOLVE(string SUBJECT,string CREATE_TIME)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                if (!string.IsNullOrEmpty(SUBJECT))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND TB_EIP_SCH_DEVOLVE.SUBJECT LIKE '%{0}%'
                                            ", SUBJECT);
                }

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_DEVOLVE.CREATE_TIME
                                    ,*
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE

                                    WHERE 1=1
                                    AND TB_EIP_SCH_DEVOLVE.SUBJECT LIKE '%校稿%'
                                    AND CONVERT(NVARCHAR,TB_EIP_SCH_DEVOLVE.CREATE_TIME,112) LIKE '{1}%'
                                    {0}
                                 
                                    ORDER BY TB_EIP_SCH_DEVOLVE.SUBJECT

                                    ", sbSqlQuery.ToString(), CREATE_TIME);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();


                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    string DEVOLVE_GUID = row.Cells["DEVOLVE_GUID"].Value.ToString().Trim();
                    string SUBJECT = row.Cells["校稿區內容"].Value.ToString().Trim();
                    SUBJECTS = row.Cells["校稿區內容"].Value.ToString().Trim();

                    SEARCH_TB_EIP_SCH_DEVOLVE_DETAIL(DEVOLVE_GUID);

                    textBox2.Text = "通知: "+ Environment.NewLine+ Environment.NewLine;
                    textBox2.Text = textBox2.Text + SUBJECT+ Environment.NewLine+ Environment.NewLine;
                    textBox2.Text = textBox2.Text + "完成";

                }
                else
                {
                    
                }
            }
            
        }
        public void SEARCH_TB_EIP_SCH_DEVOLVE_DETAIL(string DEVOLVE_GUID)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

               

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'

                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
                                    ,USER2.NAME AS '交辨人'
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID ='{0}'
                                    ORDER BY TB_EIP_SCH_DEVOLVE.CREATE_TIME

                                    ", DEVOLVE_GUID);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();


                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCH_TB_EB_USER(string NAME)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                if (!string.IsNullOrEmpty(NAME))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND TB_EB_USER.NAME LIKE '%{0}%'
                                            ", NAME);
                }

                sbSql.AppendFormat(@"  
                                     SELECT TB_EB_USER.ACCOUNT AS '工號',TB_EB_USER.NAME AS '姓名',TB_EB_USER.USER_GUID
                                    FROM [UOF].[dbo].TB_EB_USER,[UOF].[dbo].TB_EB_EMPL
                                    WHERE 1=1
                                    AND TB_EB_USER.USER_GUID=TB_EB_EMPL.USER_GUID
                                    AND TB_EB_USER.IS_SUSPENDED<>'1'
                                    AND ISNULL(TB_EB_EMPL.BIRTHDAY,'')<>''
                                    AND  TB_EB_USER.ACCOUNT  COLLATE Chinese_Taiwan_Stroke_BIN   IN (SELECT [ID] FROM [192.168.1.105].[TKRESEARCH].[dbo].[TKUOFSNEDUSERS])        
                                    {0}

                                    ORDER BY NAME

                                    ", sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["TEMPds1"];
                        dataGridView3.AutoResizeColumns();


                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void NEW_MESSAGE()
        {
            string MESSAGE_TO = "";
            string MESSAGE_FROM = "916e213c-7b2e-46e3-8821-b7066378042b";

            StringBuilder TEXTBOX = new StringBuilder();

            for (int i = 0; i < textBox2.Lines.Length; i++)
            {
                TEXTBOX.AppendFormat("<p>" + textBox2.Lines[i] + "</p>");
            }

            foreach (DataGridViewRow DR in dataGridView3.Rows)
            {
                if (Convert.ToBoolean(DR.Cells[0].Value) == true)
                {
                    MESSAGE_TO = DR.Cells["USER_GUID"].Value.ToString();

                    StringBuilder MESSAGE_CONTENT = new StringBuilder();

                    MESSAGE_CONTENT.AppendFormat(TEXTBOX.ToString());
                   


                    ADD_UOF_TB_EIP_PRIV_MESS(
                    Guid.NewGuid().ToString("")
                    , "校和完成通知: "+ SUBJECTS
                    , MESSAGE_CONTENT.ToString()
                    , MESSAGE_TO
                    , MESSAGE_FROM
                    , ""
                    , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK")
                    , ""
                    , ""
                    , "0"
                    , "0"
                    , ""
                    , Guid.NewGuid().ToString("")
                    , ""
                    );
                }
            }

            //強制送給研發
            //210018	陳奕廷	

            ADD_UOF_TB_EIP_PRIV_MESS(
            Guid.NewGuid().ToString("")
            , "校和完成通知: " + SUBJECTS
            , TEXTBOX.ToString()
            , "58902b6b-ba4c-4dc0-b13b-427892786941"
            , MESSAGE_FROM
            , ""
            , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK")
            , ""
            , ""
            , "0"
            , "0"
            , ""
            , Guid.NewGuid().ToString("")
            , ""
            );

        }

        public void ADD_UOF_TB_EIP_PRIV_MESS(
            string MESSAGE_GUID
            , string TOPIC
            , string MESSAGE_CONTENT
            , string MESSAGE_TO
            , string MESSAGE_FROM
            , string REPLY_MESSAGE_GUID
            , string CREATE_TIME
            , string READED_TIME
            , string REPLY_TIME
            , string FROM_DELETED
            , string TO_DELETED
            , string FILE_GROUP_ID
            , string MASTER_GUID
            , string EVENT_ID

            )
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(MESSAGE_TO))
                    {
                        StringBuilder SBSQL = new StringBuilder();
                        SBSQL.AppendFormat(@" 
                                            INSERT INTO [UOF].[dbo].[TB_EIP_PRIV_MESS]
                                            (
                                            [MESSAGE_GUID]
                                            ,[TOPIC]
                                            ,[MESSAGE_CONTENT]
                                            ,[MESSAGE_TO]
                                            ,[MESSAGE_FROM]
                                            ,[REPLY_MESSAGE_GUID]
                                            ,[CREATE_TIME]
                                            ,[READED_TIME]
                                            ,[REPLY_TIME]
                                            ,[FROM_DELETED]
                                            ,[TO_DELETED]
                                            ,[FILE_GROUP_ID]
                                            ,[MASTER_GUID]
                                            ,[EVENT_ID]
                                            )
                                            VALUES
                                            (
                                            NEWID()
                                            ,@TOPIC
                                            ,@MESSAGE_CONTENT
                                            ,@MESSAGE_TO
                                            ,@MESSAGE_FROM
                                            ,''
                                            ,@CREATE_TIME
                                            ,NULL
                                            ,NULL
                                            ,'0'
                                            ,'0'
                                            ,''
                                            ,NEWID()
                                            ,''
                                            )

                                            ");

                        string sql = SBSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@MESSAGE_GUID", MESSAGE_GUID);
                            cmd.Parameters.AddWithValue("@TOPIC", TOPIC);
                            cmd.Parameters.AddWithValue("@MESSAGE_CONTENT", MESSAGE_CONTENT);
                            cmd.Parameters.AddWithValue("@MESSAGE_TO", MESSAGE_TO);
                            cmd.Parameters.AddWithValue("@MESSAGE_FROM", MESSAGE_FROM);
                            cmd.Parameters.AddWithValue("@REPLY_MESSAGE_GUID", REPLY_MESSAGE_GUID);
                            cmd.Parameters.AddWithValue("@CREATE_TIME", CREATE_TIME);
                            cmd.Parameters.AddWithValue("@READED_TIME", READED_TIME);
                            cmd.Parameters.AddWithValue("@REPLY_TIME", REPLY_TIME);
                            cmd.Parameters.AddWithValue("@FROM_DELETED", FROM_DELETED);
                            cmd.Parameters.AddWithValue("@TO_DELETED", TO_DELETED);
                            cmd.Parameters.AddWithValue("@FILE_GROUP_ID", FILE_GROUP_ID);
                            cmd.Parameters.AddWithValue("@MASTER_GUID", MASTER_GUID);




                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("失敗");
            }
        }

        

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_TB_EIP_SCH_DEVOLVE(textBox6.Text, textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH_TB_EB_USER(textBox3.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            NEW_MESSAGE();

            MessageBox.Show("完成");
        }

        #endregion

       
    }
}
