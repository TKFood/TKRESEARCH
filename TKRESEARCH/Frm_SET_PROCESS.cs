using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using FastReport;
using FastReport.Data;
using System.Net.Mail;
using TKITDLL;
using System.Globalization;
using System.Collections;

namespace TKRESEARCH
{
    public partial class Frm_SET_PROCESS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string talbename = null;
        int rownum = 0;
        int result;
 


        public Frm_SET_PROCESS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH_TKMOC_REPORTMOCBOMPROCESS(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                StringBuilder SQLquery1 = new StringBuilder();               

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                     [MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[PROCESSING] AS '生產說明'
                                    FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                                    WHERE ([MB001] LIKE '%{0}%' OR [MB002]  LIKE '%{0}%')
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
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
        public void SEARCH_TKMOC_REPORTMOCBOMORIPROCESS(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                StringBuilder SQLquery1 = new StringBuilder();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                     [MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[PROCESSING] AS '生產說明'
                                    FROM [TKMOC].[dbo].[REPORTMOCBOMORIPROCESS]
                                    WHERE ([MB001] LIKE '%{0}%' OR [MB002]  LIKE '%{0}%')
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["ds"];
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                DataGridViewRow row = dataGridView1.Rows[rowindex];

                textBox2.Text = row.Cells["品號"].Value.ToString();
                textBox3.Text = row.Cells["品名"].Value.ToString();
                textBox4.Text = row.Cells["生產說明"].Value.ToString();
            }
        }
        public void ADD_REPORTMOCBOMPROCESS(string MB001, string MB002, string PROCESSING)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                                    (
                                     MB001
                                    ,MB002
                                    ,PROCESSING
                                    )
                                    VALUES
                                    (
                                     '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    )
                                    "

                                    , MB001, MB002, PROCESSING);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void UPDATE_REPORTMOCBOMPROCESS(string MB001, string MB002, string PROCESSING)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                    
                                    UPDATE  [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                                    SET MB002='{1}',PROCESSING='{2}'
                                    WHERE MB001='{0}'
                                    "

                                    , MB001, MB002, PROCESSING);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void DELETE_REPORTMOCBOMPROCESS(string MB001)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                    
                                    DELETE  [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                                    WHERE MB001='{0}'
                                    "

                                    , MB001);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                DataGridViewRow row = dataGridView2.Rows[rowindex];

                textBox5.Text = row.Cells["品號"].Value.ToString();
                textBox6.Text = row.Cells["品名"].Value.ToString();
                textBox7.Text = row.Cells["生產說明"].Value.ToString();
            }
        }

        public void ADD_REPORTMOCBOMORIPROCESS(string MB001, string MB002, string PROCESSING)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMORIPROCESS]
                                    (
                                     MB001
                                    ,MB002
                                    ,PROCESSING
                                    )
                                    VALUES
                                    (
                                     '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    )
                                    "

                                    , MB001, MB002, PROCESSING);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void UPDATE_REPORTMOCBOMORIPROCESS(string MB001, string MB002, string PROCESSING)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                    
                                    UPDATE  [TKMOC].[dbo].[REPORTMOCBOMORIPROCESS]
                                    SET MB002='{1}',PROCESSING='{2}'
                                    WHERE MB001='{0}'
                                    "

                                    , MB001, MB002, PROCESSING);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void DELETE_REPORTMOCBOMORIPROCESS(string MB001)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                    
                                    DELETE  [TKMOC].[dbo].[REPORTMOCBOMORIPROCESS]
                                    WHERE MB001='{0}'
                                    "

                                    , MB001);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_TKMOC_REPORTMOCBOMPROCESS(textBox1.Text.Trim());
            SEARCH_TKMOC_REPORTMOCBOMORIPROCESS(textBox1.Text.Trim()); 
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADD_REPORTMOCBOMPROCESS(textBox2.Text.Trim(), textBox3.Text.Trim(), textBox4.Text.Trim());
            SEARCH_TKMOC_REPORTMOCBOMPROCESS(textBox1.Text.Trim());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            UPDATE_REPORTMOCBOMPROCESS(textBox2.Text.Trim(), textBox3.Text.Trim(), textBox4.Text.Trim());
            SEARCH_TKMOC_REPORTMOCBOMPROCESS(textBox1.Text.Trim());
        }

        private void button4_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_REPORTMOCBOMPROCESS(textBox2.Text.Trim());
                SEARCH_TKMOC_REPORTMOCBOMPROCESS(textBox1.Text.Trim());

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ADD_REPORTMOCBOMORIPROCESS(textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
            SEARCH_TKMOC_REPORTMOCBOMORIPROCESS(textBox1.Text.Trim());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UPDATE_REPORTMOCBOMORIPROCESS(textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
            SEARCH_TKMOC_REPORTMOCBOMORIPROCESS(textBox1.Text.Trim());
        }
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_REPORTMOCBOMORIPROCESS(textBox5.Text.Trim());
                SEARCH_TKMOC_REPORTMOCBOMORIPROCESS(textBox1.Text.Trim());

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        #endregion


    }
}
