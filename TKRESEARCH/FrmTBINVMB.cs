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
using System.Collections;
using TKITDLL;
using System.Xml;
using System.Text.RegularExpressions;

namespace TKRESEARCH
{
    public partial class FrmTBINVMB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        int result;

        public FrmTBINVMB()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='TBINVMB'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='UPDATEUSER'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARANAME";
            comboBox2.DisplayMember = "PARANAME";
            sqlConn.Close();


        }

        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='TBINVMB'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "PARANAME";
            comboBox3.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void SEARCH(string KINDS,string MB013,string MB002,string MB001)
        {
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

                StringBuilder SQLQUERY1 = new StringBuilder();
                StringBuilder SQLQUERY2 = new StringBuilder();
                StringBuilder SQLQUERY3 = new StringBuilder();
                StringBuilder SQLQUERY4 = new StringBuilder();

                sbSql.Clear();
                SQLQUERY1.Clear();
                SQLQUERY2.Clear();
                SQLQUERY3.Clear();
                SQLQUERY4.Clear();



                if (!string.IsNullOrEmpty(KINDS)&& KINDS.Equals("空白"))
                {
                    SQLQUERY1.AppendFormat(@"
                                        AND ISNULL(MB002,'')=''
                                        ");
                }
                else if  (!string.IsNullOrEmpty(KINDS) && KINDS.Equals("非空白"))
                {
                    SQLQUERY1.AppendFormat(@"
                                        AND ISNULL(MB002,'')<>''
                                        ");
                }
                else
                {
                    SQLQUERY1.AppendFormat(@"  ");
                }
                if (!string.IsNullOrEmpty(MB013))
                {
                    SQLQUERY2.AppendFormat(@"
                                        AND MB013 LIKE '%{0}%'
                                        ", MB013);
                }
                else
                {
                    SQLQUERY2.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(MB002))
                {
                    SQLQUERY3.AppendFormat(@"
                                        AND MB002 LIKE '%{0}%'
                                        ", MB002);
                }
                else
                {
                    SQLQUERY3.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(MB001))
                {
                    SQLQUERY4.AppendFormat(@"
                                        AND MB001 LIKE '%{0}%'
                                        ", MB001);
                }
                else
                {
                    SQLQUERY4.AppendFormat(@"");
                }



                sbSql.AppendFormat(@"                             
                                    SELECT 
                                    [MB013] AS '條碼'
                                    ,[MB002] AS '品名'
                                    ,[MB001] AS '品號'
                                    ,[MB003] AS '規格(重量)'
                                    ,[MB004] AS '單位'
                                    ,[MODIDATE] AS '日期'
                                    ,[COMMENTS] AS '備註'
                                    ,[UPDATEUSER] AS '更新人員'

                                    FROM [TKRESEARCH].[dbo].[TBINVMB]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    {2}
                                    {3}
                                    ORDER BY MB013
                        

	 

                                ", SQLQUERY1.ToString(), SQLQUERY2.ToString(), SQLQUERY3.ToString(), SQLQUERY4.ToString());



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


        public void CHECK_MB013(string MB013,string UPDATEUSER)
        {
            string STR = MB013;
            char[] CHARS = STR.ToCharArray();

            int ODDS = 0;
            int EVENS = 0;
            int TOTALS = 0;
            int CAL = 0;
            int POSITION = 1;
            string CHECKCODE = "";
            string GET_CHECKCODE = "";

            //條碼+驗証碼的長度不是13碼
            //前12碼是條碼，第13碼是驗証碼
            //偶數位相加，*3
            //奇數位相加
            //奇數位相加+偶數位相加，取出/10的餘數，當驗証碼
            if (STR.Length!=13)
            {
                MessageBox.Show("條碼+驗証碼的長度不是13碼，請修改正確");
            }
            else
            {
                foreach (char C in CHARS)
                {
                    if(POSITION<=12)
                    {
                        if (POSITION % 2 == 1)
                        {
                            ODDS += int.Parse(C.ToString());
                        }
                        else if (POSITION % 2 == 0)
                        {
                            EVENS += int.Parse(C.ToString());
                        }
                    }
                    
                    if(POSITION==13)
                    {
                        GET_CHECKCODE = Convert.ToString(C);
                    }

                    POSITION = POSITION + 1;
                }

                EVENS = EVENS * 3;
                TOTALS = ODDS + EVENS;
                CAL = TOTALS % 10;
                CAL = 10 - CAL;
                CHECKCODE = Convert.ToString(CAL);         

                //MessageBox.Show(CHECKCODE);

                if(!CHECKCODE.Equals(GET_CHECKCODE))
                {
                    MessageBox.Show("驗証碼錯諤");
                }
                else if (CHECKCODE.Equals(GET_CHECKCODE))
                {
                    ADD_MB013(MB013, UPDATEUSER);
                }
            }
           

        }

        public void ADD_MB013(string MB013,string UPDATEUSER)
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
                                       INSERT INTO [TKRESEARCH].[dbo].[TBINVMB]
                                        ([MB013],[MB002],[MB001],[UPDATEUSER])
                                        VALUES
                                        ('{0}','','','{1}')

                                        ", MB013, UPDATEUSER);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("失敗");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("完成");
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


        public void UPDATE_MB002(string MB013,string MB002,string MB003,string MB004,string COMMENTS,string UPDATEUSER)
        {
            if(!string.IsNullOrEmpty(MB013)&&!string.IsNullOrEmpty(MB002))
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
                                        UPDATE  [TKRESEARCH].[dbo].[TBINVMB]
                                        SET MB002='{1}',MB003='{2}',MB004='{3}',COMMENTS='{4}',[UPDATEUSER]='{5}',[MODIDATE]=CONVERT(NVARCHAR,GETDATE(),112)
                                        WHERE MB013='{0}'
                                        ", MB013,MB002, MB003, MB004, COMMENTS, UPDATEUSER);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        MessageBox.Show("完成");
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

        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox5.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox5.Text = row.Cells["條碼"].Value.ToString();
                    textBox6.Text = row.Cells["品名"].Value.ToString();
                    textBox7.Text = row.Cells["品號"].Value.ToString();

                    FIND_ERP_INBMB_MB013(row.Cells["條碼"].Value.ToString());
                }
                else
                {
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                }
            }
        }

        public void UPDATE_MB001( string MB013, string MB001)
        {
            if (!string.IsNullOrEmpty(MB001) && !string.IsNullOrEmpty(MB013))
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
                                        UPDATE  [TKRESEARCH].[dbo].[TBINVMB]
                                        SET MB001='{1}'
                                        WHERE MB013='{0}'
                                        ", MB013, MB001);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        //MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        //MessageBox.Show("完成");
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
        }

       

        public void FIND_ERP_INVMB(string MB001,string MB013)
        {
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
     
                sbSql.Clear();

                if(!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"                             
                                        SELECT * 
                                        FROM  [TK].dbo.INVMB
                                        WHERE MB001='{0}' 

                                        ",MB001);
                }




                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    MessageBox.Show("品號不存在ERP，請先設定ERP的品號");
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        UPDATE_ERP_INVMB(MB001, MB013);
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

        public void UPDATE_ERP_INVMB(string MB001,string MB013)
        {
            if(!string.IsNullOrEmpty(MB001)&& !string.IsNullOrEmpty(MB013))
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
                                        UPDATE  [TK].[dbo].[INVMB]
                                        SET MB013='{1}'
                                        WHERE MB001='{0}'
                                        ", MB001, MB013);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        MessageBox.Show("完成");
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
        }

        public void FIND_ERP_INBMB_MB013(string MB013)
        {
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
             

                sbSql.Clear();

                sbSql.AppendFormat(@"                             
                                    SELECT 
                                    [MB013] AS '條碼'
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'                                   
                                    FROM [TK].[dbo].[INVMB]
                                    WHERE 1=1
                                    AND MB013='{0}'
                                    ORDER BY MB013                 

	 

                                ", MB013);



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

                        //字體
                        dataGridView2.DefaultCellStyle.Font = new Font("Arial", 12, FontStyle.Bold);
                        // 修改預設儲存格的背景顏色和前景顏色
                        dataGridView2.DefaultCellStyle.BackColor = Color.LightGray;
                        dataGridView2.DefaultCellStyle.ForeColor = Color.Blue;
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

        public void UPDATE_MB001_NULL(string MB013, string MB001)
        {
            if (!string.IsNullOrEmpty(MB013))
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
                                        UPDATE  [TKRESEARCH].[dbo].[TBINVMB]
                                        SET MB001=NULL
                                        WHERE MB013='{0}'
                                        ", MB013, MB001);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        //MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        //MessageBox.Show("完成");
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
        }

        public void UPDATE_MB002_NULL(string MB013, string MB002)
        {
            if (!string.IsNullOrEmpty(MB013))
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
                                        UPDATE  [TKRESEARCH].[dbo].[TBINVMB]
                                        SET MB002=NULL,MB003=NULL,MB004=NULL,COMMENTS=NULL
                                        WHERE MB013='{0}'
                                        ", MB013, MB002);

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                        //MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        //MessageBox.Show("完成");
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
        }

        public void SEARCH2(string KINDS, string MB013, string MB002, string MB001)
        {
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

                StringBuilder SQLQUERY1 = new StringBuilder();
                StringBuilder SQLQUERY2 = new StringBuilder();
                StringBuilder SQLQUERY3 = new StringBuilder();
                StringBuilder SQLQUERY4 = new StringBuilder();

                sbSql.Clear();
                SQLQUERY1.Clear();
                SQLQUERY2.Clear();
                SQLQUERY3.Clear();
                SQLQUERY4.Clear();



                if (!string.IsNullOrEmpty(KINDS) && KINDS.Equals("空白"))
                {
                    SQLQUERY1.AppendFormat(@"
                                        AND ISNULL(MB002,'')=''
                                        ");
                }
                else if (!string.IsNullOrEmpty(KINDS) && KINDS.Equals("非空白"))
                {
                    SQLQUERY1.AppendFormat(@"
                                        AND ISNULL(MB002,'')<>''
                                        ");
                }
                else
                {
                    SQLQUERY1.AppendFormat(@"  ");
                }
                if (!string.IsNullOrEmpty(MB013))
                {
                    SQLQUERY2.AppendFormat(@"
                                        AND MB013 LIKE '%{0}%'
                                        ", MB013);
                }
                else
                {
                    SQLQUERY2.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(MB002))
                {
                    SQLQUERY3.AppendFormat(@"
                                        AND MB002 LIKE '%{0}%'
                                        ", MB002);
                }
                else
                {
                    SQLQUERY3.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(MB001))
                {
                    SQLQUERY4.AppendFormat(@"
                                        AND MB001 LIKE '%{0}%'
                                        ", MB001);
                }
                else
                {
                    SQLQUERY4.AppendFormat(@"");
                }



                sbSql.AppendFormat(@"                             
                                    SELECT 
                                    [MB013] AS '條碼'
                                    ,[MB002] AS '品名'
                                    ,[MB001] AS '品號'
                                    ,[MB003] AS '規格(重量)'
                                    ,[MB004] AS '單位'
                                    ,[MODIDATE] AS '日期'
                                    ,[COMMENTS] AS '備註'
                                    ,[UPDATEUSER] AS '更新人員'

                                    FROM [TKRESEARCH].[dbo].[TBINVMB]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    {2}
                                    {3}
                                    ORDER BY MB013
                        

	 

                                ", SQLQUERY1.ToString(), SQLQUERY2.ToString(), SQLQUERY3.ToString(), SQLQUERY4.ToString());



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["ds"];
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(comboBox1.Text.ToString(),textBox1.Text.ToString(), textBox2.Text.ToString(),textBox3.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CHECK_MB013(textBox4.Text.Trim(),comboBox2.Text.ToString());
            SEARCH(comboBox1.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATE_MB002(textBox5.Text, textBox6.Text, textBox8.Text, textBox9.Text, textBox10.Text,comboBox2.Text.ToString());
            SEARCH(comboBox1.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString());

        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATE_MB001(textBox5.Text, textBox7.Text);
            FIND_ERP_INVMB(textBox7.Text,textBox5.Text);
            SEARCH(comboBox1.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("清空?", "清空?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                UPDATE_MB001_NULL(textBox5.Text, "");
                SEARCH(comboBox1.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

           
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("清空?", "清空?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                UPDATE_MB002_NULL(textBox5.Text, "");
                SEARCH(comboBox1.Text.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SEARCH2(comboBox3.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString());
        }

        #endregion


    }
}
