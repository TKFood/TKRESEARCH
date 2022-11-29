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

namespace TKRESEARCH
{
    public partial class FrmINVMBCALCOST : Form
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

 


        DataSet ds = new DataSet();
      
        int rownum = 0;
        int result;
    

        string ID;


        public FrmINVMBCALCOST()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
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
            Sequel.AppendFormat(@"SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='CALCOSTPRODS'
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
            Sequel.AppendFormat(@"SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='CALCOSTPRODS'
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
        public void SEARCH(string PRODNAMES,string ISCLOESED)
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

                if (!string.IsNullOrEmpty(PRODNAMES))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT  
                                        [PRODNAMES] AS '產品名稱'
                                        ,[SPECS] AS '規格及重量'
                                        ,[COSTRAW] AS '原料成本'
                                        ,[COSTMATERIL] AS '物料成本'
                                        ,[COSTART] AS '人工成本(製造+內外包)'
                                        ,[COSTMANU] AS '製造費用' 
                                        ,[COSTTOTALS] AS '單位成本合計'
                                        ,CONVERT(NVARCHAR,[CREATEDATE],112) AS '建立日期'
                                        ,[ISCLOESED] AS '結案'
                                        ,[ID]

                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                        WHERE [PRODNAMES] LIKE '%{0}%'
                                        AND [ISCLOESED] IN ('{1}')
                                        ORDER BY CONVERT(NVARCHAR,[CREATEDATE],112)
                                        ", PRODNAMES, ISCLOESED);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT  
                                        [PRODNAMES] AS '產品名稱'
                                        ,[SPECS] AS '規格及重量'
                                        ,[COSTRAW] AS '原料成本'
                                        ,[COSTMATERIL] AS '物料成本'
                                        ,[COSTART] AS '人工成本(製造+內外包)'
                                        ,[COSTMANU] AS '製造費用' 
                                        ,[COSTTOTALS] AS '單位成本合計'
                                        ,CONVERT(NVARCHAR,[CREATEDATE],112) AS '建立日期'
                                        ,[ISCLOESED] AS '結案'
                                        ,[ID]

                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                        WHERE [ISCLOESED] IN ('{0}')

                                        ORDER BY CONVERT(NVARCHAR,[CREATEDATE],112)
                                   
                                    ", ISCLOESED);
                }


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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBoxID.Text = null;

            textBox5.Text = null;
            textBox6.Text = null;
            textBoxID2.Text = null;

            textBox16.Text = null;
            textBox17.Text = null;         
            textBoxID4.Text = null;

            textBox24.Text = null;
            textBox28.Text = null;
            textBoxID6.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;               

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox2.Text = row.Cells["規格及重量"].Value.ToString();
                    textBoxID.Text = row.Cells["ID"].Value.ToString();

                    textBox5.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox6.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID2.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG2(textBox5.Text.Trim());

                    textBox16.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox17.Text = row.Cells["產品名稱"].Value.ToString();                   
                    textBoxID4.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG3(textBox16.Text);

                    textBox24.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox28.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID6.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG4(textBox24.Text);

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBoxID.Text = null;

                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBoxID2.Text = null;

                    textBox16.Text = null;
                    textBox17.Text = null;                  
                    textBoxID4.Text = null;

                    textBox24.Text = null;
                    textBox28.Text = null;
                    textBoxID6.Text = null;
                }
            }
        }

        public void ADDCALCOSTPRODS(string PRODNAMES,string SPECS)
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
                                    INSERT INTO  [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                    ([PRODNAMES],[SPECS])
                                    VALUES
                                    ('{0}','{1}')

                                    ", PRODNAMES, SPECS);

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

        public void UPDATECALCOSTPRODS(string PRODNAMES, string SPECS,string ID,string ISCLOESED)
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
                                   UPDATE [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                    SET [PRODNAMES]='{1}',[SPECS]='{2}',[ISCLOESED]='{3}'
                                    WHERE [ID]='{0}'

                                    ", ID, PRODNAMES, SPECS, ISCLOESED);

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

        public void SEARCHDG2(string PRODNAMES)
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

                if (!string.IsNullOrEmpty(PRODNAMES))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [MID] AS 'ID'
                                        ,[PRODNAMES] AS '產品名稱'
                                        ,[MB001] AS '使用品號'
                                        ,[MB002] AS '使用品名'
                                        ,[INS] AS '投入重量'
                                        ,[PRICES] AS '單價'
                                        ,[TMONEYS] AS '金額'
                                        ,[REMARK] AS '備註'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS1RAW]
                                        WHERE [PRODNAMES] LIKE '%{0}%'
                                        ORDER BY [MID],[PRODNAMES],[MB001]
                                        ", PRODNAMES);
                }
               

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

        public void ADDCALCOSTPRODS1RAW(string MID
                                        , string PRODNAMES
                                        , string MB001
                                        , string MB002
                                        , string INS
                                        , string PRICES
                                        , string TMONEYS
                                        , string REMARK
                                        )
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
                                   
                                    INSERT INTO  [TKRESEARCH].[dbo].[CALCOSTPRODS1RAW]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MB001]
                                    ,[MB002]
                                    ,[INS]
                                    ,[PRICES]
                                    ,[TMONEYS]
                                    ,[REMARK]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    )
                                    ", MID
                                        , PRODNAMES
                                        , MB001
                                        , MB002
                                        , INS
                                        , PRICES
                                        , TMONEYS
                                        , REMARK);

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

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox8.Text = FINDMB002(textBox7.Text.ToString());
            textBox10.Text = FINDPRICES(textBox7.Text.ToString());
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = (Convert.ToDecimal(textBox9.Text) * Convert.ToDecimal(textBox10.Text)).ToString();
        }
        public string FINDMB002(string MB001)
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

                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT MB001,MB002 
                                        FROM [TK].dbo.INVMB
                                        WHERE MB001='{0}'

                                        ", MB001);
                }


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["MB002"].ToString();

                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }
        public string FINDPRICES(string MB001)
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

                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT MB001,MB002,MB050 
                                        FROM [TK].dbo.INVMB
                                        WHERE MB001='{0}'

                                        ", MB001);
                }


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["MB050"].ToString();
                    
                }
                else
                {
                    return "0";
                }

            }
            catch
            {
                return "0";
            }
            finally
            {

            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
           
            textBoxID3.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox13.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox14.Text = row.Cells["使用品號"].Value.ToString();
                    textBox15.Text = row.Cells["使用品名"].Value.ToString();
                    textBoxID3.Text = row.Cells["ID"].Value.ToString();                  

                }
                else
                {
                    textBoxID3.Text = null;
                    textBox13.Text = null;
                    textBox14.Text = null;
                    textBox15.Text = null;
                }
            }
        }

        public void DELCALCOSTPRODS1RAW(string MID                                     
                                     , string MB001
                                     
                                     )
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
                                   
                                    DELETE  [TKRESEARCH].[dbo].[CALCOSTPRODS1RAW]
                                    WHERE [MID]='{0}' AND [MB001]='{1}'
                                    ", MID                                       
                                        , MB001
                                       );

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

        public void SEARCHDG3(string PRODNAMES)
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

                if (!string.IsNullOrEmpty(PRODNAMES))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [MID] AS 'ID'
                                        ,[PRODNAMES] AS '產品名稱'
                                        ,[MB001] AS '使用物料品號'
                                        ,[MB002] AS '使用物料品名'
                                        ,[INS] AS '投入重量'
                                        ,[PRICES] AS '單價'
                                        ,[TMONEYS] AS '金額'
                                        ,[REMARK] AS '備註'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS2MATERIL]
                                        WHERE [PRODNAMES] LIKE '%{0}%'
                                        ORDER BY [MID],[PRODNAMES],[MB001]

                                        ", PRODNAMES);
                }


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

        public void ADDCALCOSTPRODS2MATERIL(string MID
                                     , string PRODNAMES
                                     , string MB001
                                     , string MB002
                                     , string INS
                                     , string PRICES
                                     , string TMONEYS
                                     , string REMARK
                                     )
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
                                   
                                    INSERT INTO  [TKRESEARCH].[dbo].[CALCOSTPRODS2MATERIL]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MB001]
                                    ,[MB002]
                                    ,[INS]
                                    ,[PRICES]
                                    ,[TMONEYS]
                                    ,[REMARK]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    )
                                    ", MID
                                        , PRODNAMES
                                        , MB001
                                        , MB002
                                        , INS
                                        , PRICES
                                        , TMONEYS
                                        , REMARK);

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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBoxID5.Text = null;
            textBox25.Text = null;
            textBox26.Text = null;
            textBox27.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox25.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox26.Text = row.Cells["使用物料品號"].Value.ToString();
                    textBox27.Text = row.Cells["使用物料品名"].Value.ToString();
                    textBoxID5.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBoxID5.Text = null;
                    textBox25.Text = null;
                    textBox26.Text = null;
                    textBox27.Text = null;
                }
            }
        }

        public void DELCALCOSTPRODS2MATERIL(string MID
                                    , string MB001

                                    )
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
                                   
                                    DELETE  [TKRESEARCH].[dbo].[CALCOSTPRODS2MATERIL]
                                    WHERE [MID]='{0}' AND [MB001]='{1}'
                                    ", MID
                                        , MB001
                                       );

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
        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            textBox19.Text = FINDMB002(textBox18.Text.ToString());
            textBox21.Text = FINDPRICES(textBox18.Text.ToString());
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            textBox22.Text = (Convert.ToDecimal(textBox20.Text) * Convert.ToDecimal(textBox21.Text)).ToString();
        }

        public void SEARCHDG4(string PRODNAMES)
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

                if (!string.IsNullOrEmpty(PRODNAMES))
                {
                    sbSql.AppendFormat(@"  
                                       SELECT 
                                        [MID] AS 'ID'
                                        ,[PRODNAMES] AS '產品名稱'
                                        ,[MB001] AS '使用鹹蛋黃品號'
                                        ,[MB002] AS '使用鹹蛋黃品名'
                                        ,[INS] AS '投入重量'
                                        ,[NOTYIELD] AS '秏損率'
                                        ,[AFTERYIELDINS] AS '投入重量秏損後'
                                        ,[PRICES] AS '單價'
                                        ,[TMONEYS] AS '金額'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS3EGG]
                                        WHERE [PRODNAMES] LIKE '%{0}%'
                                        ORDER BY [MID],[PRODNAMES],[MB001]

                                        ", PRODNAMES);
                }


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds.Tables["ds"];

                        dataGridView4.AutoResizeColumns();


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

        public void ADDCALCOSTPRODS3EGG(string MID
                                   , string PRODNAMES
                                   , string MB001
                                   , string MB002
                                   , string INS
                                   , string NOTYIELD
                                   , string AFTERYIELDINS
                                   , string PRICES
                                   , string TMONEYS
                                   
                                   )
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
                                   
                                    INSERT INTO  [TKRESEARCH].[dbo].[CALCOSTPRODS3EGG]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MB001]
                                    ,[MB002]
                                    ,[INS]
                                    ,[NOTYIELD]
                                    ,[AFTERYIELDINS]
                                    ,[PRICES]
                                    ,[TMONEYS]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    )
                                    ", MID
                                        , PRODNAMES
                                        , MB001
                                        , MB002
                                        , INS
                                        , NOTYIELD
                                        , AFTERYIELDINS
                                        , PRICES
                                        , TMONEYS
                                        );

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

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBoxID7.Text = null;
            textBox36.Text = null;
            textBox37.Text = null;
            textBox38.Text = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox36.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox37.Text = row.Cells["使用鹹蛋黃品號"].Value.ToString();
                    textBox38.Text = row.Cells["使用鹹蛋黃品名"].Value.ToString();
                    textBoxID7.Text = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    textBoxID7.Text = null;
                    textBox36.Text = null;
                    textBox37.Text = null;
                    textBox38.Text = null;
                }
            }
        }
        public void DELCALCOSTPRODS3EGG(string MID
                                   , string MB001

                                   )
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
                                   
                                    DELETE  [TKRESEARCH].[dbo].[CALCOSTPRODS3EGG]
                                    WHERE [MID]='{0}' AND [MB001]='{1}'
                                    ", MID
                                        , MB001
                                       );

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
        private void textBox29_TextChanged(object sender, EventArgs e)
        {
            textBox30.Text = FINDMB002(textBox29.Text.ToString());
            textBox34.Text = FINDPRICES(textBox29.Text.ToString());
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox32.Text))
            {
                textBox33.Text = (Convert.ToDecimal(textBox31.Text) - (Convert.ToDecimal(textBox31.Text) * (Convert.ToDecimal(textBox32.Text) / 100))).ToString();
            }

            if (!string.IsNullOrEmpty(textBox34.Text))
            {
                textBox35.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox33.Text)).ToString();
            }
            
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox32.Text))
            {
                textBox33.Text = (Convert.ToDecimal(textBox31.Text) - (Convert.ToDecimal(textBox31.Text) * (Convert.ToDecimal(textBox32.Text) / 100))).ToString();
            }

            if (!string.IsNullOrEmpty(textBox34.Text))
            {
                textBox35.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox33.Text)).ToString();
            }
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox34.Text))
            {
                textBox35.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox33.Text)).ToString();
            }
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox34.Text))
            {
                textBox35.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox33.Text)).ToString();
            }
        }
        public void SETTEXTBOX1()
        {
            textBox3.Text = null;
            textBox4.Text = null;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox999.Text.Trim(),comboBox1.Text.ToString());
        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS(textBox3.Text, textBox4.Text);
            SETTEXTBOX1();

            SEARCH(textBox999.Text.Trim(), comboBox1.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UPDATECALCOSTPRODS(textBox1.Text, textBox2.Text,textBoxID.Text,comboBox2.Text);

            SEARCH(textBox999.Text.Trim(), comboBox1.Text.ToString());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHDG2(textBox5.Text.Trim());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS1RAW(textBoxID2.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim(), textBox8.Text.Trim(), textBox9.Text.Trim(), textBox10.Text.Trim(), textBox11.Text.Trim(), textBox12.Text.Trim());
            SEARCHDG2(textBox5.Text.Trim());
        }

        private void button6_Click(object sender, EventArgs e)
        {          
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCALCOSTPRODS1RAW(textBoxID3.Text.Trim(), textBox14.Text.Trim());
                SEARCHDG2(textBox5.Text.Trim());


            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SEARCHDG3(textBox16.Text); 
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS2MATERIL(textBoxID4.Text.Trim(), textBox17.Text.Trim(), textBox18.Text.Trim(), textBox19.Text.Trim(), textBox20.Text.Trim(), textBox21.Text.Trim(), textBox22.Text.Trim(), textBox23.Text.Trim());
            SEARCHDG3(textBox16.Text.Trim());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCALCOSTPRODS2MATERIL(textBoxID5.Text.Trim(), textBox26.Text.Trim());
                SEARCHDG3(textBox16.Text.Trim());


            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            SEARCHDG4(textBox24.Text);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS3EGG(textBoxID6.Text.Trim(), textBox28.Text.Trim(), textBox29.Text.Trim(), textBox30.Text.Trim(), textBox31.Text.Trim(), textBox32.Text.Trim(), textBox33.Text.Trim(), textBox34.Text.Trim(), textBox35.Text.Trim());
            SEARCHDG4(textBox24.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCALCOSTPRODS3EGG(textBoxID7.Text.Trim(), textBox37.Text.Trim());
                SEARCHDG4(textBox24.Text);


            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }




        #endregion

       
    }
}
