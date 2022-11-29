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

            SETCAL();
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

        public void SETCAL()
        {
            textBox41.Text=FINDDEVALUESWATER();
            textBox42.Text = FINDDEVALUESNG();

            textBox57.Text= FINDDEVALUESHUMAN();
            textBox71.Text = FINDDEVALUESHUMAN();
            textBox83.Text = FINDDEVALUESHUMAN();
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
                                        ,[TOTALRAWS] AS '總投入原料重量'
                                        ,[WATERLOSTS] AS '水份蒸發損秏率'
                                        ,[NGLOSTS] AS '不良率'
                                        ,[AFETERRAWS] AS '烘焙後總重量'
                                        ,[UNITCOSTS] AS '每單位原料成本'
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

            textBox39.Text = null;
            textBoxID8.Text = null;

            textBox61.Text = null;
            textBoxID9.Text = null;

            textBox62.Text = null;
            textBoxID10.Text = null;

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

                    textBox39.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID8.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG5(textBox39.Text);
                    SEARCHDG6(textBox39.Text);

                    textBox61.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID9.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG7(textBox61.Text);

                    textBox62.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID10.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG8(textBox62.Text);

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

                    textBox39.Text = null;
                    textBoxID8.Text = null;

                    textBox61.Text = null;
                    textBoxID9.Text = null;

                    textBox62.Text = null;
                    textBoxID10.Text = null;

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

        public void SEARCHDG5(string PRODNAMES )
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
                                        ,[TOTALRAWS] AS '總投入原料重量'
                                        ,[WATERLOSTS] AS '水份蒸發損秏率'
                                        ,[NGLOSTS] AS '不良率'
                                        ,[AFETERRAWS] AS '烘焙後總重量'
                                        ,[UNITCOSTS] AS '每單位原料成本'
                                        ,CONVERT(NVARCHAR,[CREATEDATE],112) AS '建立日期'
                                        ,[ISCLOESED] AS '結案'
                                        ,[ID]

                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                        WHERE [PRODNAMES] LIKE '%{0}%'                               
                                        ORDER BY CONVERT(NVARCHAR,[CREATEDATE],112)
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
                    dataGridView5.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds.Tables["ds"];

                        dataGridView5.AutoResizeColumns();


                    }

                }

            }
            catch
            {

            }
            finally
            {

            }

            ///計算原料、物料、鹹蛋黃的金額加總
            CALUNITCOSTSTMONEYS();
            SETUNITCOSTS();

        }
        public void UPDATECALCOSTPRODS(string ID
                                    , string TOTALRAWS
                                    , string WATERLOSTS
                                    , string NGLOSTS
                                    , string AFETERRAWS
                                    , string UNITCOSTS

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
                                   UPDATE [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                    SET TOTALRAWS='{1}',WATERLOSTS='{2}',NGLOSTS='{3}',AFETERRAWS='{4}',UNITCOSTS='{5}'
                                    WHERE [ID]='{0}'
                                  
                                      "
                                    , ID
                                    , TOTALRAWS
                                    , WATERLOSTS
                                    , NGLOSTS
                                    , AFETERRAWS
                                    , UNITCOSTS);

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

        public void SETUNITCOSTS()
        {
            decimal TOTALRAWS = 0;
            decimal WATERLOSTS = 0;
            decimal NGLOSTS = 0;
            decimal AFETERRAWS = 0;
            decimal UNITCOSTS = 0;

            textBox44.Text = "0";

            if (!string.IsNullOrEmpty(textBox40.Text)& !string.IsNullOrEmpty(textBox41.Text)& !string.IsNullOrEmpty(textBox42.Text)& !string.IsNullOrEmpty(textBox43.Text)& !string.IsNullOrEmpty(textBox44.Text))
            {
                TOTALRAWS = Convert.ToDecimal(textBox40.Text)+ Convert.ToDecimal(textBox48.Text);
                WATERLOSTS = Convert.ToDecimal(textBox41.Text);
                NGLOSTS = Convert.ToDecimal(textBox42.Text);
                AFETERRAWS = Convert.ToDecimal(textBox43.Text);
                UNITCOSTS = Convert.ToDecimal(textBox44.Text);

                AFETERRAWS = TOTALRAWS * (1 - WATERLOSTS / 100 - NGLOSTS / 100);
                textBox43.Text = AFETERRAWS.ToString();

                if(AFETERRAWS>0)
                {
                    UNITCOSTS = (Convert.ToDecimal(textBox45.Text) + Convert.ToDecimal(textBox46.Text) + Convert.ToDecimal(textBox47.Text)) / (AFETERRAWS*1000);
                    textBox44.Text = Math.Round(UNITCOSTS,3).ToString();
                }
               
            }
        }

        public string CALCOSTPRODS1RAWTMONEYS(string MID)
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

                if (!string.IsNullOrEmpty(MID))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT ISNULL(SUM([TMONEYS]),0) TMONEYS     
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS1RAW]
                                        WHERE MID='{0}'
                                        ", MID);
                }



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["TMONEYS"].ToString();
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

        public string CALCOSTPRODS1RAWTINS(string MID)
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

                if (!string.IsNullOrEmpty(MID))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT ISNULL(SUM([INS]),0) INS     
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS1RAW]
                                        WHERE MID='{0}'
                                        ", MID);
                }



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["INS"].ToString();
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
        public string CALCALCOSTPRODS2MATERILTMONEYS(string MID)
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

                if (!string.IsNullOrEmpty(MID))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT ISNULL(SUM([TMONEYS]),0) TMONEYS     
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS2MATERIL]
                                        WHERE MID='{0}'
                                        ", MID);
                }



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["TMONEYS"].ToString();
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
        public string CALCALCOSTPRODS3EGGTMONEYS(string MID)
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

                if (!string.IsNullOrEmpty(MID))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT ISNULL(SUM([TMONEYS]),0) TMONEYS     
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS3EGG]
                                        WHERE MID='{0}'
                                        ", MID);
                }



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["TMONEYS"].ToString();
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

        public string CALCALCOSTPRODS3EGGTINS(string MID)
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

                if (!string.IsNullOrEmpty(MID))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT ISNULL(SUM([INS]),0) INS     
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS3EGG]
                                        WHERE MID='{0}'
                                        ", MID);
                }



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["INS"].ToString();
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

        public void CALUNITCOSTSTMONEYS()
        {
            textBox45.Text = CALCOSTPRODS1RAWTMONEYS(textBoxID8.Text);
            textBox46.Text = CALCALCOSTPRODS2MATERILTMONEYS(textBoxID8.Text);
            textBox47.Text = CALCALCOSTPRODS3EGGTMONEYS(textBoxID8.Text);
            textBox40.Text = CALCOSTPRODS1RAWTINS(textBoxID8.Text);
            textBox48.Text = CALCALCOSTPRODS3EGGTINS(textBoxID8.Text);
        }

        public string FINDDEVALUESWATER()
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
                                       SELECT [KINDS]
                                        ,[TYPE]
                                        ,[DEVALUES]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='1' AND [TYPE]='水份蒸發率'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["DEVALUES"].ToString();
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
        public string FINDDEVALUESNG()
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
                                       SELECT [KINDS]
                                        ,[TYPE]
                                        ,[DEVALUES]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='1' AND [TYPE]='不良率'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["DEVALUES"].ToString();
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

        public string FINDDEVALUESHUMAN()
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
                                       SELECT [KINDS]
                                        ,[TYPE]
                                        ,[DEVALUES]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='1' AND [TYPE]='標準工時人'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["DEVALUES"].ToString();
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
        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            SETUNITCOSTS();
        }

        private void textBox42_TextChanged(object sender, EventArgs e)
        {
            SETUNITCOSTS();
        }

        public void SEARCHDG6(string PRODNAMES)
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
                                        ,[MANUHR1] AS '製程時間(時)-1'
                                        ,[MANUHUMAN1] AS '製程時間(時)-1所需人員(人)'
                                        ,[MANUHR2] AS '製程時間(時)-2'
                                        ,[MANUHUMAN2] AS '製程時間(時)-2所需人員(人)'
                                        ,[MANUHR3] AS '製程時間(時)-3'
                                        ,[MANUHUMAN3] AS '製程時間(時)-3所需人員(人)'
                                        ,[TMANUHR] AS '合計-製程時間(時)'
                                        ,[TMANUBUMAN] AS '合計-所需人員'
                                        ,[HRSEPCS] AS '標準工時/人'
                                        ,[THUMANCOSTS] AS '人工總計'
                                        ,[OUTSPEC] AS '標準產出數量'
                                        ,[HUMANCOSTS] AS '每單位製造人工成本'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS4PRODYIELD]
                                        WHERE [PRODNAMES]='{0}'

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
                    dataGridView6.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView6.DataSource = ds.Tables["ds"];

                        dataGridView6.AutoResizeColumns();


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

        public void CALHUMANCALCOSTPRODS4PRODYIELD()
        {
            if(!string.IsNullOrEmpty(textBox49.Text)& !string.IsNullOrEmpty(textBox50.Text)& !string.IsNullOrEmpty(textBox51.Text)& !string.IsNullOrEmpty(textBox52.Text)& !string.IsNullOrEmpty(textBox53.Text)& !string.IsNullOrEmpty(textBox54.Text))
            {
                textBox56.Text = (Convert.ToInt32(textBox50.Text) + Convert.ToInt32(textBox52.Text) + Convert.ToInt32(textBox54.Text)).ToString();
                textBox55.Text = (Convert.ToInt32(textBox49.Text) + Convert.ToInt32(textBox51.Text) + Convert.ToInt32(textBox53.Text)).ToString();
                int HUMAN = (Convert.ToInt32(textBox49.Text) * Convert.ToInt32(textBox50.Text) + Convert.ToInt32(textBox51.Text) * Convert.ToInt32(textBox52.Text) + Convert.ToInt32(textBox53.Text) * Convert.ToInt32(textBox54.Text));
                textBox58.Text = (Convert.ToDecimal(textBox57.Text) * HUMAN).ToString();
                //textBox58.Text = (Convert.ToInt32(textBox57.Text) * (Convert.ToInt32(textBox49.Text)* Convert.ToInt32(textBox50.Text)+ Convert.ToInt32(textBox51.Text)* Convert.ToInt32(textBox52.Text)+ Convert.ToInt32(textBox53.Text)* Convert.ToInt32(textBox54.Text))).ToString();
            }
        }
        private void textBox49_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS4PRODYIELD();
        }

        public void ADDCALCOSTPRODS4PRODYIELD(string MID
                                    , string PRODNAMES
                                    , string MANUHR1
                                    , string MANUHUMAN1
                                    , string MANUHR2
                                    , string MANUHUMAN2
                                    , string MANUHR3
                                    , string MANUHUMAN3
                                    , string TMANUHR
                                    , string TMANUBUMAN
                                    , string HRSEPCS
                                    , string THUMANCOSTS
                                    , string OUTSPEC
                                    , string HUMANCOSTS

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
                                    DELETE 
                                    [TKRESEARCH].[dbo].[CALCOSTPRODS4PRODYIELD]
                                    WHERE [MID]='{0}'

                                    INSERT INTO [TKRESEARCH].[dbo].[CALCOSTPRODS4PRODYIELD]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MANUHR1]
                                    ,[MANUHUMAN1]
                                    ,[MANUHR2]
                                    ,[MANUHUMAN2]
                                    ,[MANUHR3]
                                    ,[MANUHUMAN3]
                                    ,[TMANUHR]
                                    ,[TMANUBUMAN]
                                    ,[HRSEPCS]
                                    ,[THUMANCOSTS]
                                    ,[OUTSPEC]
                                    ,[HUMANCOSTS]
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
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    )
                                      "
                                    , MID
                                    , PRODNAMES
                                    , MANUHR1
                                    , MANUHUMAN1
                                    , MANUHR2
                                    , MANUHUMAN2
                                    , MANUHR3
                                    , MANUHUMAN3
                                    , TMANUHR
                                    , TMANUBUMAN
                                    , HRSEPCS
                                    , THUMANCOSTS
                                    , OUTSPEC
                                    , HUMANCOSTS
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
        public void SEARCHDG7(string PRODNAMES)
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
                                        ,[MANUHR1] AS '製程時間(時)-1'
                                        ,[MANUHUMAN1] AS '製程時間(時)-1所需人員(人)'
                                        ,[MANUHR2] AS '製程時間(時)-2'
                                        ,[MANUHUMAN2] AS '製程時間(時)-2所需人員(人)'
                                        ,[MANUHR3] AS '製程時間(時)-3'
                                        ,[MANUHUMAN3] AS '製程時間(時)-3所需人員(人)'
                                        ,[TMANUHR] AS '合計-製程時間(時)'
                                        ,[TMANUBUMAN] AS '合計-所需人員'
                                        ,[HRSEPCS] AS '標準工時/人'
                                        ,[THUMANCOSTS] AS '人工總計'
                                        ,[OUTSPEC] AS '標準產出數量'
                                        ,[HUMANCOSTS] AS '每單位製造人工成本'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS5MANU]
                                        WHERE [PRODNAMES]='{0}'

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
                    dataGridView7.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView7.DataSource = ds.Tables["ds"];

                        dataGridView7.AutoResizeColumns();


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

        public void CALHUMANCALCOSTPRODS5MANU()
        {
            if (!string.IsNullOrEmpty(textBox63.Text) & !string.IsNullOrEmpty(textBox64.Text) & !string.IsNullOrEmpty(textBox65.Text) & !string.IsNullOrEmpty(textBox66.Text) & !string.IsNullOrEmpty(textBox67.Text) & !string.IsNullOrEmpty(textBox68.Text))
            {
                textBox69.Text = (Convert.ToInt32(textBox63.Text) + Convert.ToInt32(textBox65.Text) + Convert.ToInt32(textBox67.Text)).ToString();
                textBox70.Text = (Convert.ToInt32(textBox64.Text) + Convert.ToInt32(textBox66.Text) + Convert.ToInt32(textBox68.Text)).ToString();
                int HUMAN = (Convert.ToInt32(textBox63.Text) * Convert.ToInt32(textBox64.Text) + Convert.ToInt32(textBox65.Text) * Convert.ToInt32(textBox66.Text) + Convert.ToInt32(textBox67.Text) * Convert.ToInt32(textBox68.Text));
                textBox72.Text = (Convert.ToDecimal(textBox71.Text) * HUMAN).ToString();

                if(Convert.ToDecimal(textBox72.Text)>0 & Convert.ToDecimal(textBox73.Text)>0)
                {
                    textBox74.Text = (Convert.ToDecimal(textBox72.Text) / Convert.ToDecimal(textBox73.Text)).ToString();
                }
                
                //textBox58.Text = (Convert.ToInt32(textBox57.Text) * (Convert.ToInt32(textBox49.Text)* Convert.ToInt32(textBox50.Text)+ Convert.ToInt32(textBox51.Text)* Convert.ToInt32(textBox52.Text)+ Convert.ToInt32(textBox53.Text)* Convert.ToInt32(textBox54.Text))).ToString();
            }
        }

        private void textBox63_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox64_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox65_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox66_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox67_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox68_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }
        private void textBox73_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }

        private void textBox72_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS5MANU();
        }


        public void ADDCALCOSTPRODS5MANU(string MID
                                    , string PRODNAMES
                                    , string MANUHR1
                                    , string MANUHUMAN1
                                    , string MANUHR2
                                    , string MANUHUMAN2
                                    , string MANUHR3
                                    , string MANUHUMAN3
                                    , string TMANUHR
                                    , string TMANUBUMAN
                                    , string HRSEPCS
                                    , string THUMANCOSTS
                                    , string OUTSPEC
                                    , string HUMANCOSTS

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
                                    DELETE 
                                    [TKRESEARCH].[dbo].[CALCOSTPRODS5MANU]
                                    WHERE [MID]='{0}'

                                    INSERT INTO [TKRESEARCH].[dbo].[CALCOSTPRODS5MANU]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MANUHR1]
                                    ,[MANUHUMAN1]
                                    ,[MANUHR2]
                                    ,[MANUHUMAN2]
                                    ,[MANUHR3]
                                    ,[MANUHUMAN3]
                                    ,[TMANUHR]
                                    ,[TMANUBUMAN]
                                    ,[HRSEPCS]
                                    ,[THUMANCOSTS]
                                    ,[OUTSPEC]
                                    ,[HUMANCOSTS]
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
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    )
                                      "
                                    , MID
                                    , PRODNAMES
                                    , MANUHR1
                                    , MANUHUMAN1
                                    , MANUHR2
                                    , MANUHUMAN2
                                    , MANUHR3
                                    , MANUHUMAN3
                                    , TMANUHR
                                    , TMANUBUMAN
                                    , HRSEPCS
                                    , THUMANCOSTS
                                    , OUTSPEC
                                    , HUMANCOSTS
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

        public void SEARCHDG8(string PRODNAMES)
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
                                        ,[MANUHR1] AS '製程時間(時)-1'
                                        ,[MANUHUMAN1] AS '製程時間(時)-1所需人員(人)'
                                        ,[MANUHR2] AS '製程時間(時)-2'
                                        ,[MANUHUMAN2] AS '製程時間(時)-2所需人員(人)'
                                        ,[MANUHR3] AS '製程時間(時)-3'
                                        ,[MANUHUMAN3] AS '製程時間(時)-3所需人員(人)'
                                        ,[TMANUHR] AS '合計-製程時間(時)'
                                        ,[TMANUBUMAN] AS '合計-所需人員'
                                        ,[HRSEPCS] AS '標準工時/人'
                                        ,[THUMANCOSTS] AS '人工總計'
                                        ,[OUTSPEC] AS '標準產出數量'
                                        ,[HUMANCOSTS] AS '每單位製造人工成本'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS6INPACK]
                                        WHERE [PRODNAMES]='{0}'

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
                    dataGridView8.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView8.DataSource = ds.Tables["ds"];

                        dataGridView8.AutoResizeColumns();


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

        public void CALHUMANCALCOSTPRODS6INPACK()
        {
            if (!string.IsNullOrEmpty(textBox75.Text) & !string.IsNullOrEmpty(textBox76.Text) & !string.IsNullOrEmpty(textBox77.Text) & !string.IsNullOrEmpty(textBox78.Text) & !string.IsNullOrEmpty(textBox79.Text) & !string.IsNullOrEmpty(textBox80.Text))
            {
                textBox81.Text = (Convert.ToInt32(textBox75.Text) + Convert.ToInt32(textBox77.Text) + Convert.ToInt32(textBox79.Text)).ToString();
                textBox82.Text = (Convert.ToInt32(textBox76.Text) + Convert.ToInt32(textBox78.Text) + Convert.ToInt32(textBox80.Text)).ToString();
                int HUMAN = (Convert.ToInt32(textBox75.Text) * Convert.ToInt32(textBox76.Text) + Convert.ToInt32(textBox77.Text) * Convert.ToInt32(textBox78.Text) + Convert.ToInt32(textBox79.Text) * Convert.ToInt32(textBox80.Text));
                textBox84.Text = (Convert.ToDecimal(textBox83.Text) * HUMAN).ToString();

                if (Convert.ToDecimal(textBox84.Text) > 0 & Convert.ToDecimal(textBox85.Text) > 0)
                {
                    textBox86.Text = (Convert.ToDecimal(textBox84.Text) / Convert.ToDecimal(textBox85.Text)).ToString();
                }

                //textBox58.Text = (Convert.ToInt32(textBox57.Text) * (Convert.ToInt32(textBox49.Text)* Convert.ToInt32(textBox50.Text)+ Convert.ToInt32(textBox51.Text)* Convert.ToInt32(textBox52.Text)+ Convert.ToInt32(textBox53.Text)* Convert.ToInt32(textBox54.Text))).ToString();
            }
        }

        private void textBox75_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }

        private void textBox76_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }

        private void textBox77_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }

        private void textBox78_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }

        private void textBox79_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }

        private void textBox80_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }
        private void textBox85_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS6INPACK();
        }
        public void ADDCALCOSTPRODS6INPACK(string MID
                                   , string PRODNAMES
                                   , string MANUHR1
                                   , string MANUHUMAN1
                                   , string MANUHR2
                                   , string MANUHUMAN2
                                   , string MANUHR3
                                   , string MANUHUMAN3
                                   , string TMANUHR
                                   , string TMANUBUMAN
                                   , string HRSEPCS
                                   , string THUMANCOSTS
                                   , string OUTSPEC
                                   , string HUMANCOSTS

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
                                    DELETE 
                                    [TKRESEARCH].[dbo].[CALCOSTPRODS6INPACK]
                                    WHERE [MID]='{0}'

                                    INSERT INTO [TKRESEARCH].[dbo].[CALCOSTPRODS6INPACK]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[MANUHR1]
                                    ,[MANUHUMAN1]
                                    ,[MANUHR2]
                                    ,[MANUHUMAN2]
                                    ,[MANUHR3]
                                    ,[MANUHUMAN3]
                                    ,[TMANUHR]
                                    ,[TMANUBUMAN]
                                    ,[HRSEPCS]
                                    ,[THUMANCOSTS]
                                    ,[OUTSPEC]
                                    ,[HUMANCOSTS]
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
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    )
                                      "
                                    , MID
                                    , PRODNAMES
                                    , MANUHR1
                                    , MANUHUMAN1
                                    , MANUHR2
                                    , MANUHUMAN2
                                    , MANUHR3
                                    , MANUHUMAN3
                                    , TMANUHR
                                    , TMANUBUMAN
                                    , HRSEPCS
                                    , THUMANCOSTS
                                    , OUTSPEC
                                    , HUMANCOSTS
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

        private void button13_Click(object sender, EventArgs e)
        {
            SEARCHDG5(textBox39.Text);
            SEARCHDG6(textBox39.Text);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            UPDATECALCOSTPRODS(textBoxID8.Text, textBox40.Text, textBox41.Text, textBox42.Text, textBox43.Text, textBox44.Text);
            SEARCHDG5(textBox39.Text);
            SEARCHDG6(textBox39.Text);
        }


        private void button15_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS4PRODYIELD(textBoxID8.Text, textBox39.Text,textBox49.Text,textBox50.Text, textBox51.Text, textBox52.Text, textBox53.Text, textBox54.Text, textBox55.Text, textBox56.Text, textBox57.Text, textBox58.Text, textBox59.Text, textBox60.Text);
            SEARCHDG5(textBox39.Text);
            SEARCHDG6(textBox39.Text);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            SEARCHDG7(textBox61.Text);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS5MANU(textBoxID9.Text, textBox61.Text, textBox63.Text, textBox64.Text, textBox65.Text, textBox66.Text, textBox67.Text, textBox68.Text, textBox69.Text, textBox70.Text, textBox71.Text, textBox72.Text, textBox73.Text, textBox74.Text);
            SEARCHDG7(textBox61.Text);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            SEARCHDG8(textBox62.Text);
        }

        private void button19_Click(object sender, EventArgs e)
        {

            ADDCALCOSTPRODS6INPACK(textBoxID10.Text, textBox62.Text, textBox75.Text, textBox76.Text, textBox77.Text, textBox78.Text, textBox79.Text, textBox80.Text, textBox81.Text, textBox82.Text, textBox83.Text, textBox84.Text, textBox85.Text, textBox86.Text);
            SEARCHDG8(textBox62.Text);
        }


        #endregion


    }
}
