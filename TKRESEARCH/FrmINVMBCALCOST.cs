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
            textBox96.Text = FINDDEVALUESHUMAN();
            textBox104.Text = FINDDEVALUESMANUHUMAN();
            textBox111.Text = FINDDEVALUESEGGS();
            textBox112.Text = FINDREMARKSEGGS();
            textBox113.Text = FINDREMARKS1();
            textBox114.Text = FINDREMARKS2();
            //textBox115.Text = FINDREMARKS3();
            textBox115.Text = FINDREMARKSOUTPACK().ToString();
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

            textBox87.Text = null;
            textBoxID11.Text = null;

            textBox100.Text = null;
            textBox101.Text = null;
            textBoxID12.Text = null;

            textBox116.Text = null;
            textBoxID13.Text = null;

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

                    textBox87.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID11.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG9(textBox87.Text);

                    textBox100.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox101.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID12.Text = row.Cells["ID"].Value.ToString();
                    SEARCHDG10(textBox100.Text);

                    textBox116.Text = row.Cells["產品名稱"].Value.ToString();
                    textBoxID13.Text = row.Cells["ID"].Value.ToString();

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

                    textBox87.Text = null;
                    textBoxID11.Text = null;

                    textBox100.Text = null;
                    textBox101.Text = null;
                    textBoxID12.Text = null;

                    textBox116.Text = null;
                    textBoxID13.Text = null;

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
            textBox9.Text = "0";
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrEmpty(textBox10.Text))
            {
                textBox11.Text = (Convert.ToDecimal(textBox9.Text) * Convert.ToDecimal(textBox10.Text)).ToString();
            }
            
        }
        public string FINDMB001(string MB002)
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

                if (!string.IsNullOrEmpty(MB002))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT MB001,MB002 
                                        FROM [TK].dbo.INVMB
                                        WHERE (MB001 LIKE '1%' OR MB001 LIKE '2%')
                                        AND MB002 LIKE '%{0}%'

                                        ", MB002);
                }


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["MB001"].ToString();

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

            if (dataGridView2.CurrentRow != null)
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
                                     , string MB002
                                     
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
                                    WHERE [MID]='{0}' AND [MB002]='{1}'
                                    ", MID                                       
                                        , MB002
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
                                    , string MB002

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
                                    WHERE [MID]='{0}' AND [MB002]='{1}'
                                    ", MID
                                        , MB002
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
            textBox20.Text = "0";
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox20.Text) && !string.IsNullOrEmpty(textBox21.Text))
            {
                textBox22.Text = (Convert.ToDecimal(textBox20.Text) * Convert.ToDecimal(textBox21.Text)).ToString();
            }
            
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
            textBox110.Text = CALCALCOSTPRODS3EGGTINS(textBoxID8.Text);
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

        public string FINDDEVALUESMANUHUMAN()
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
                                        WHERE [KINDS]='1' AND [TYPE]='標準製費'
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

        public string FINDDEVALUESEGGS()
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
                                        WHERE [KINDS]='1' AND [TYPE]='烤後鹹蛋黃一桶片數'
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

        public string FINDREMARKSEGGS()
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
                                        ,[REMARKS]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='B' AND [TYPE]='烤後鹹蛋黃'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["REMARKS"].ToString();
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

        public string FINDREMARKS1()
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
                                        ,[REMARKS]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='B' AND [TYPE]='製程'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["REMARKS"].ToString();
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
        public string FINDREMARKS2()
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
                                        ,[REMARKS]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='B' AND [TYPE]='內包'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["REMARKS"].ToString();
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
        public string FINDREMARKS3()
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
                                        ,[REMARKS]
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='B' AND [TYPE]='外包'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["REMARKS"].ToString();
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

         public StringBuilder FINDREMARKSOUTPACK()
        {
            StringBuilder MESS = new StringBuilder();

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
                                        ([TYPE]+':'+[REMARKS]) AS 'REMARKS'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODSTYPE]
                                        WHERE [KINDS]='A'
                                        ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    
                    //return ds.Tables["ds"].Rows[0]["REMARKS"].ToString();
                    foreach(DataRow DR in ds.Tables["ds"].Rows)
                    {
                        MESS.AppendFormat(@"{0}", DR["REMARKS"].ToString());
                        MESS.AppendLine();
                    }

                    return MESS;
                }
                else
                {
                    return MESS;
                }

            }
            catch
            {
                return MESS;
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

        public void SEARCHDG9(string PRODNAMES)
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
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS7OUTPACK]
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
                    dataGridView9.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView9.DataSource = ds.Tables["ds"];

                        dataGridView9.AutoResizeColumns();


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

        public void CALHUMANCALCOSTPRODS7OUTPACK()
        {
            if (!string.IsNullOrEmpty(textBox88.Text) & !string.IsNullOrEmpty(textBox89.Text) & !string.IsNullOrEmpty(textBox90.Text) & !string.IsNullOrEmpty(textBox91.Text) & !string.IsNullOrEmpty(textBox92.Text) & !string.IsNullOrEmpty(textBox93.Text))
            {
                textBox94.Text = (Convert.ToInt32(textBox88.Text) + Convert.ToInt32(textBox90.Text) + Convert.ToInt32(textBox92.Text)).ToString();
                textBox94.Text = (Convert.ToInt32(textBox89.Text) + Convert.ToInt32(textBox91.Text) + Convert.ToInt32(textBox93.Text)).ToString();
                int HUMAN = (Convert.ToInt32(textBox88.Text) * Convert.ToInt32(textBox89.Text) + Convert.ToInt32(textBox90.Text) * Convert.ToInt32(textBox91.Text) + Convert.ToInt32(textBox92.Text) * Convert.ToInt32(textBox93.Text));
                textBox97.Text = (Convert.ToDecimal(textBox96.Text) * HUMAN).ToString();

                if (Convert.ToDecimal(textBox97.Text) > 0 & Convert.ToDecimal(textBox98.Text) > 0)
                {
                    textBox99.Text = (Convert.ToDecimal(textBox97.Text) / Convert.ToDecimal(textBox98.Text)).ToString();
                }

                //textBox58.Text = (Convert.ToInt32(textBox57.Text) * (Convert.ToInt32(textBox49.Text)* Convert.ToInt32(textBox50.Text)+ Convert.ToInt32(textBox51.Text)* Convert.ToInt32(textBox52.Text)+ Convert.ToInt32(textBox53.Text)* Convert.ToInt32(textBox54.Text))).ToString();
            }
        }

        public void ADDCALCOSTPRODS7OUTPACK(string MID
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
                                    [TKRESEARCH].[dbo].[CALCOSTPRODS7OUTPACK]
                                    WHERE [MID]='{0}'

                                    INSERT INTO [TKRESEARCH].[dbo].[CALCOSTPRODS7OUTPACK]
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

        private void textBox88_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox89_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox90_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox91_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox92_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox93_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        private void textBox98_TextChanged(object sender, EventArgs e)
        {
            CALHUMANCALCOSTPRODS7OUTPACK();
        }

        public void SEARCHDG10(string PRODNAMES)
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
                                        [MID]  AS 'ID'
                                        ,[PRODNAMES] AS '產品名稱'
                                        ,[PRODWORKS] AS '生產站'
                                        ,[HRS] AS '每小時製費/人'
                                        ,[HUMANS] AS '所需人數'
                                        ,[TOTALHRHUMANS] AS '8小時製造費用金額'
                                        ,[PRODS] AS '8小時生產數量'
                                        ,[UNITCOSTS] AS '每單位製造費用'
                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS8PRODUSEDMANU]
                                        WHERE [PRODNAMES]='{0}'
                                        ORDER BY [PRODWORKS]

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
                    dataGridView9.DataSource = null;

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView10.DataSource = ds.Tables["ds"];

                        dataGridView10.AutoResizeColumns();


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
        
        public void CALCOSTPRODS8PRODUSEDMANU()
        {
            if(!string.IsNullOrEmpty(textBox104.Text)& !string.IsNullOrEmpty(textBox105.Text)& !string.IsNullOrEmpty(textBox106.Text)& !string.IsNullOrEmpty(textBox107.Text))
            {
                textBox106.Text = (Convert.ToDecimal(textBox104.Text)* Convert.ToDecimal(textBox105.Text)).ToString();

                if(Convert.ToDecimal(textBox104.Text)>0& Convert.ToDecimal(textBox105.Text)>0& Convert.ToDecimal(textBox107.Text)>0)
                {
                    textBox103.Text = ((Convert.ToDecimal(textBox104.Text) * Convert.ToDecimal(textBox105.Text)) / Convert.ToDecimal(textBox107.Text)).ToString();
                }
                
            }
        }
        private void textBox104_TextChanged(object sender, EventArgs e)
        {
            CALCOSTPRODS8PRODUSEDMANU();
        }

        private void textBox105_TextChanged(object sender, EventArgs e)
        {
            CALCOSTPRODS8PRODUSEDMANU();
        }

        private void textBox106_TextChanged(object sender, EventArgs e)
        {
            CALCOSTPRODS8PRODUSEDMANU();
        }

        private void textBox107_TextChanged(object sender, EventArgs e)
        {
            CALCOSTPRODS8PRODUSEDMANU();
        }
        public void ADDCALCOSTPRODS8PRODUSEDMANU(string MID
                                                , string PRODNAMES
                                                , string PRODWORKS
                                                , string HRS
                                                , string HUMANS
                                                , string TOTALHRHUMANS
                                                , string PRODS
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
                                    INSERT INTO [TKRESEARCH].[dbo].[CALCOSTPRODS8PRODUSEDMANU]
                                    (
                                    [MID]
                                    ,[PRODNAMES]
                                    ,[PRODWORKS]
                                    ,[HRS]
                                    ,[HUMANS]
                                    ,[TOTALHRHUMANS]
                                    ,[PRODS]
                                    ,[UNITCOSTS]
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
                                    , PRODWORKS
                                    , HRS
                                    , HUMANS
                                    , TOTALHRHUMANS
                                    , PRODS
                                    , UNITCOSTS
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
        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
           
            textBox108.Text = null;
            textBox109.Text = null;

            if (dataGridView10.CurrentRow != null)
            {
                int rowindex = dataGridView10.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView10.Rows[rowindex];
                    textBox108.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox109.Text = row.Cells["生產站"].Value.ToString();
                    
                    

                }
                else
                {
                    textBox108.Text = null;
                    textBox109.Text = null;
                }
            }
        }

        public void DELCALCOSTPRODS8PRODUSEDMANU(string MID                                               
                                                , string PRODWORKS
                                                
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
                                    DELETE  [TKRESEARCH].[dbo].[CALCOSTPRODS8PRODUSEDMANU]
                                    WHERE [MID]='{0}' AND [PRODWORKS]='{1}'
                                    
                                    ", MID                                    
                                    , PRODWORKS
                                  
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

        public void CALATFERCOOKEGGS()
        {
            if (!string.IsNullOrEmpty(textBox58.Text) & !string.IsNullOrEmpty(textBox59.Text) & !string.IsNullOrEmpty(textBox110.Text) & !string.IsNullOrEmpty(textBox111.Text))
            {
                if (Convert.ToDecimal(textBox58.Text) > 0 & Convert.ToDecimal(textBox59.Text) > 0 & Convert.ToDecimal(textBox110.Text) > 0 & Convert.ToDecimal(textBox111.Text) > 0)
                {
                    textBox60.Text = ((Convert.ToDecimal(textBox58.Text) / Convert.ToDecimal(textBox59.Text)) * Convert.ToDecimal(textBox110.Text)/ Convert.ToDecimal(textBox111.Text)).ToString();
                }

            }
        }
        private void textBox58_TextChanged(object sender, EventArgs e)
        {
            CALATFERCOOKEGGS();
        }

        private void textBox59_TextChanged(object sender, EventArgs e)
        {
            CALATFERCOOKEGGS();
        }

        private void textBox110_TextChanged(object sender, EventArgs e)
        {
            CALATFERCOOKEGGS();
        }

        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            CALATFERCOOKEGGS();
        }

        private void textBox102_TextChanged(object sender, EventArgs e)
        {
            textBox107.Text = "0";
            textBox108.Text = "0";
        }

        public void UPDATECALCOSTPRODSCOSTTOTALS(string ID
                                        , string COSTRAW
                                        , string COSTMATERIL
                                        , string COSTART
                                        , string COSTMANU
                                        , string COSTTOTALS
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
                                    SET [COSTRAW]='{1}',[COSTMATERIL]='{2}',[COSTART]='{3}',[COSTMANU]='{4}',[COSTTOTALS]='{5}'
                                    WHERE [ID]='{0}'
                                    ", ID
                                    , COSTRAW
                                    , COSTMATERIL
                                    , COSTART
                                    , COSTMANU
                                    , COSTTOTALS
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

        public void CALCALCOSTPRODS()
        {
            if (!string.IsNullOrEmpty(textBox117.Text) & !string.IsNullOrEmpty(textBox118.Text) & !string.IsNullOrEmpty(textBox119.Text) & !string.IsNullOrEmpty(textBox120.Text))
            {
                textBox121.Text = ((Convert.ToDecimal(textBox117.Text) + Convert.ToDecimal(textBox118.Text)) + Convert.ToDecimal(textBox119.Text) +Convert.ToDecimal(textBox120.Text)).ToString();

            }
        }
        private void textBox117_TextChanged(object sender, EventArgs e)
        {
            CALCALCOSTPRODS();
        }

        private void textBox118_TextChanged(object sender, EventArgs e)
        {
            CALCALCOSTPRODS();
        }

        private void textBox119_TextChanged(object sender, EventArgs e)
        {
            CALCALCOSTPRODS();
        }

        private void textBox120_TextChanged(object sender, EventArgs e)
        {
            CALCALCOSTPRODS();
        }
        public void SETTEXTBOX1()
        {
            textBox3.Text = null;
            textBox4.Text = null;
        }

        public void SETFASTREPORT(string ID)
        {
            StringBuilder SQL = new StringBuilder();
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();
            StringBuilder SQL4 = new StringBuilder();
            StringBuilder SQL5 = new StringBuilder();
            StringBuilder SQL6 = new StringBuilder();
            StringBuilder SQL7 = new StringBuilder();
            StringBuilder SQL8 = new StringBuilder();

            SQL = SETSQL(ID);
            SQL1 = SETSQL1(ID);
            SQL2 = SETSQL2(ID);
            SQL3 = SETSQL3(ID);
            SQL4 = SETSQL4(ID);
            SQL5 = SETSQL5(ID);
            SQL6 = SETSQL6(ID);
            SQL7 = SETSQL7(ID);
            SQL8 = SETSQL8(ID);

            Report report1 = new Report();
            report1.Load(@"REPORT\新品成本試算表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL1.ToString();
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL2.ToString();
            TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
            table3.SelectCommand = SQL3.ToString();
            TableDataSource table4 = report1.GetDataSource("Table4") as TableDataSource;
            table4.SelectCommand = SQL4.ToString();
            TableDataSource table5 = report1.GetDataSource("Table5") as TableDataSource;
            table5.SelectCommand = SQL5.ToString();
            TableDataSource table6 = report1.GetDataSource("Table6") as TableDataSource;
            table6.SelectCommand = SQL6.ToString();
            TableDataSource table7 = report1.GetDataSource("Table7") as TableDataSource;
            table7.SelectCommand = SQL7.ToString();
            TableDataSource table8 = report1.GetDataSource("Table8") as TableDataSource;
            table8.SelectCommand = SQL8.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();

        }
        public StringBuilder SETSQL(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
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
                            WHERE [ID]='{0}'

                           

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL1(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
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
                            WHERE [MID]='{0}'

                            ORDER BY [MID],[PRODNAMES],[MB001]

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL2(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
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
                            WHERE [MID]='{0}'
                            ORDER BY [MID],[PRODNAMES],[MB001]

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL3(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
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
                            WHERE [MID]='{0}'
                            ORDER BY [MID],[PRODNAMES],[MB001]

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL4(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                           
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
                            WHERE [MID]='{0}' 

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL5(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 

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
                            WHERE [MID]='{0}' 
                           

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL6(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                            

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
                            WHERE [MID]='{0}'   
                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL7(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                           
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
                            FROM [TKRESEARCH].[dbo].[CALCOSTPRODS7OUTPACK]
                            WHERE [MID]='{0}'   

                            ", ID);


            return SB;
        }
        public StringBuilder SETSQL8(string ID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                            [MID] AS 'ID'
                            ,[PRODNAMES] AS '產品名稱'
                            ,[PRODWORKS] AS '生產站'
                            ,[HRS] AS '每小時製費/人'
                            ,[HUMANS] AS '所需人數'
                            ,[TOTALHRHUMANS] AS '8小時製造費用金額'
                            ,[PRODS] AS '8小時生產數量'
                            ,[UNITCOSTS] AS '每單位製造費用'
                            FROM [TKRESEARCH].[dbo].[CALCOSTPRODS8PRODUSEDMANU]
                            WHERE [MID]='{0}' 
                            ORDER BY [PRODWORKS]

                            ", ID);


            return SB;
        }

        public void SETFASTREPORT2(string SDATE, string MB001, string MB003,string MB002)
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);




            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(SDATE, MB001, MB003, MB002);
            Report report1 = new Report();
            report1.Load(@"REPORT\平均成本.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string YM, string MB001, string MB003,string MB002)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(MB001))
            {
                SQUERY.AppendFormat(@"
                                    AND TA001 LIKE '%{0}%'
                                    ", MB001);
            }
            else
            {
                SQUERY.AppendFormat(@"

                                    ");
            }

            if (!string.IsNullOrEmpty(MB003))
            {
                SQUERY.AppendFormat(@"
                                     AND MB003 LIKE '%{0}%'
                                    ", MB003);
            }
            else
            {
                SQUERY.AppendFormat(@"

                                    ");
            }

            if (!string.IsNullOrEmpty(MB002))
            {
                SQUERY.AppendFormat(@"
                                     AND MB002 LIKE '%{0}%'
                                    ", MB002);
            }
            else
            {
                SQUERY.AppendFormat(@"

                                    ");
            }

            SB.AppendFormat(@" 

                            SELECT TA002 AS '年月',TA001 AS '品號',MB002 AS '品名',MB003 AS '規格',生產入庫數,ME005 在製約量_材料,本階人工成本,本階製造費用,ME007 材料成本,ME008 人工成本,ME009 製造費用,ME010 加工費用
                            ,((ME007+ME008+ME009+ME010)/(生產入庫數+ME005)) 單位成本, ((ME007)/(生產入庫數+ME005)) 單位材料成本, ((ME008)/(生產入庫數+ME005)) 單位人工成本,((ME009)/(生產入庫數+ME005)) 單位製造成本,((ME010)/(生產入庫數+ME005)) 單位加工成本
                            ,MB068
                            ,(CASE WHEN MB068 IN ('09') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均包裝人工成本
                            ,(CASE WHEN MB068 IN ('09') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均包裝製造費用
                            ,(CASE WHEN MB068 IN ('03') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均小線人工成本
                            ,(CASE WHEN MB068 IN ('03') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均小線製造費用
                            ,(CASE WHEN MB068 IN ('02') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均大線人工成本
                            ,(CASE WHEN MB068 IN ('02') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均大線製造費用
                            FROM 
                            (
                            SELECT TA002,TA001,SUM(TA012) '生產入庫數',SUM(TA016-TA019) AS '本階人工成本',SUM(TA017-TA020) AS '本階製造費用'
                            FROM [TK].dbo.CSTTA
                            WHERE TA002 LIKE '{0}%'
                            GROUP BY TA002,TA001
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                            WHERE 1=1

                           {1} 
                          
                            AND (生產入庫數+ME005)>0
                            ORDER BY TA001,TA002
                              
                            ", YM, SQUERY.ToString()); 

            return SB;

        }

        public void SETFASTREPORT3(string SDATE, string MB001, string MB002)
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);




            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(SDATE, MB001, MB002);
            Report report1 = new Report();
            report1.Load(@"REPORT\品號平均年度單位成本.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string YM, string MB001, string MB002)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(MB001))
            {
                SQUERY.AppendFormat(@"
                                    AND 成品品號 LIKE '%{0}%'
                                    ", MB001);
            }
            else
            {
                SQUERY.AppendFormat(@"

                                    ");
            }
            

            if (!string.IsNullOrEmpty(MB002))
            {
                SQUERY.AppendFormat(@"
                                     AND 成品品名 LIKE '%{0}%'
                                    ", MB002);
            }
            else
            {
                SQUERY.AppendFormat(@"

                                    ");
            }

            SB.AppendFormat(@" 

                            SELECT *
                            ,(CASE WHEN 總成品平均成本>0 THEN 分攤成本/總成品平均成本 ELSE 0 END) AS '各百分比' 
                            FROM 
                            (
                            SELECT '{0}'AS '年度',MC001 AS '成品品號',MB1.MB002  AS '成品品名' ,MC004,MD003  AS '使用品號',MB2.MB002  AS '使用品名',MD006,MD007
                            ,總成品平均成本
                            ,材料平均成本
                            ,人工平均成本
                            ,製造平均成本
                            ,加工平均成本
                            ,各採購單位成本
                            ,總採購單位成本
                            ,總半成品重
                            ,(CASE WHEN 總成品平均成本>0 THEN (CASE WHEN MB2.MB001 LIKE '3%' THEN ((材料平均成本-總採購單位成本)*MD006/MD007/總半成品重) ELSE 各採購單位成本 END) ELSE 0 END) AS '分攤成本' 
                            ,(CASE WHEN MD003 LIKE '1%' THEN '1原料'  WHEN MD003 LIKE '2%' THEN '2物料' WHEN MD003 LIKE '3%' THEN '3半成品'END ) AS '分類'
                            FROM
                            (
                            SELECT MC001,MC004,MD003,MD006,MD007
                            ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                            ,ISNULL((SELECT AVG((ME007)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND ME002 LIKE '{0}%'),0) AS '材料平均成本'
                            ,ISNULL((SELECT AVG((ME008)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND ME002 LIKE '{0}%'),0) AS '人工平均成本'
                            ,ISNULL((SELECT AVG((ME009)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND ME002 LIKE '{0}%'),0) AS '製造平均成本'
                            ,ISNULL((SELECT AVG((ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND ME002 LIKE '{0}%'),0) AS '加工平均成本'
                            ,(CASE WHEN ( MB2.MB001 LIKE '1%' OR MB2.MB001 LIKE '2%') AND MB2.MB064>0 AND MB2.MB065 >0 THEN MB2.MB065/MB2.MB064*MD006/MD007/MC004 ELSE MB2.MB050 END ) AS '各採購單位成本'
                            ,(SELECT SUM (CASE WHEN  ( MB001 LIKE '1%' OR MB001 LIKE '2%') AND MB064>0 AND MB065 >0 THEN MB065/MB064*MD006/MD007/MC004 ELSE MB050 END) FROM [TK].dbo.BOMMC MC, [TK].dbo.BOMMD MD ,[TK].dbo.INVMB MB WHERE  MC.MC001=MD.MD001 AND MD.MD003=MB.MB001 AND MC.MC001=BOMMC.MC001)   AS '總採購單位成本'
                            ,ISNULL((SELECT SUM (MD006/MD007) FROM [TK].dbo.BOMMC MC, [TK].dbo.BOMMD MD ,[TK].dbo.INVMB MB WHERE  MC.MC001=MD.MD001 AND MD.MD003=MB.MB001 AND MC.MC001=BOMMC.MC001 AND MB.MB001 LIKE '3%'),0)  AS '總半成品重'
                            FROM [TK].dbo.BOMMC
                            LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=BOMMC.MC001
                            , [TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=BOMMD.MD003
                            WHERE MC001=MD001
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=TEMP.MC001
                            LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=TEMP.MD003
                            UNION ALL
                            SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                            ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,ISNULL((SELECT AVG((ME008)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '成本'
                            ,'4人工' AS '分類'
                            FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                            WHERE  MC001=MB001
                            AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 
                            UNION ALL
                            SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                            ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,ISNULL((SELECT AVG((ME009)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '成本'
                            ,'5製造' AS '分類'
                            FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                            WHERE  MC001=MB001
                            AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 
                            UNION ALL
                            SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                            ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,0
                            ,ISNULL((SELECT AVG((ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND ME002 LIKE '{0}%'),0) AS '成本'
                            ,'6加工' AS '分類'
                            FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                            WHERE  MC001=MB001
                            AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 
                            ) AS TEMP2
                            WHERE 1=1
                            {1}

                            ORDER BY 成品品號,分類,使用品號
                            
                        ", YM, SQUERY.ToString());

            return SB;

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrEmpty(textBox10.Text))
            {
                textBox11.Text = (Convert.ToDecimal(textBox9.Text) * Convert.ToDecimal(textBox10.Text)).ToString();
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox20.Text) && !string.IsNullOrEmpty(textBox21.Text))
            {
                textBox22.Text = (Convert.ToDecimal(textBox20.Text) * Convert.ToDecimal(textBox21.Text)).ToString();
            }
        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox7.Text = FINDMB001(textBox8.Text.ToString());
                textBox10.Text = FINDPRICES(textBox7.Text.ToString());
                textBox9.Text = "0";
            }
        }

        private void textBox19_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox18.Text = FINDMB001(textBox19.Text.ToString());
                textBox21.Text = FINDPRICES(textBox18.Text.ToString());
                textBox20.Text = "0";
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            FrmINVMBCALCOSTSUB SUB_FrmINVMBCALCOSTSUB = new FrmINVMBCALCOSTSUB();
            SUB_FrmINVMBCALCOSTSUB.ShowDialog();
            textBox8.Text = SUB_FrmINVMBCALCOSTSUB.TextBoxMsg;
            textBox7.Text = FINDMB001(textBox8.Text.ToString());
            textBox10.Text = FINDPRICES(textBox7.Text.ToString());
            textBox9.Text = "0";
        }

        private void button28_Click(object sender, EventArgs e)
        {
            FrmINVMBCALCOSTSUB SUB_FrmINVMBCALCOSTSUB = new FrmINVMBCALCOSTSUB();
            SUB_FrmINVMBCALCOSTSUB.ShowDialog();
            textBox19.Text = SUB_FrmINVMBCALCOSTSUB.TextBoxMsg;
            textBox18.Text = FINDMB001(textBox19.Text.ToString());
            textBox21.Text = FINDPRICES(textBox18.Text.ToString());
            textBox20.Text = "0";
        }

        private void textBox123_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox123.Text))
            {
                textBox125.Text = null;
            }
        }

        private void textBox125_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox125.Text))
            {
                textBox123.Text = null;
            }

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
                DELCALCOSTPRODS1RAW(textBoxID3.Text.Trim(), textBox15.Text.Trim());
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
                DELCALCOSTPRODS2MATERIL(textBoxID5.Text.Trim(), textBox27.Text.Trim());
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
        private void button20_Click(object sender, EventArgs e)
        {
            SEARCHDG9(textBox87.Text);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS7OUTPACK(textBoxID11.Text, textBox87.Text, textBox88.Text, textBox89.Text, textBox90.Text, textBox91.Text, textBox92.Text, textBox93.Text, textBox94.Text, textBox95.Text, textBox96.Text, textBox97.Text, textBox98.Text, textBox99.Text);
            SEARCHDG9(textBox87.Text);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            SEARCHDG10(textBox100.Text);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS8PRODUSEDMANU(textBoxID12.Text, textBox101.Text, textBox102.Text, textBox104.Text, textBox105.Text , textBox106.Text , textBox107.Text , textBox103.Text        );
            SEARCHDG10(textBox100.Text);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DELCALCOSTPRODS8PRODUSEDMANU(textBoxID12.Text, textBox109.Text);
            SEARCHDG10(textBox100.Text);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            UPDATECALCOSTPRODSCOSTTOTALS(textBoxID13.Text, textBox117.Text, textBox118.Text, textBox119.Text, textBox120.Text, textBox121.Text);
        }

        private void button25_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBoxID13.Text);
        }

        private void button61_Click(object sender, EventArgs e)
        {
            tabControl2.SelectedTab = tabControl2.TabPages["tabPage12"];
            SETFASTREPORT2(dateTimePicker1.Value.ToString("yyyy"), textBox123.Text, textBox124.Text, textBox125.Text);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            tabControl2.SelectedTab = tabControl2.TabPages["tabPage13"];
            SETFASTREPORT3(dateTimePicker1.Value.ToString("yyyy"), textBox123.Text , textBox125.Text); 
        }
    }


    #endregion



}
