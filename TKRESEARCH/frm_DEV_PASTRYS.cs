﻿using System;
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
    public partial class Frm_DEV_PASTRYS : Form
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
        string DATES = null;
        string DirectoryNAME = null;
        string pathFileSALESMONEYS = null;
        string NO = "";

        public Frm_DEV_PASTRYS()
        {
            InitializeComponent();

            comboBox1load();
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
                                WHERE [KIND]='TB_DEV_PASTRYS' 
                                ORDER BY [PARANAME]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void SEARCH_TB_DEV_PASTRYS(string NO, string NAMES)
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
                StringBuilder SQLquery2 = new StringBuilder();

                if (!string.IsNullOrEmpty(NO))
                {
                    SQLquery1.AppendFormat(@" AND CONVERT(NVARCHAR,[DEVCARESTEDATES],112) LIKE '%{0}%'", NO);
                }
                else
                {
                    SQLquery1.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(NAMES))
                {
                    SQLquery2.AppendFormat(@" AND [NAMES] LIKE '%{0}%'", NAMES);
                }
                else
                {
                    SQLquery2.AppendFormat(@" ");
                }

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [NO] AS '編號'
                                    ,[NAMES] AS '產品名稱'
                                    ,[SPECS] AS '規格(g)'
                                    ,CONVERT(NVARCHAR,[DEVCARESTEDATES],112) AS '開發日期'
                                    ,[BEFROECOOKEDSPCS] AS '烤前長*寬*厚(cm)'
                                    ,[AFERCOOKEDSPCS] AS '烤後長*寬*厚(cm)'
                                    ,[BEFORECOOKEDWEIGHTS] AS '烤前重量(g)'
                                    ,[AFTERCOOKEDWEIGHTS] AS '烤後重量(g)'
                                    ,[COOKEDTEMP] AS '烘焙溫度(℃)'
                                    ,[COOKEDTIMES] AS '烘焙時間(m)'
                                    ,[TOTALS] AS '總產量(片or公斤)'
                                    ,[THICKNESS] AS '延壓厚度(cm)'
                                    ,[PCTS] AS '配比(水麵:油酥)'
                                    ,[COMMETNS] AS '工作流程'
                                    ,[MB001] AS '品號'
                                    ,[ID] 
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ORDER BY [NO]
                                    ", SQLquery1.ToString(), SQLquery2.ToString());

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

        public void SEARCH_TB_DEV_PASTRYS2(string NO)
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
                StringBuilder SQLquery2 = new StringBuilder();

                if (!string.IsNullOrEmpty(NO))
                {
                    SQLquery1.AppendFormat(@" AND [NO] LIKE '%{0}%'", NO);
                }
                else
                {
                    SQLquery1.AppendFormat(@" ");
                }
              

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [NO] AS '編號'
                                    ,[NAMES] AS '產品名稱'
                                    ,[SPECS] AS '規格(g)'
                                    ,CONVERT(NVARCHAR,[DEVCARESTEDATES],112) AS '開發日期'
                                    ,[BEFROECOOKEDSPCS] AS '烤前長*寬*厚(cm)'
                                    ,[AFERCOOKEDSPCS] AS '烤後長*寬*厚(cm)'
                                    ,[BEFORECOOKEDWEIGHTS] AS '烤前重量(g)'
                                    ,[AFTERCOOKEDWEIGHTS] AS '烤後重量(g)'
                                    ,[COOKEDTEMP] AS '烘焙溫度(℃)'
                                    ,[COOKEDTIMES] AS '烘焙時間(m)'
                                    ,[TOTALS] AS '總產量(片or公斤)'
                                    ,[THICKNESS] AS '延壓厚度(cm)'
                                    ,[PCTS] AS '配比(水麵:油酥)'
                                    ,[COMMETNS] AS '工作流程'
                                    ,[ID] 
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    WHERE 1=1
                                    {0}
                                   
                                    ORDER BY [NO]
                                    ", SQLquery1.ToString());

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
            NO = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                DataGridViewRow row = dataGridView1.Rows[rowindex];

                NO = row.Cells["編號"].Value.ToString();
                SEARCH_TB_DEV_PASTRYS_DETAILS(NO);

                SETTEXT_TAB2();

                textBox2T1.Text = row.Cells["編號"].Value.ToString();
                textBox2T2.Text = row.Cells["產品名稱"].Value.ToString();
                textBox2T3.Text = row.Cells["規格(g)"].Value.ToString();
                textBox2T4.Text = row.Cells["烤前長*寬*厚(cm)"].Value.ToString();
                textBox2T5.Text = row.Cells["烤後長*寬*厚(cm)"].Value.ToString();
                textBox2T6.Text = row.Cells["烤前重量(g)"].Value.ToString();
                textBox2T7.Text = row.Cells["烤後重量(g)"].Value.ToString();
                textBox2T8.Text = row.Cells["烘焙溫度(℃)"].Value.ToString();
                textBox2T9.Text = row.Cells["烘焙時間(m)"].Value.ToString();
                textBox2T10.Text = row.Cells["總產量(片or公斤)"].Value.ToString();
                textBox2T11.Text = row.Cells["延壓厚度(cm)"].Value.ToString();
                textBox2T12.Text = row.Cells["配比(水麵:油酥)"].Value.ToString();
                textBox2T13.Text = row.Cells["工作流程"].Value.ToString();
                textBox2T14.Text = row.Cells["品號"].Value.ToString();

                //dateTimePicker2.Value= row.Cells["開發日期"].Value.ToString();
                DateTime dateTime;
                if (DateTime.TryParseExact(row.Cells["開發日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
                {
                    //Console.WriteLine(dateTime.ToString()); // 输出转换后的日期时间
                    dateTimePicker2.Value = dateTime;
                }

               
            }


        }

        public void SEARCH_TB_DEV_PASTRYS_DETAILS(string NO)
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

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                  SELECT
                                    [NO] AS '編號'
                                    ,[KINDS] AS '品項'
                                    ,[SEQ] AS '投料順序'
                                    ,[CODE] AS '代號'
                                    ,[SUPPLIERS] AS '供應商'
                                    ,[NAMES] AS '原料品項'
                                    ,[PCTS] AS '各自百分比(%)'
                                    ,[WEIGHTS] AS '各自重量(g)'
                                    ,[TPCTS] AS '加總後百分比(%)'
                                    ,[TWEIGHTS] AS '加總後重量(g)'
                                    ,[MB001] AS '品號'
                                    , [ID]
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    WHERE [NO]='{0}'
                                    ORDER BY [KINDS],[CODE]
                                  
                                    ", NO);

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

        public void SEARCH_TB_DEV_PASTRYS_DETAILS2(string NO)
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

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                  SELECT
                                    [NO] AS '編號'
                                    ,[KINDS] AS '品項'
                                    ,[SEQ] AS '投料順序'
                                    ,[CODE] AS '代號'
                                    ,[SUPPLIERS] AS '供應商'
                                    ,[NAMES] AS '原料品項'
                                    ,[PCTS] AS '各自百分比(%)'
                                    ,[WEIGHTS] AS '各自重量(g)'
                                    ,[TPCTS] AS '加總後百分比(%)'
                                    ,[TWEIGHTS] AS '加總後重量(g)'
                                    ,[MB001] AS '品號'
                                    , [ID]
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    WHERE [NO]='{0}'
                                    ORDER BY [KINDS],[CODE]
                                  
                                    ", NO);

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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                DataGridViewRow row = dataGridView3.Rows[rowindex];

                NO = row.Cells["編號"].Value.ToString();
                string ID = row.Cells["ID"].Value.ToString();
               
                SETTEXT_TAB2B();
                  
                textBox2T30.Text = row.Cells["編號"].Value.ToString();
                
                textBox2T32.Text = row.Cells["投料順序"].Value.ToString();
                textBox2T33.Text = row.Cells["代號"].Value.ToString();
                textBox2T34.Text = row.Cells["供應商"].Value.ToString();
                textBox2T35.Text = row.Cells["原料品項"].Value.ToString();
                textBox2T36.Text = row.Cells["各自百分比(%)"].Value.ToString();
                textBox2T37.Text = row.Cells["各自重量(g)"].Value.ToString();
                textBox2T38.Text = row.Cells["加總後百分比(%)"].Value.ToString();
                textBox2T39.Text = row.Cells["加總後重量(g)"].Value.ToString();
                textBox2T40.Text = ID;
                textBox2T41.Text = row.Cells["品號"].Value.ToString();

                comboBox1.Text= row.Cells["品項"].Value.ToString();



            }
        }

        public void UPDATE_TB_DEV_PASTRYS_DETAILS(
            string ID
            , string NO
            , string KINDS
            , string SEQ
            , string CODE
            , string SUPPLIERS
            , string NAMES
            , string PCTS
            , string WEIGHTS
            , string TPCTS
            , string TWEIGHTS
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
                                    UPDATE  [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    SET
                                    NO='{1}'
                                    ,KINDS='{2}'
                                    ,SEQ='{3}'
                                    ,CODE='{4}'
                                    ,SUPPLIERS='{5}'
                                    ,NAMES='{6}'
                                    ,PCTS='{7}'
                                    ,WEIGHTS='{8}'
                                    ,TPCTS='{9}'
                                    ,TWEIGHTS='{10}'
                                    WHERE ID='{0}'
                                    ", ID
                                    , NO
                                    , KINDS
                                    , SEQ
                                    , CODE
                                    , SUPPLIERS
                                    , NAMES
                                    , PCTS
                                    , WEIGHTS
                                    , TPCTS
                                    , TWEIGHTS
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

        public void INSERT_TB_DEV_PASTRYS_DETAILS(
            string ID
            , string NO
            , string KINDS
            , string SEQ
            , string CODE
            , string SUPPLIERS
            , string NAMES
            , string PCTS
            , string WEIGHTS
            , string TPCTS
            , string TWEIGHTS
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
                                   INSERT INTO [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    (
                                    NO
                                    ,KINDS
                                    ,SEQ
                                    ,CODE
                                    ,SUPPLIERS
                                    ,NAMES
                                    ,PCTS
                                    ,WEIGHTS
                                    ,TPCTS
                                    ,TWEIGHTS
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
                                    )
                                    "
                                    , NO
                                    , KINDS
                                    , SEQ
                                    , CODE
                                    , SUPPLIERS
                                    , NAMES
                                    , PCTS
                                    , WEIGHTS
                                    , TPCTS
                                    , TWEIGHTS
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
        public void DELETE_TB_DEV_PASTRYS_DETAILS(string ID)
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
                                    DELETE  [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]                                   
                                    WHERE ID='{0}'
                                    ", ID
                                 
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

        public void INSERT_TB_DEV_PASTRYS(
            string NO
            , string NAMES
            , string SPECS
            , string DEVCARESTEDATES
            , string BEFROECOOKEDSPCS
            , string AFERCOOKEDSPCS
            , string BEFORECOOKEDWEIGHTS
            , string AFTERCOOKEDWEIGHTS
            , string COOKEDTEMP
            , string COOKEDTIMES
            , string TOTALS
            , string THICKNESS
            , string PCTS
            , string COMMETNS

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
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    (
                                    NO
                                    ,NAMES
                                    ,SPECS
                                    ,DEVCARESTEDATES
                                    ,BEFROECOOKEDSPCS
                                    ,AFERCOOKEDSPCS
                                    ,BEFORECOOKEDWEIGHTS
                                    ,AFTERCOOKEDWEIGHTS
                                    ,COOKEDTEMP
                                    ,COOKEDTIMES
                                    ,TOTALS
                                    ,THICKNESS
                                    ,PCTS
                                    ,COMMETNS
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
                                     , NO
                                    , NAMES
                                    , SPECS
                                    , DEVCARESTEDATES
                                    , BEFROECOOKEDSPCS
                                    , AFERCOOKEDSPCS
                                    , BEFORECOOKEDWEIGHTS
                                    , AFTERCOOKEDWEIGHTS
                                    , COOKEDTEMP
                                    , COOKEDTIMES
                                    , TOTALS
                                    , THICKNESS
                                    , PCTS
                                    , COMMETNS
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

        public void UPDATE_TB_DEV_PASTRYS(
             string NO
            , string NAMES
            , string SPECS
            , string DEVCARESTEDATES
            , string BEFROECOOKEDSPCS
            , string AFERCOOKEDSPCS
            , string BEFORECOOKEDWEIGHTS
            , string AFTERCOOKEDWEIGHTS
            , string COOKEDTEMP
            , string COOKEDTIMES
            , string TOTALS
            , string THICKNESS
            , string PCTS
            , string COMMETNS

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

                                    UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    SET
                                    [NAMES]='{1}'
                                    ,[SPECS]='{2}'
                                    ,[DEVCARESTEDATES]='{3}'
                                    ,[BEFROECOOKEDSPCS]='{4}'
                                    ,[AFERCOOKEDSPCS]='{5}'
                                    ,[BEFORECOOKEDWEIGHTS]='{6}'
                                    ,[AFTERCOOKEDWEIGHTS]='{7}'
                                    ,[COOKEDTEMP]='{8}'
                                    ,[COOKEDTIMES]='{9}'
                                    ,[TOTALS]='{10}'
                                    ,[THICKNESS]='{11}'
                                    ,[PCTS]='{12}'
                                    ,[COMMETNS]='{13}'
                                    WHERE  [NO]='{0}'                                    
                                    "
                                     , NO
                                    , NAMES
                                    , SPECS
                                    , DEVCARESTEDATES
                                    , BEFROECOOKEDSPCS
                                    , AFERCOOKEDSPCS
                                    , BEFORECOOKEDWEIGHTS
                                    , AFTERCOOKEDWEIGHTS
                                    , COOKEDTEMP
                                    , COOKEDTIMES
                                    , TOTALS
                                    , THICKNESS
                                    , PCTS
                                    , COMMETNS
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

        public void DELETE_TB_DEV_PASTRYS(string NO)
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
                                    DELETE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]                                  
                                    WHERE  [NO]='{0}'                                    
                                    "
                                     , NO
                                
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

        public string GETMAXNO(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                        SELECT
                                        ISNULL(MAX(NO),'0')  AS 'NO'
                                        FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                        WHERE [NO] LIKE '{0}%'
                                        ORDER BY [NO] DESC
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        string NEWNO = ds1.Tables["ds1"].Rows[0]["NO"].ToString();

                        if (NEWNO.Equals("0"))
                        {
                            return NO + "-"+"001";
                        }

                        else
                        {
                            int serno = Convert.ToInt16(NEWNO.Substring(6, 3));
                            serno = serno + 1;
                            string temp = serno.ToString();
                            temp = temp.PadLeft(3, '0');
                            return NO + "-" + temp.ToString();
                        }
                       

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void SETTEXT_TAB2()
        {
            textBox2T1.Text = null;
            textBox2T2.Text = null;
            textBox2T3.Text = null;
            textBox2T4.Text = null;
            textBox2T5.Text = null;
            textBox2T6.Text = null;
            textBox2T7.Text = null;
            textBox2T8.Text = null;
            textBox2T9.Text = null;
            textBox2T10.Text = null;
            textBox2T11.Text = null;
            textBox2T12.Text = null;
            textBox2T13.Text = null;

        }

        public void SETTEXT_TAB2B()
        {
            textBox2T30.Text = null;
            //textBox2T31.Text = null;
            textBox2T32.Text = null;
            textBox2T33.Text = null;
            textBox2T34.Text = null;
            textBox2T35.Text = null;
            textBox2T36.Text = null;
            textBox2T37.Text = null;
            textBox2T38.Text = null;
            textBox2T39.Text = null;
            textBox2T40.Text = null;
        

        }
        public DataTable FIND_NO_KINDS(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                      SELECT [NO],[KINDS]
                                        FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                        WHERE [NO] = '{0}' 
                                        GROUP BY [NO],[KINDS]
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        /// <summary>
        ///  //計算百分比
        /// </summary>
        /// <param name="NO"></param>
        /// <param name="KINDS"></param>
        public void CAL_TB_DEV_PASTRYS_DETAILS_PCTS(string NO)
        {
            DataTable DT=FIND_NO_KINDS(NO);

            foreach (DataRow DR in DT.Rows)
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
                                   WITH CTE AS (
                                        SELECT 
                                            [ID],
                                            [NO],
                                            [KINDS],
                                            [SEQ],
                                            [CODE],
                                            [SUPPLIERS],
                                            [NAMES],
                                            [PCTS],
                                            [WEIGHTS],
                                            [TPCTS],
                                            [TWEIGHTS],
                                            SUM([WEIGHTS]) OVER (PARTITION BY [NO], [KINDS]) AS SUMWEIGHTS,
                                            CONVERT(decimal(16,4),([WEIGHTS]/SUM([WEIGHTS]) OVER (PARTITION BY [NO], [KINDS]))) AS CALPCTS,
                                            MAX([SEQ]) OVER (PARTITION BY [NO], [KINDS]) AS MAXSEQ
                                        FROM 
                                            [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                        WHERE 
                                            [NO] = '{0}' AND [KINDS] = '{1}'
                                    ) 

                                    SELECT 
                                        *,
                                        (SELECT SUM(CALPCTS) FROM CTE CTE2 WHERE CTE2.SEQ<CTE.SEQ) AS SUMPERPCT,
                                        CONVERT(decimal(16,4),(CASE WHEN CTE.SEQ<>MAXSEQ THEN CALPCTS ELSE 1-(SELECT SUM(CALPCTS) FROM CTE CTE2 WHERE CTE2.SEQ<CTE.SEQ) END)) AS FINALPCTS
                                    INTO 
                                        #TEMP_UPDATED_PASTRYS_DETAILS  -- 将结果插入临时表中
                                    FROM  
                                        CTE;

                                    -- 使用临时表来更新原始表
                                    UPDATE 
                                        TKRESEARCH.dbo.TB_DEV_PASTRYS_DETAILS
                                    SET 
                                        TB_DEV_PASTRYS_DETAILS.PCTS = CTE.FINALPCTS
                                    FROM 
                                        #TEMP_UPDATED_PASTRYS_DETAILS CTE
                                    WHERE 
                                        CTE.ID = TB_DEV_PASTRYS_DETAILS.ID;

                                    -- 删除临时表
                                    DROP TABLE #TEMP_UPDATED_PASTRYS_DETAILS;
                              
                                    "
                                         , DR["NO"].ToString(), DR["KINDS"].ToString()

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
            
        }

        public void CAL_MOC_USE(string NO)
        {
            string NAMES = "";
            string BASESNUM = "0";
            string MULTIS = "0";
            string PASTRYSPCTS = "0";
            string WATERSWEIGHTS = "0";

            //找出要調的 油酥，基礎 麵粉
            DataTable DT1 = FIND_TB_DEV_PASTRYS_BASES(NO);
            if(DT1!= null && DT1.Rows.Count>=1)
            {
                foreach(DataRow DR in DT1.Rows)
                {
                    NAMES = DR["NAMES"].ToString();
                    BASESNUM = DR["BASESNUM"].ToString();

                    //在 油酥，找出 麵粉 的用量，用66計算出的倍數
                    DataTable DT2 = FIND_TB_DEV_PASTRYS_DETAILS_BASES(NO, NAMES, BASESNUM);
                    if(DT2!=null && DT2.Rows.Count>=1)
                    {
                        MULTIS = DT2.Rows[0]["MULTIS"].ToString();
                        //在 油酥，用倍數計算除 麵粉 以外的用量，並更新總百分比
                        if (Convert.ToDecimal(MULTIS) > 0)
                        {
                            //用 油酥 的總重，用水麵、油酥的佔比，再算出桶的總水麵用量是多少
                            UPDATE_TB_DEV_PASTRYS_DETAILS_TWEIGHTS(NO, MULTIS, NAMES, BASESNUM);
                            //用總水麵用量是多少乘各水麵的佔比，算出桶的用量，並更新總百分比
                            DataTable DT3 = FIND_TB_DEV_PASTRYS_PCTS(NO);
                            if(DT3!=null && DT3.Rows.Count>=1)
                            {
                                PASTRYSPCTS= DT3.Rows[0]["PCTS"].ToString();
                                WATERSWEIGHTS = DT3.Rows[0]["WATERSWEIGHTS"].ToString();

                                if (Convert.ToDecimal(WATERSWEIGHTS) >0)
                                {
                                    UPDATE_TB_DEV_PASTRYS_DETAILS_TWEIGHTS_WATERS(NO,NAMES, WATERSWEIGHTS);
                                }
                            }
                        }
                    }
                }
            }            
        }

        public DataTable FIND_TB_DEV_PASTRYS_BASES(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                      SELECT 
                                    [TB_DEV_PASTRYS_BASES].[NAMES],[BASESNUM]
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_BASES],[TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    WHERE [TB_DEV_PASTRYS_BASES].NAMES=[TB_DEV_PASTRYS_DETAILS].NAMES
                                    AND [TB_DEV_PASTRYS_DETAILS].NO='{0}'
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataTable FIND_TB_DEV_PASTRYS_DETAILS_BASES(string NO,string NAMES,string BASESNUM)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                     SELECT 
                                    [ID]
                                    ,[NO]
                                    ,[KINDS]
                                    ,[SEQ]
                                    ,[CODE]
                                    ,[SUPPLIERS]
                                    ,[NAMES]
                                    ,[PCTS]
                                    ,[WEIGHTS]
                                    ,[TPCTS]
                                    ,[TWEIGHTS]
                                    ,CONVERT(decimal(16,4),{2}/[WEIGHTS]) AS 'MULTIS'
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    WHERE [NO]='{0}' AND [KINDS]='油酥' AND [NAMES]='{1}'
                                    AND EXISTS (SELECT * FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS] WHERE [NAMES]='{1}')
                                        ", NO, NAMES, BASESNUM);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATE_TB_DEV_PASTRYS_DETAILS_TWEIGHTS(string NO,string MULTIS ,string NAMES,string BASBASESNUM)
        {
            DataTable DT = FIND_NO_KINDS(NO);

            foreach (DataRow DR in DT.Rows)
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
                                        UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                        SET [TWEIGHTS]=[WEIGHTS]*{2},[TPCTS]=[PCTS]
                                        WHERE [NO]='{0}' AND [KINDS]='油酥' AND  [NAMES]<>'{1}'
                                        AND EXISTS (SELECT * FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS] WHERE [NAMES]='{1}')

                                        UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                        SET [TWEIGHTS]={3},[TPCTS]=[PCTS]
                                        WHERE [NO]='{0}' AND [KINDS]='油酥' AND  [NAMES]='{1}'
                                        ", NO, NAMES, MULTIS, BASBASESNUM


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

        }

        public DataTable FIND_TB_DEV_PASTRYS_PCTS(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT *
                                    ,(CASE WHEN [PCTS]>0 AND TWEIGHTS>0 THEN CONVERT(DECIMAL(16,4),TWEIGHTS/[PCTS]) ELSE 0 END ) AS 'WATERSWEIGHTS'
                                    FROM 
                                    (
                                    SELECT 
                                    [ID]
                                    ,[NO]
                                    ,[PCTS]
                                    ,(SELECT SUM([TWEIGHTS]) FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS] WHERE [TB_DEV_PASTRYS_DETAILS].NO=[TB_DEV_PASTRYS].NO AND [TB_DEV_PASTRYS_DETAILS].[KINDS]='油酥') AS 'TWEIGHTS'
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    WHERE [NO]='{0}'
                                    ) AS TEMP
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void UPDATE_TB_DEV_PASTRYS_DETAILS_TWEIGHTS_WATERS(string NO,string NAMES, string WATERSWEIGHTS)
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
                                        UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                        SET [TWEIGHTS]=[PCTS]*{2},[TPCTS]=[PCTS]
                                        WHERE [NO]='{0}' AND [KINDS]='水麵' 
                                        AND EXISTS (SELECT * FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS] WHERE [NAMES]='{1}')
                                        ", NO, NAMES, WATERSWEIGHTS


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

        public void UPDATE_TB_DEV_PASTRYS_BASES_MB001(string NO,string MB001)
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    SET [MB001]='{1}'
                                    WHERE [NO]='{0}'
                                        ", NO, MB001

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
        public void UPDATE_TB_DEV_PASTRYS_DETAILS_MB001(string ID, string MB001)
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    SET [MB001]='{1}'
                                    WHERE [ID]='{0}'
                                        ", ID, MB001

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


        public DataTable CHECK_MB001(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT *
                                    FROM 
                                    (
                                    SELECT 
                                    [ID]
                                    ,[NO]
                                    ,[MB001]
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
                                    WHERE NO='{0}' AND ISNULL([MB001],'')=''
                                    UNION ALL
                                    SELECT 
                                    [ID]
                                    ,[NO]
                                    ,[MB001]
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS]
                                    WHERE NO='{0}' AND ISNULL([MB001],'')=''
                                    ) AS TEMP
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_BOMMJ_BOMMK(string NO)
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

INSERT INTO [TK].[dbo].[BOMMJ]
(
[COMPANY]
,[CREATOR]
,[USR_GROUP]
,[CREATE_DATE]
,[MODIFIER]
,[MODI_DATE]
,[FLAG]
,[CREATE_TIME]
,[MODI_TIME]
,[TRANS_TYPE]
,[TRANS_NAME]
,[sync_date]
,[sync_time]
,[sync_mark]
,[sync_count]
,[DataUser]
,[DataGroup]
,[MJ001]
,[MJ004]
,[MJ005]
,[MJ006]
,[MJ007]
,[MJ008]
,[MJ009]
,[MJ010]
,[MJ011]
,[MJ012]
,[MJ013]
,[MJ014]
,[MJ015]
,[MJ016]
,[MJ017]
,[MJ018]
,[MJ019]
,[MJ020]
,[MJ021]
,[MJ022]
)
SELECT 
'TK' [COMPANY]
,'160033' [CREATOR]
,'104000' [USR_GROUP]
,CONVERT(NVARCHAR,GETDATE(),112) [CREATE_DATE]
,'160033' [MODIFIER]
,CONVERT(NVARCHAR,GETDATE(),112) [MODI_DATE]
,'0' [FLAG]
,CONVERT(VARCHAR(8), GETDATE(), 108) [CREATE_TIME]
,CONVERT(VARCHAR(8), GETDATE(), 108) [MODI_TIME]
,'P001' [TRANS_TYPE]
,'BOMI11' [TRANS_NAME]
,'' [sync_date]
,'' [sync_time]
,'' [sync_mark]
,'0' [sync_count]
,'' [DataUser]
,'104000' [DataGroup]
,[MB001] [MJ001]
,(SELECT SUM([TWEIGHTS]) FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS] WHERE [TB_DEV_PASTRYS_DETAILS].NO=[TB_DEV_PASTRYS].NO)[MJ004]
,'A510' [MJ005]
,'' [MJ006]
,'' [MJ007]
,'' [MJ008]
,'0' [MJ009]
,'' [MJ010]
,'N' [MJ011]
,'0' [MJ012]
,'0' [MJ013]
,'' [MJ014]
,'' [MJ015]
,'' [MJ016]
,'' [MJ017]
,'0' [MJ018]
,'' [MJ019]
,'' [MJ020]
,'' [MJ021]
,'0' [MJ022]
FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
WHERE NO='{0}'
AND MB001 NOT IN (SELECT [MJ001] FROM [TK].[dbo].[BOMMJ])

SELECT 
'TK' [COMPANY]
,'160033' [CREATOR]
,'104000' [USR_GROUP]
,CONVERT(NVARCHAR,GETDATE(),112) [CREATE_DATE]
,'160033' [MODIFIER]
,CONVERT(NVARCHAR,GETDATE(),112) [MODI_DATE]
,'0' [FLAG]
,CONVERT(VARCHAR(8), GETDATE(), 108) [CREATE_TIME]
,CONVERT(VARCHAR(8), GETDATE(), 108) [MODI_TIME]
,'P001' [TRANS_TYPE]
,'BOMI11' [TRANS_NAME]
,'' [sync_date]
,'' [sync_time]
,'' [sync_mark]
,'0' [sync_count]
,'' [DataUser]
,'104000' [DataGroup]
,[TB_DEV_PASTRYS].MB001 [MK001]
,RIGHT(REPLICATE('0', 4) + CAST([SEQ] * 10 AS VARCHAR(4)), 4) [MK002]
,[TB_DEV_PASTRYS_DETAILS].[MB001] [MK003]
,[TWEIGHTS] [MK006]
,'1' [MK007]
,'0' [MK008]
,'****' [MK009]
,'' [MK010]
,'' [MK011]
,'' [MK012]
,'N' [MK013]
,'N' [MK014]
,'' [MK015]
,'' [MK016]
,'1' [MK017]
,'0' [MK018]
,'1' [MK019]
,'' [MK020]
,'' [MK021]
,'' [MK022]
,'' [MK023]
,'' [MK024]
,'' [MK025]
,'' [MK026]
,'' [MK027]
,'0' [MK028]
,'0' [MK029]
,'' [MK030]
,'' [MK031]
,'' [MK032]
,'0' [MK033]
,'' [MK034]
,'' [MK035]
,'' [MK036]
,'' [MK037]
,'0' [MK038]
FROM [TKRESEARCH].[dbo].[TB_DEV_PASTRYS_DETAILS],[TKRESEARCH].[dbo].[TB_DEV_PASTRYS]
WHERE [TB_DEV_PASTRYS].NO=[TB_DEV_PASTRYS_DETAILS].NO
AND [TB_DEV_PASTRYS_DETAILS].NO='{0}'
AND ISNULL([TB_DEV_PASTRYS_DETAILS].MB001,'')<>''
                                  
                                        ", NO

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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_PASTRYS(dateTimePicker1.Value.ToString("yyyy"),textBox1.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // 或者使用 SelectedTab 屬性直接指定 Tab 頁面物件
            tabControl1.SelectedTab = tabPage2;
            // 在某個地方調用 PerformClick() 方法來觸發按鈕的點擊事件
            button3.PerformClick();
            //MessageBox.Show(NO);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_PASTRYS2(textBox2T1.Text);
            SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);

        }

        private void button7_Click(object sender, EventArgs e)
        {
            UPDATE_TB_DEV_PASTRYS(
                 textBox2T1.Text
                , textBox2T2.Text
                , textBox2T3.Text
                , dateTimePicker2.Value.ToString("yyyy/MM/dd")
                , textBox2T4.Text
                , textBox2T5.Text
                , textBox2T6.Text
                , textBox2T7.Text
                , textBox2T8.Text
                , textBox2T9.Text
                , textBox2T10.Text
                , textBox2T11.Text
                , textBox2T12.Text
                , textBox2T13.Text
                );

            SEARCH_TB_DEV_PASTRYS2(textBox2T1.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TB_DEV_PASTRYS(textBox2T1.Text);
                SEARCH_TB_DEV_PASTRYS2(textBox2T1.Text);

                SETTEXT_TAB2();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            INSERT_TB_DEV_PASTRYS(
            textBox2T1.Text
            , textBox2T2.Text
            , textBox2T3.Text
            , dateTimePicker2.Value.ToString("yyyy/MM/dd")
            , textBox2T4.Text
            , textBox2T5.Text
            , textBox2T6.Text
            , textBox2T7.Text
            , textBox2T8.Text
            , textBox2T9.Text
            , textBox2T10.Text
            , textBox2T11.Text
            , textBox2T12.Text
            , textBox2T13.Text

            );

            SEARCH_TB_DEV_PASTRYS2(textBox2T1.Text);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            INSERT_TB_DEV_PASTRYS_DETAILS(
                ""
            , textBox2T30.Text
            , comboBox1.Text
            , textBox2T32.Text
            , textBox2T33.Text
            , textBox2T34.Text
            , textBox2T35.Text
            , textBox2T36.Text
            , textBox2T37.Text
            , textBox2T38.Text
            , textBox2T39.Text
            );
           
            CAL_TB_DEV_PASTRYS_DETAILS_PCTS(textBox2T30.Text);
            SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            UPDATE_TB_DEV_PASTRYS_DETAILS(
            textBox2T40.Text
            , textBox2T30.Text
            , comboBox1.Text
            , textBox2T32.Text
            , textBox2T33.Text
            , textBox2T34.Text
            , textBox2T35.Text
            , textBox2T36.Text
            , textBox2T37.Text
            , textBox2T38.Text
            , textBox2T39.Text
            );

            CAL_TB_DEV_PASTRYS_DETAILS_PCTS(textBox2T30.Text);
            SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);
        }

        private void button10_Click(object sender, EventArgs e)
        {           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TB_DEV_PASTRYS_DETAILS(textBox2T40.Text);
                CAL_TB_DEV_PASTRYS_DETAILS_PCTS(textBox2T30.Text);
                SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string DATES = dateTimePicker2.Value.ToString("yyyy-MM");
            DATES = DATES.Substring(2,5);
            string NO = GETMAXNO(DATES);
            textBox2T1.Text = NO;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            CAL_MOC_USE(textBox2T1.Text);
            SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);
        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2T14.Text.Trim()))
            {
                UPDATE_TB_DEV_PASTRYS_BASES_MB001(textBox2T1.Text.Trim(), textBox2T14.Text.Trim());
                SEARCH_TB_DEV_PASTRYS2(textBox2T1.Text);
            }  
            else
            {
                MessageBox.Show("未填寫BOM品號");
            }            
        
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox2T41.Text.Trim()))
            {
                UPDATE_TB_DEV_PASTRYS_DETAILS_MB001(textBox2T40.Text.Trim(), textBox2T41.Text.Trim());

                SEARCH_TB_DEV_PASTRYS_DETAILS2(textBox2T1.Text);
            }
            else
            {
                MessageBox.Show("未填寫BOM品號");
            }

        }
        private void button15_Click(object sender, EventArgs e)
        {
            //CHECK_MB001
            string NO = textBox2T1.Text.Trim();
            DataTable DT =CHECK_MB001(NO);

            if(DT!=null && DT.Rows.Count>=1)
            {
                MessageBox.Show(NO+Environment.NewLine+"有品號未填寫");
            }
            else
            {
                ADD_BOMMJ_BOMMK(NO);
            }
        }
        #endregion


    }
}
