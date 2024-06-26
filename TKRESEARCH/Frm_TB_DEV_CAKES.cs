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
    public partial class Frm_TB_DEV_CAKES : Form
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


        public Frm_TB_DEV_CAKES()
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
                                WHERE [KIND]='TB_DEV_CAKES' 
                                ORDER BY [PARAID]
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
        public void SEARCH_TB_DEV_CAKES(string NO, string NAMES)
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
                                    ,[TOTALSWEIGHTS] AS '總產量(片or公斤)'
                                    ,[MODELS] AS '模具'             
                                    ,[MOQS] AS '最低生產批量'
                                    ,[COMMETNS] AS '工作流程'
                                    ,[EGGSLEVELS] AS '蛋白打發程度'
                                    ,[EGGSTIMES] AS '蛋白時間'
                                    ,[EGGSSPEEDS] AS '蛋白轉速'
                                    ,[EGGSWEIGHTS] AS '蛋白重量'
                                    ,[FILLINGWEIGHTS] AS '餡重量'
                                    ,[DECWEIGHTS] AS '裝飾重量'
                                    ,[MB001] AS '品號'
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES]
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            NO = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                DataGridViewRow row = dataGridView1.Rows[rowindex];

                NO = row.Cells["編號"].Value.ToString();
                SEARCH_TB_DEV_CAKES_DETAILS(NO);

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
                textBox2T11.Text = row.Cells["模具"].Value.ToString();
                textBox2T12.Text = row.Cells["工作流程"].Value.ToString();
                textBox2T13.Text = row.Cells["品號"].Value.ToString();
                textBox2T14.Text = row.Cells["蛋白打發程度"].Value.ToString();
                textBox2T15.Text = row.Cells["蛋白時間"].Value.ToString();
                textBox2T16.Text = row.Cells["蛋白轉速"].Value.ToString();
                textBox2T17.Text = row.Cells["蛋白重量"].Value.ToString();
                textBox2T18.Text = row.Cells["餡重量"].Value.ToString();
                textBox2T19.Text = row.Cells["裝飾重量"].Value.ToString();

                //dateTimePicker2.Value= row.Cells["開發日期"].Value.ToString();
                DateTime dateTime;
                if (DateTime.TryParseExact(row.Cells["開發日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
                {
                    //Console.WriteLine(dateTime.ToString()); // 输出转换后的日期时间
                    dateTimePicker2.Value = dateTime;
                }


            }
        }

        public void SEARCH_TB_DEV_CAKES_DETAILS(string NO)
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
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
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

        public void SEARCH_TB_DEV_CAKES2(string NO)
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
                                    ,[TOTALSWEIGHTS] AS '總產量(片or公斤)'
                                    ,[MODELS] AS '模具'             
                                    ,[MOQS] AS '最低生產批量'
                                    ,[COMMETNS] AS '工作流程'
                                    ,[EGGSLEVELS] AS '蛋白打發程度'
                                    ,[EGGSTIMES] AS '蛋白時間'
                                    ,[EGGSSPEEDS] AS '蛋白轉速'
                                    ,[EGGSWEIGHTS] AS '蛋白重量'
                                    ,[FILLINGWEIGHTS] AS '餡重量'
                                    ,[DECWEIGHTS] AS '裝飾重量'
                                    ,[MB001] AS '品號'
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES]       
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

        public void SEARCH_TB_DEV_CAKES_DETAILS2(string NO)
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
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
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
                                        FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES]
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
                            return NO + "-" + "001";
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

        public void INSERT_TB_DEV_CAKES(
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
                , string TOTALSWEIGHTS
                , string MODELS
                , string MOQS
                , string COMMETNS
                , string MB001
                , string EGGSLEVELS
                , string EGGSTIMES
                , string EGGSSPEEDS
                , string EGGSWEIGHTS
                , string FILLINGWEIGHTS
                , string DECWEIGHTS
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
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_DEV_CAKES]
                                    (
                                    [NO]
                                    ,[NAMES]
                                    ,[SPECS]
                                    ,[DEVCARESTEDATES]
                                    ,[BEFROECOOKEDSPCS]
                                    ,[AFERCOOKEDSPCS]
                                    ,[BEFORECOOKEDWEIGHTS]
                                    ,[AFTERCOOKEDWEIGHTS]
                                    ,[COOKEDTEMP]
                                    ,[COOKEDTIMES]
                                    ,[TOTALSWEIGHTS]
                                    ,[MODELS]
                                    ,[MOQS]
                                    ,[COMMETNS]
                                    ,[MB001]
                                    ,[EGGSLEVELS]
                                    ,[EGGSTIMES]
                                    ,[EGGSSPEEDS]
                                    ,[EGGSWEIGHTS]
                                    ,[FILLINGWEIGHTS]
                                    ,[DECWEIGHTS]
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
                                    ,'{14}'
                                    ,'{15}'
                                    ,'{16}'
                                    ,'{17}'
                                    ,'{18}'
                                    ,'{19}'
                                    ,'{20}'
                                  
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
                                    , TOTALSWEIGHTS
                                    , MODELS
                                    , MOQS
                                    , COMMETNS
                                    , MB001
                                    , EGGSLEVELS
                                    , EGGSTIMES
                                    , EGGSSPEEDS
                                    , EGGSWEIGHTS
                                    , FILLINGWEIGHTS
                                    , DECWEIGHTS
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

        public void UPDATE_TB_DEV_CAKES(
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
                , string TOTALSWEIGHTS
                , string MODELS
                , string MOQS
                , string COMMETNS
                , string MB001
                , string EGGSLEVELS
                , string EGGSTIMES
                , string EGGSSPEEDS
                , string EGGSWEIGHTS
                , string FILLINGWEIGHTS
                , string DECWEIGHTS

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
                                    UPDATE  [TKRESEARCH].[dbo].[TB_DEV_CAKES]
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
                                    ,[TOTALSWEIGHTS]='{10}'
                                    ,[MODELS]='{11}'
                                    ,[MOQS]='{12}'
                                    ,[COMMETNS]='{13}'
                                    ,[MB001]='{14}'
                                    ,[EGGSLEVELS]='{15}'
                                    ,[EGGSTIMES]='{16}'
                                    ,[EGGSSPEEDS]='{17}'
                                    ,[EGGSWEIGHTS]='{18}'
                                    ,[FILLINGWEIGHTS]='{19}'
                                    ,[DECWEIGHTS]='{20}'
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
                                    , TOTALSWEIGHTS
                                    , MODELS
                                    , MOQS
                                    , COMMETNS
                                    , MB001
                                    , EGGSLEVELS
                                    , EGGSTIMES
                                    , EGGSSPEEDS
                                    , EGGSWEIGHTS
                                    , FILLINGWEIGHTS
                                    , DECWEIGHTS
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
        public void DELETE_TB_DEV_CAKES(string NO)
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
                                    DELETE [TKRESEARCH].[dbo].[TB_DEV_CAKES]                                  
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

        public void UPDATE_TB_DEV_CAKES_MB001(string NO, string MB001)
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_DEV_CAKES]
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

        public void INSERT_TB_DEV_CAKES_DETAILS(
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

                if (string.IsNullOrEmpty(PCTS))
                {
                    PCTS = "1";
                }
                if (string.IsNullOrEmpty(WEIGHTS))
                {
                    WEIGHTS = "0";
                }
                if (string.IsNullOrEmpty(TPCTS))
                {
                    TPCTS = "1";
                }
                if (string.IsNullOrEmpty(TWEIGHTS))
                {
                    TWEIGHTS = "0";
                }

                sbSql.AppendFormat(@" 
                                   INSERT INTO [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
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

                comboBox1.Text = row.Cells["品項"].Value.ToString();



            }
        }

        public void UPDATE_TB_DEV_CAKES_DETAILS(
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
                                    UPDATE  [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
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

        public void DELETE_TB_DEV_CAKES_DETAILS(string ID)
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
                                    DELETE  [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]                                   
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

        public void UPDATE_TB_DEV_CAKES_DETAILS_MB001(string ID, string MB001)
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
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
            textBox2T14.Text = null;
            textBox2T15.Text = null;
            textBox2T16.Text = null;
            textBox2T17.Text = null;
            textBox2T18.Text = null;
            textBox2T19.Text = null;

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

        private void textBox2T1_TextChanged(object sender, EventArgs e)
        {
            textBox2T30.Text = textBox2T1.Text;
        }

        public void CAL_TB_DEV_CAKES_DETAILS(string NO)
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
                                    SELECT  
                                    [NO]
                                    ,[KINDS]
                                    ,SUM([WEIGHTS]) AS 'KINDSWEIGHTS'
                                    INTO #KindsTotalWeight
                                    FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
                                    WHERE [NO]='{0}'
                                    GROUP BY [NO],[KINDS]

                                    UPDATE  [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS]
                                    SET [TWEIGHTS]=[WEIGHTS]
                                    ,[PCTS]=CONVERT(decimal(16,4),[WEIGHTS]/#KindsTotalWeight.KINDSWEIGHTS)
                                    ,[TPCTS]=CONVERT(decimal(16,4),[WEIGHTS]/(SELECT SUM([WEIGHTS]) FROM [TKRESEARCH].[dbo].[TB_DEV_CAKES_DETAILS] AS DE2 WHERE DE2.NO=[TB_DEV_CAKES_DETAILS].NO))
                                    FROM #KindsTotalWeight 
                                    WHERE [TB_DEV_CAKES_DETAILS].NO=#KindsTotalWeight.NO AND  [TB_DEV_CAKES_DETAILS].KINDS=#KindsTotalWeight.KINDS 
                                    AND [TB_DEV_CAKES_DETAILS].[NO]='{0}'

                              
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_CAKES(dateTimePicker1.Value.ToString("yyyy"), textBox1.Text);
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
            SEARCH_TB_DEV_CAKES2(textBox2T1.Text);
            SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);
        }
        private void button8_Click(object sender, EventArgs e)
        {
            INSERT_TB_DEV_CAKES_DETAILS(
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

            CAL_TB_DEV_CAKES_DETAILS(textBox2T1.Text);
            SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string DATES = dateTimePicker2.Value.ToString("yyyy-MM");
            DATES = DATES.Substring(2, 5);
            string NO = GETMAXNO(DATES);
            textBox2T1.Text = NO;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            INSERT_TB_DEV_CAKES(
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
            , textBox2T20.Text
            , textBox2T14.Text           
            , textBox2T15.Text
            , textBox2T16.Text
            , textBox2T17.Text
            , textBox2T18.Text
            , textBox2T19.Text

            );

            SEARCH_TB_DEV_CAKES2(textBox2T1.Text);
        }
        private void button7_Click(object sender, EventArgs e)
        {
            UPDATE_TB_DEV_CAKES(
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
                , textBox2T20.Text
                , textBox2T14.Text
                , textBox2T15.Text
                , textBox2T16.Text
                , textBox2T17.Text
                , textBox2T18.Text
                , textBox2T19.Text
                );

            SEARCH_TB_DEV_CAKES2(textBox2T1.Text);
        }
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TB_DEV_CAKES(textBox2T1.Text);
                SEARCH_TB_DEV_CAKES2(textBox2T1.Text);

                SETTEXT_TAB2();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2T14.Text.Trim()))
            {
                UPDATE_TB_DEV_CAKES_MB001(textBox2T1.Text.Trim(), textBox2T20.Text.Trim());
                SEARCH_TB_DEV_CAKES2(textBox2T1.Text);
            }
            else
            {
                MessageBox.Show("未填寫BOM品號");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            UPDATE_TB_DEV_CAKES_DETAILS(
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

            CAL_TB_DEV_CAKES_DETAILS(textBox2T1.Text);
            SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);
        }
        private void button10_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TB_DEV_CAKES_DETAILS(textBox2T40.Text);

                CAL_TB_DEV_CAKES_DETAILS(textBox2T1.Text);
                SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button14_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2T41.Text.Trim()))
            {
                UPDATE_TB_DEV_CAKES_DETAILS_MB001(textBox2T40.Text.Trim(), textBox2T41.Text.Trim());

                SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);
            }
            else
            {
                MessageBox.Show("未填寫BOM品號");
            }
        }
        private void button12_Click(object sender, EventArgs e)
        {
            CAL_TB_DEV_CAKES_DETAILS(textBox2T1.Text);
            SEARCH_TB_DEV_CAKES_DETAILS2(textBox2T1.Text);

            MessageBox.Show("完成");
        }

        #endregion


    }
}
