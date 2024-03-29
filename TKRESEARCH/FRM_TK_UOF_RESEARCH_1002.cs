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
using System.Collections;
using TKITDLL;
using System.Xml;
using System.Text.RegularExpressions;

namespace TKRESEARCH
{
    public partial class FRM_TK_UOF_RESEARCH_1002 : Form
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


        public FRM_TK_UOF_RESEARCH_1002()
        {
            InitializeComponent();

            SET_DATES();
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
                                WHERE [KIND]='TK_UOF_RESEARCH_1002'
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
                                WHERE [KIND]='TK_UOF_RESEARCH_1002'
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
                                WHERE [KIND]='TK_UOF_RESEARCH_1002'
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

        public void SET_DATES()
        {
            textBox1.Text = DateTime.Now.ToString("yyyy");
            textBox3.Text = DateTime.Now.ToString("yyyy");
        }

        public void SEARCH(string YYYY, string RDFrm1002PD, string ISCLOSED)
        {

            string RDF1002SN = YYYY.Substring(2, 2);

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

                sbSql.Clear();
                SQLQUERY1.Clear();
                SQLQUERY2.Clear();
                SQLQUERY3.Clear();



                if (!string.IsNullOrEmpty(RDF1002SN))
                {
                    SQLQUERY1.AppendFormat(@"
                                        AND RDF1002SN LIKE '%{0}%'
                                        ", RDF1002SN);
                }
                else
                {
                    SQLQUERY1.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(RDFrm1002PD))
                {
                    SQLQUERY2.AppendFormat(@"
                                        AND RDFrm1002PD LIKE '%{0}%'
                                        ", RDFrm1002PD);
                }
                else
                {
                    SQLQUERY2.AppendFormat(@"");
                }
                if (!string.IsNullOrEmpty(ISCLOSED))
                {
                    SQLQUERY3.AppendFormat(@"
                                        AND ISCLOSED LIKE '%{0}%'
                                        ", ISCLOSED);
                }
                else
                {
                    SQLQUERY3.AppendFormat(@"");
                }



                sbSql.AppendFormat(@" 
                            
                            SELECT 
                            [RDF1002SN] AS '表單編號'
                            ,[NAME] AS '申請人'
                            ,[RDFrm1002DATE1] AS '預計設計須完成日(需求單位填寫)'
                            ,[RDFrm1002DATE2] AS '預計設計上校稿日(行銷單位填寫)'
                            ,[RDFrm1002CS] AS '設計別'
                            ,[RDFrm1002DP] AS '需求部門'
                            ,[RDFrm1002PD] AS '產品名稱'

                            ,[ISCLOSED] AS '是否結案'
                            ,[INPROCESSING] AS '處理進度'

                            ,[RDFrm1002ST] AS '產品規格' 
                            ,[RDFrm1002G7T1] AS '預計出貨日期'
                            ,[RDFrm1002G7T2] AS '預計上市日期'
                            ,[RDFrm1002G7T3] AS '預計銷售通路/國家別'
                            ,[RDFrm1002G7T4] AS '預估量（最小單位）'
                            ,[RDFrm1002G7T5] AS '商品屬性'
                            ,[RDFrm1002G5T6] AS '產品包裝形式'
                            ,[RDFrm1002DS] AS '設計需求具體內容'
                      
 
                            FROM [TKRESEARCH].[dbo].[TK_UOF_RESEARCH_1002]
                            WHERE 1=1
                            {0}
                             {1}
                             {2}
                            ORDER BY [RDF1002SN]

	 

                            ", SQLQUERY1.ToString(), SQLQUERY2.ToString(), SQLQUERY3.ToString());



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
        public void SETFASTREPORT(string YYYY,string RDFrm1002PD,string ISCLOSED)
        {
            string RDF1002SN = YYYY.Substring(2, 2);


            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL(RDF1002SN, RDFrm1002PD, ISCLOSED);      

            Report report1 = new Report();
            report1.Load(@"REPORT\13.研發類表單1002.設計需求內容清單.frx");

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
            
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();

        }
        public StringBuilder SETSQL(string YY, string RDFrm1002PD, string ISCLOSED)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQLQUERY1 = new StringBuilder();
            StringBuilder SQLQUERY2 = new StringBuilder();
            StringBuilder SQLQUERY3 = new StringBuilder();

            if (!string.IsNullOrEmpty(YY))
            {
                SQLQUERY1.AppendFormat(@"
                                        AND RDF1002SN LIKE '%{0}%'
                                        ", YY);
            }
            else
            {
                SQLQUERY1.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(RDFrm1002PD))
            {
                SQLQUERY2.AppendFormat(@"
                                        AND RDFrm1002PD LIKE '%{0}%'
                                        ", RDFrm1002PD);
            }
            else
            {
                SQLQUERY2.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(ISCLOSED))
            {
                SQLQUERY3.AppendFormat(@"
                                        AND ISCLOSED LIKE '%{0}%'
                                        ", ISCLOSED);
            }
            else
            {
                SQLQUERY3.AppendFormat(@"");
            }



            SB.AppendFormat(@" 
                            
                            SELECT 
                            [RDF1002SN] AS '表單編號'
                            ,[NAME] AS '申請人'
                            ,[RDFrm1002DATE1] AS '預計設計須完成日(需求單位填寫)'
                            ,[RDFrm1002DATE2] AS '預計設計上校稿日(行銷單位填寫)'
                            ,[RDFrm1002CS] AS '設計別'
                            ,[RDFrm1002DP] AS '需求部門'
                            ,[RDFrm1002PD] AS '產品名稱'
                            ,[RDFrm1002ST] AS '產品規格'
                            ,[RDFrm1002G7T1] AS '預計出貨日期'
                            ,[RDFrm1002G7T2] AS '預計上市日期'
                            ,[RDFrm1002G7T3] AS '預計銷售通路/國家別'
                            ,[RDFrm1002G7T4] AS '預估量（最小單位）'
                            ,[RDFrm1002G7T5] AS '商品屬性'
                            ,[RDFrm1002G5T6] AS '產品包裝形式'
                            ,[RDFrm1002DS] AS '設計需求具體內容'
                            ,[INPROCESSING] AS '處理進度'
                            ,[ISCLOSED] AS '是否結案'
 
                            FROM [TKRESEARCH].[dbo].[TK_UOF_RESEARCH_1002]
                            WHERE 1=1
                            {0}
                             {1}
                             {2}
                            ORDER BY [RDF1002SN]

	 

                            ", SQLQUERY1.ToString(), SQLQUERY2.ToString(), SQLQUERY3.ToString());


            return SB;
        }


        public void NEW_TKRESEARCH_TK_UOF_RESEARCH_1002()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            string THISYEARS = DateTime.Now.ToString("yyyy");
            //取西元年後2位
            THISYEARS = THISYEARS.Substring(2, 2);
            //THISYEARS = "21";
            string THISYEARSDAYS = DateTime.Now.ToString("yyyy") + "0101";

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

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



                //核準過TASK_RESULT='0'
                //AND DOC_NBR  LIKE 'QC1002{0}%'

                sbSql.AppendFormat(@"  
                                    SELECT DOC_NBR,*
                                    FROM [UOF].dbo.TB_WKF_TASK
                                    WHERE 1=1
                                    AND TASK_STATUS='2'
                                    AND TASK_RESULT='0'
                                    AND DOC_NBR  LIKE 'RD1002%'
                                    AND DOC_NBR >='RD1002230400001'
                                    AND DOC_NBR COLLATE Chinese_Taiwan_Stroke_BIN NOT IN (SELECT  [RDF1002SN] FROM [192.168.1.105].[TKRESEARCH].[dbo].[TK_UOF_RESEARCH_1002])
                                       
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    foreach (DataRow dr in ds1.Tables["ds1"].Rows)
                    {
                        SEARCHUOFTB_WKF_TASK_RD1002(dr["DOC_NBR"].ToString());
                    }
                }
                else
                {

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

        //找出UOF表單的資料，將CURRENT_DOC的內容，轉成xmlDoc
        //從xmlDoc找出各節點的Attributes
        public void SEARCHUOFTB_WKF_TASK_RD1002(string DOC_NBR)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

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

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                    SELECT * 
                                    FROM [UOF].DBO.TB_WKF_TASK 
                                    LEFT JOIN [UOF].[dbo].[TB_EB_USER] ON [TB_EB_USER].USER_GUID=TB_WKF_TASK.USER_GUID
                                    WHERE DOC_NBR LIKE '{0}%'
                              
                                    ", DOC_NBR);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    string RDF1002SN = "";
                    string NAME = "";
                    string RDFrm1002DATE1 = "";
                    string RDFrm1002DATE2 = "";
                    string RDFrm1002CS = "";
                    string RDFrm1002DP = "";
                    string RDFrm1002PD = "";
                    string RDFrm1002ST = "";
                    string RDFrm1002DS = "";

                    string RDFrm1002G7T1 = "";
                    string RDFrm1002G7T2 = "";
                    string RDFrm1002G7T3 = "";
                    string RDFrm1002G7T4 = "";
                    string RDFrm1002G7T5 = "";
                    string RDFrm1002G5T6 = "";



                    XmlDocument xmlDoc = new XmlDocument();

                    xmlDoc.LoadXml(ds1.Tables["ds1"].Rows[0]["CURRENT_DOC"].ToString());



                    //XmlNode node = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='ID']");
                    try
                    {
                        RDF1002SN = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDF1002SN']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        //找出表單申請人 
                        NAME = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002DATE1']").Attributes["fillerName"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002DATE1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002DATE1']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002DATE2 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002DATE2']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002CS = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002CS']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002DP = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002DP']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002PD = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002PD']").Attributes["fieldValue"].Value;
                    }
                    catch { }
                    try
                    {
                        RDFrm1002ST = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002ST']").Attributes["fieldValue"].Value;
                    }
                    catch { }

                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002DS']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"&#xD;", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"&#xA;", "");

                        RDFrm1002DS = fieldValue3;
                    }
                    catch { }

                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G7T1']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"&#xD;", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"&#xA;", "");

                        RDFrm1002G7T1 = fieldValue3;
                    }
                    catch { }
                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G7T2']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"[\W_]+", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"[0-9A-Za-z]+", "");

                        RDFrm1002G7T2 = fieldValue3;
                    }
                    catch { }
                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G7T3']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"[\W_]+", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"[0-9A-Za-z]+", "");

                        RDFrm1002G7T3 = fieldValue3;
                    }
                    catch { }
                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G7T4']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"[\W_]+", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"[0-9A-Za-z]+", "");

                        RDFrm1002G7T4 = fieldValue3;
                    }
                    catch { }
                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G7T5']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"[\W_]+", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"[0-9A-Za-z]+", "");

                        RDFrm1002G7T5 = fieldValue3;
                    }
                    catch { }
                    try
                    {
                        //把html語法去除 
                        //QCFrm002Cmf = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='QCFrm002Cmf']").Attributes["fieldValue"].Value;

                        string fieldValue1 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='RDFrm1002G7']/DataGrid/Row/Cell[@fieldId='RDFrm1002G5T6']").Attributes["fieldValue"].Value;

                        string fieldValue2 = Regex.Replace(fieldValue1, @"[\W_]+", "");
                        string fieldValue3 = Regex.Replace(fieldValue2, @"[0-9A-Za-z]+", "");

                        RDFrm1002G5T6 = fieldValue3;
                    }
                    catch { }
                    //string OK = "";
                    ADD_TK_UOF_RESEARCH_1002(
                                             RDF1002SN
                                            , NAME
                                            , RDFrm1002DATE1
                                            , RDFrm1002DATE2
                                            , RDFrm1002CS
                                            , RDFrm1002DP
                                            , RDFrm1002PD
                                            , RDFrm1002ST
                                            , RDFrm1002G7T1
                                            , RDFrm1002G7T2
                                            , RDFrm1002G7T3
                                            , RDFrm1002G7T4
                                            , RDFrm1002G7T5
                                            , RDFrm1002G5T6
                                            , RDFrm1002DS
                                           );


                }
                else
                {

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


        public void ADD_TK_UOF_RESEARCH_1002(
                                            string RDF1002SN
                                            , string NAME
                                            , string RDFrm1002DATE1
                                            , string RDFrm1002DATE2
                                            , string RDFrm1002CS
                                            , string RDFrm1002DP
                                            , string RDFrm1002PD
                                            , string RDFrm1002ST
                                            , string RDFrm1002G7T1
                                            , string RDFrm1002G7T2
                                            , string RDFrm1002G7T3
                                            , string RDFrm1002G7T4
                                            , string RDFrm1002G7T5
                                            , string RDFrm1002G5T6
                                            , string RDFrm1002DS
                                            )
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

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
                                    INSERT INTO [TKRESEARCH].[dbo].[TK_UOF_RESEARCH_1002]
                                    (
                                    [RDF1002SN]
                                    ,[NAME]
                                    ,[RDFrm1002DATE1]
                                    ,[RDFrm1002DATE2]
                                    ,[RDFrm1002CS]
                                    ,[RDFrm1002DP]
                                    ,[RDFrm1002PD]
                                    ,[RDFrm1002ST]
                                    ,[RDFrm1002G7T1]
                                    ,[RDFrm1002G7T2]
                                    ,[RDFrm1002G7T3]
                                    ,[RDFrm1002G7T4]
                                    ,[RDFrm1002G7T5]
                                    ,[RDFrm1002G5T6]
                                    ,[RDFrm1002DS]
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
                                    )
                                    ", RDF1002SN
                                    , NAME
                                    , RDFrm1002DATE1
                                    , RDFrm1002DATE2
                                    , RDFrm1002CS
                                    , RDFrm1002DP
                                    , RDFrm1002PD
                                    , RDFrm1002ST
                                    , RDFrm1002G7T1
                                    , RDFrm1002G7T2
                                    , RDFrm1002G7T3
                                    , RDFrm1002G7T4
                                    , RDFrm1002G7T5
                                    , RDFrm1002G5T6
                                    , RDFrm1002DS);

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SET_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox5.Text = row.Cells["表單編號"].Value.ToString();
                    textBox6.Text = row.Cells["申請人"].Value.ToString();
                    textBox7.Text = row.Cells["預計設計須完成日(需求單位填寫)"].Value.ToString();
                    textBox8.Text = row.Cells["預計設計上校稿日(行銷單位填寫)"].Value.ToString();
                    textBox9.Text = row.Cells["設計別"].Value.ToString();
                    textBox10.Text = row.Cells["需求部門"].Value.ToString();
                    textBox11.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox12.Text = row.Cells["產品規格"].Value.ToString();
                    textBox13.Text = row.Cells["預計出貨日期"].Value.ToString();
                    textBox14.Text = row.Cells["預計上市日期"].Value.ToString();
                    textBox15.Text = row.Cells["預計銷售通路/國家別"].Value.ToString();
                    textBox16.Text = row.Cells["預估量（最小單位）"].Value.ToString();
                    textBox17.Text = row.Cells["商品屬性"].Value.ToString();
                    textBox18.Text = row.Cells["產品包裝形式"].Value.ToString();
                    textBox19.Text = row.Cells["設計需求具體內容"].Value.ToString();
                    textBox20.Text = row.Cells["處理進度"].Value.ToString();
                    comboBox3.Text = row.Cells["是否結案"].Value.ToString();

                }
                else
                {
                    
                }
            }

        }

        public void SET_NULL()
        {
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
            textBox20.Text = null;

         
        }

        public void UPDATE_TK_UOF_RESEARCH_1002(string RDF1002SN,string INPROCESSING,string ISCLOSED)
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
                                        UPDATE [TKRESEARCH].[dbo].[TK_UOF_RESEARCH_1002]
                                        SET [INPROCESSING]='{1}',[ISCLOSED]='{2}'
                                        WHERE [RDF1002SN]='{0}'

                                        ", RDF1002SN, INPROCESSING, ISCLOSED);

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
            SETFASTREPORT(textBox1.Text,textBox2.Text,comboBox1.Text.ToString());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH(textBox3.Text.Trim(), textBox4.Text, comboBox2.Text.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            NEW_TKRESEARCH_TK_UOF_RESEARCH_1002();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            UPDATE_TK_UOF_RESEARCH_1002(textBox5.Text,textBox20.Text,comboBox3.Text);
            SEARCH(textBox3.Text.Trim(), textBox4.Text, comboBox2.Text.ToString());
        }

        #endregion


    }
}
