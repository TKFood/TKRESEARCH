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

        #endregion


    }
}
