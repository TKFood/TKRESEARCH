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

        public FrmTBINVMB()
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
                else
                {
                    SQLQUERY1.AppendFormat(@"");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(comboBox1.Text.ToString(),textBox1.Text.ToString(), textBox2.Text.ToString(),textBox3.Text.ToString());
        }

        #endregion


    }
}
