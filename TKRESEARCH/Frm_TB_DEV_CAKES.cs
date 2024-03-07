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
        }

        #region FUNCTION
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_TB_DEV_CAKES(dateTimePicker1.Value.ToString("yyyy"), textBox1.Text);
        }
        #endregion

    
    }
}
