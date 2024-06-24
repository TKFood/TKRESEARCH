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
    public partial class FrmTB_DEV_NEWLISTS : Form
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
   


        public FrmTB_DEV_NEWLISTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string yyyyMM, string NAMES)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            string YY = yyyyMM.Substring(2, 2);
            string MM = yyyyMM.Substring(4, 2);
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

                if (!string.IsNullOrEmpty(NAMES))
                {
                    SQLquery1.AppendFormat(@" AND [NAMES] LIKE '%{0}%' ", NAMES);
                }
                else
                {
                    SQLquery1.AppendFormat(@" ");

                    if(!string.IsNullOrEmpty(yyyyMM))
                    {
                        SQLquery2.AppendFormat(@" AND [NO] LIKE '%{0}%'", YY+"-"+ MM);
                    }
                }
               

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [NO] AS '編號'
                                    ,[NAMES] AS '商品'
                                    ,[SPECS] AS '規格'
                                    ,[COMMENTS] AS '需求'
                                    ,[INGREDIENTS] AS '成份'
                                    ,[COSTS] AS '成本'
                                    ,[MOQS] AS 'MOQ'
                                    ,[MANUPRODS] AS '一天產能量'
                                    ,CONVERT(NVARCHAR,[GETDATES],112)  AS '送樣日期'
                                    ,[REPLY] AS '業務回覆'
                                    ,CONVERT(NVARCHAR,[CARESTEDATES],112) AS '建立日期'
                                    ,[ID]

                                    FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
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
            SETTEXT();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                DataGridViewRow row = dataGridView1.Rows[rowindex];

                textBoxid.Text = row.Cells["ID"].Value.ToString();
                textBox2.Text = row.Cells["編號"].Value.ToString();
                textBox3.Text = row.Cells["商品"].Value.ToString();
                textBox4.Text = row.Cells["規格"].Value.ToString();
                textBox5.Text = row.Cells["需求"].Value.ToString();
                textBox6.Text = row.Cells["成份"].Value.ToString();
                textBox7.Text = row.Cells["成本"].Value.ToString();
                textBox8.Text = row.Cells["MOQ"].Value.ToString();
                textBox9.Text = row.Cells["一天產能量"].Value.ToString();
                textBox10.Text = row.Cells["業務回覆"].Value.ToString();

                //dateTimePicker2.Value= row.Cells["開發日期"].Value.ToString();
                DateTime dateTime2;
                if (DateTime.TryParseExact(row.Cells["建立日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime2))
                {                 
                    dateTimePicker2.Value = dateTime2;
                }
                DateTime dateTime3;
                if (DateTime.TryParseExact(row.Cells["送樣日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime3))
                {                   
                    dateTimePicker3.Value = dateTime3;
                }


            }
        }

        public void SETTEXT()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBoxid.Text = null;
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
                                        FROM  [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMM"),textBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string DATES = DateTime.Now.ToString("yyyy-MM");
            DATES = DATES.Substring(2, 5);
            string NO = GETMAXNO(DATES);
            textBox2.Text = NO;
        }
        private void button5_Click(object sender, EventArgs e)
        {

        }
        private void button6_Click(object sender, EventArgs e)
        {

        }
        private void button7_Click(object sender, EventArgs e)
        {

        }



        #endregion

       
    }
}
