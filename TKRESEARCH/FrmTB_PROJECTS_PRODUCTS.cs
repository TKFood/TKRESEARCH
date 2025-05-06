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
using System.Xml;
using System.Xml.Linq;

namespace TKRESEARCH
{
    public partial class FrmTB_PROJECTS_PRODUCTS : Form
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

        public FrmTB_PROJECTS_PRODUCTS()
        {
            InitializeComponent();
         
        }


        #region FUNCTION
        private void FrmTB_PROJECTS_PRODUCTS_Load(object sender, EventArgs e)
        {
            comboBox1load();
            comboBox2load();
        }
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
                                WHERE [KIND]='TB_PROJECTS_PRODUCTS_ISCLOSED'
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
                                SELECT OWNER
                                FROM 
                                (
	                                SELECT '全部' AS 'OWNER'
	                                UNION ALL
	                                SELECT
	                                [OWNER]      
	                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
	                                GROUP BY [OWNER]
                                ) AS TEMP
                                ORDER BY OWNER
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("OWNER", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "OWNER";
            comboBox2.DisplayMember = "OWNER";
            sqlConn.Close();


        }
        public void SEARCH(string ISCLOSED, string OWNER,string PROJECTNAMES)
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

                StringBuilder QUERYS = new StringBuilder();
                StringBuilder QUERYS2 = new StringBuilder();
                StringBuilder QUERYS3 = new StringBuilder();


                sbSql.Clear();

                //DropDownList_ISCLOSED
                if (ISCLOSED.Equals("全部"))
                {
                    QUERYS.AppendFormat(@"");
                }
                else if (ISCLOSED.Equals("進行中"))
                {
                    QUERYS.AppendFormat(@" AND [ISCLOSED]='N' ");
                }
                else if (ISCLOSED.Equals("已完成"))
                {
                    QUERYS.AppendFormat(@" AND [ISCLOSED]='Y' ");
                }
                //DropDownList_OWNER
                if (OWNER.Equals("全部"))
                {
                    QUERYS2.AppendFormat(@"");
                }
                else
                {
                    QUERYS2.AppendFormat(@" AND OWNER=N'{0}' ", OWNER);
                }
                //TextBox1
                if (!string.IsNullOrEmpty(PROJECTNAMES))
                {
                    QUERYS3.AppendFormat(@" AND PROJECTNAMES LIKE '%{0}%' ", PROJECTNAMES);
                }
                else
                {
                    QUERYS3.AppendFormat(@"");
                }

                sbSql.AppendFormat(@"  
                                        SELECT 
                                        [NO] AS '專案編號'
                                        ,[KINDS] AS '分類'
                                        ,[PROJECTNAMES] AS '項目名稱'                                        
                                        ,[OWNER] AS '專案負責人'
                                        ,[STATUS] AS '研發進度回覆'
                                        ,[TASTESREPLYS] AS '業務進度回覆'
                                        ,[DESIGNER] AS '設計負責人'
                                        ,[DESIGNREPLYS] AS '設計回覆'
                                        ,[STAGES] AS '專案階段'
                                        ,[ISCLOSED] AS '是否結案'
                                        ,[DOC_NBR] AS '表單編號'
                                        ,CONVERT(NVARCHAR,[UPDATEDATES],112) AS '更新日'                                        
                                        ,[ID]
                                        
                                        FROM [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                        WHERE 1=1
                                        {0}
                                        {1}
                                        {2}
                                        ORDER BY [OWNER],[NO]
                                         ", QUERYS.ToString(), QUERYS2.ToString(), QUERYS3.ToString());

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
                textBox2.Text = row.Cells["研發進度回覆"].Value.ToString().Replace("\n", "\r\n"); 
                textBox3.Text = row.Cells["業務進度回覆"].Value.ToString().Replace("\n", "\r\n"); 
                textBox4.Text = row.Cells["設計回覆"].Value.ToString().Replace("\n", "\r\n");
                textBox5.Text = row.Cells["項目名稱"].Value.ToString().Replace("\n", "\r\n"); 

            }
        }

        public void UPDATE_TB_PROJECTS_PRODUCTS_COMMENTS(string ID,string STATUS,string TASTESREPLYS,string DESIGNREPLYS)
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    UPDATE  [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    SET [STATUS]=@STATUS,
                                    [TASTESREPLYS]=@TASTESREPLYS,
                                    [DESIGNREPLYS]=@DESIGNREPLYS
                                    WHERE [ID]=@ID
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值
                cmd.Parameters.AddWithValue("@ID", ID);
                cmd.Parameters.AddWithValue("@STATUS", STATUS);
                cmd.Parameters.AddWithValue("@TASTESREPLYS", TASTESREPLYS);
                cmd.Parameters.AddWithValue("@DESIGNREPLYS", DESIGNREPLYS);


                // 執行插入語句
                result = cmd.ExecuteNonQuery();

                // 處理交易
                if (result == 0)
                {
                    tran.Rollback();    // 交易取消
                }
                else
                {
                    tran.Commit();      // 執行交易
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

        public void SETTEXT()
        {
            textBoxid.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string ISCLOSED = comboBox1.Text.Trim();
            string OWNER = comboBox2.Text.Trim();
            string PROJECTNAMES = textBox1.Text.Trim();

            SEARCH(ISCLOSED, OWNER, PROJECTNAMES);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string ID = textBoxid.Text.Trim();
            string STATUS= textBox2.Text.Replace("\r\n", "\n"); 
            string TASTESREPLYS = textBox3.Text.Replace("\r\n", "\n");
            string DESIGNREPLYS = textBox4.Text.Replace("\r\n", "\n");
            UPDATE_TB_PROJECTS_PRODUCTS_COMMENTS(ID, STATUS, TASTESREPLYS, DESIGNREPLYS);

            string ISCLOSED = comboBox1.Text.Trim();
            string OWNER = comboBox2.Text.Trim();
            string PROJECTNAMES = textBox1.Text.Trim();

            SEARCH(ISCLOSED, OWNER, PROJECTNAMES);
        }

        #endregion


    }
}
