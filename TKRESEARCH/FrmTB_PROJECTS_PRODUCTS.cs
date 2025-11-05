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
            comboBox1_load();
            comboBox2_load();
            comboBox3_load();
            comboBox4_load();
            comboBox5_load();
            comboBox6_load();
            comboBox7_load();
            comboBox8_load();
            comboBox9_load();
            comboBox10_load();
            comboBox11_load();
            comboBox12_load();
            comboBox13_load();
            comboBox14_load();

            SETTEXT_ADD();
        }
        public void comboBox1_load()
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

        public void comboBox2_load()
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
                                OWNER
                                FROM
                                (
                                    SELECT 
                                        '全部' AS OWNER,      
                                        0 AS SortOrder
                                    UNION ALL
                                    SELECT
                                        [OWNER],       
                                        1 AS SortOrder
                                    FROM 
                                        [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    GROUP BY 
                                        [OWNER]
                                ) AS TEMP
                                ORDER BY 
                                    TEMP.SortOrder, -- 先按 SortOrder 排序，確保 0 (全部) 在前
                                    TEMP.OWNER;     -- 再按 OWNER 字母順序排序 (對 '全部' 之後的項目生效)
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
        public void comboBox3_load()
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
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "PARANAME";
            comboBox3.DisplayMember = "PARANAME";
            sqlConn.Close();


        }

        public void comboBox4_load()
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
                                OWNER
                                FROM
                                (
                                    SELECT 
                                        '全部' AS OWNER,      
                                        0 AS SortOrder
                                    UNION ALL
                                    SELECT
                                        [OWNER],       
                                        1 AS SortOrder
                                    FROM 
                                        [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    GROUP BY 
                                        [OWNER]
                                ) AS TEMP
                                ORDER BY 
                                    TEMP.SortOrder, -- 先按 SortOrder 排序，確保 0 (全部) 在前
                                    TEMP.OWNER;     -- 再按 OWNER 字母順序排序 (對 '全部' 之後的項目生效)
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("OWNER", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "OWNER";
            comboBox4.DisplayMember = "OWNER";
            sqlConn.Close();


        }
        public void comboBox5_load()
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
                                [KINDS]
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_KINDS]
                                ORDER BY [ID]
                                 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KINDS", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "KINDS";
            comboBox5.DisplayMember = "KINDS";
            sqlConn.Close();


        }
        public void comboBox6_load()
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
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "OWNER";
            comboBox6.DisplayMember = "OWNER";
            sqlConn.Close();


        }
        public void comboBox7_load()
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
                                SELECT DESIGNER
                                FROM 
                                (                                
                                SELECT
                                DESIGNER      
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                GROUP BY DESIGNER
                                ) AS TEMP
                                ORDER BY DESIGNER
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("DESIGNER", typeof(string));
            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "DESIGNER";
            comboBox7.DisplayMember = "DESIGNER";
            sqlConn.Close();


        }
        public void comboBox8_load()
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
                                ,[STAGES]
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_STAGES]
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("STAGES", typeof(string));
            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "STAGES";
            comboBox8.DisplayMember = "STAGES";
            sqlConn.Close();


        }
        public void comboBox9_load()
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
                                FROM[TKRESEARCH].[dbo].[TBPARA]
                                WHERE[KIND] = 'TB_PROJECTS_PRODUCTS_ISCLOSEDYN'
                                ORDER BY[PARAID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox9.DataSource = dt.DefaultView;
            comboBox9.ValueMember = "PARANAME";
            comboBox9.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void comboBox10_load()
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
                                [KINDS]
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_KINDS]
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KINDS", typeof(string));
            da.Fill(dt);
            comboBox10.DataSource = dt.DefaultView;
            comboBox10.ValueMember = "KINDS";
            comboBox10.DisplayMember = "KINDS";
            sqlConn.Close();


        }
        public void comboBox11_load()
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
            comboBox11.DataSource = dt.DefaultView;
            comboBox11.ValueMember = "OWNER";
            comboBox11.DisplayMember = "OWNER";
            sqlConn.Close();


        }
        public void comboBox12_load()
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
                                SELECT DESIGNER
                                FROM 
                                (                                
                                SELECT
                                DESIGNER      
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                GROUP BY DESIGNER
                                ) AS TEMP
                                ORDER BY DESIGNER
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("DESIGNER", typeof(string));
            da.Fill(dt);
            comboBox12.DataSource = dt.DefaultView;
            comboBox12.ValueMember = "DESIGNER";
            comboBox12.DisplayMember = "DESIGNER";
            sqlConn.Close();


        }
        public void comboBox13_load()
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
                                ,[STAGES]
                                FROM [TKRESEARCH].[dbo].[TB_PROJECTS_STAGES]
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("STAGES", typeof(string));
            da.Fill(dt);
            comboBox13.DataSource = dt.DefaultView;
            comboBox13.ValueMember = "STAGES";
            comboBox13.DisplayMember = "STAGES";
            sqlConn.Close();


        }
        public void comboBox14_load()
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
                                FROM[TKRESEARCH].[dbo].[TBPARA]
                                WHERE[KIND] = 'TB_PROJECTS_PRODUCTS_ISCLOSEDYN'
                                ORDER BY[PARAID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox14.DataSource = dt.DefaultView;
            comboBox14.ValueMember = "PARANAME";
            comboBox14.DisplayMember = "PARANAME";
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
                                        ,[PURREPLYS] AS '採購回覆'
                                        ,[QCREPLYS] AS '品保回覆'
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
                        //dataGridView1.AutoResizeColumns();
                        // 1. 設定預設的單元格樣式，啟用文字換行 (WrapMode)
                        // 這會影響到整個 DataGridView 中所有單元格的文本顯示方式。
                        dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        // 2. 設定行高自動調整模式 (AutoSizeRowsMode)
                        // 這會告訴 DataGridView 根據單元格的內容（特別是換行後的內容）來自動調整行高。
                        dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        // 3. 讓欄寬完全依照所有單元格的內容來決定
                        // **注意：** 這可能會導致 DataGridView 總寬度超出螢幕或容器邊界，出現水平滾動條。
                        //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                        // 隱藏  欄位
                        dataGridView1.Columns["研發進度回覆"].Visible = false;
                        dataGridView1.Columns["業務進度回覆"].Visible = false;
                        dataGridView1.Columns["設計回覆"].Visible = false;
                        dataGridView1.Columns["採購回覆"].Visible = false;
                        dataGridView1.Columns["品保回覆"].Visible = false;

                        if (dataGridView1.Columns.Contains("項目名稱"))
                        {
                            dataGridView1.Columns["項目名稱"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }
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

                HandleRowSelectionChanged_GV1(dataGridView1.CurrentRow);
            }
        }
        ///把原本的 SelectionChanged 抽出來變一個方法
        private void HandleRowSelectionChanged_GV1(DataGridViewRow row)
        {
            // 原本的內容寫這裡，例如填入欄位的值到 TextBox 等
            textBoxid.Text = row.Cells["ID"].Value.ToString();
            textBox2.Text = row.Cells["研發進度回覆"].Value.ToString().Replace("\n", "\r\n");
            textBox3.Text = row.Cells["業務進度回覆"].Value.ToString().Replace("\n", "\r\n");
            textBox4.Text = row.Cells["設計回覆"].Value.ToString().Replace("\n", "\r\n");
            textBox5.Text = row.Cells["項目名稱"].Value.ToString().Replace("\n", "\r\n");
            textBox16.Text = row.Cells["品保回覆"].Value.ToString().Replace("\n", "\r\n");
            textBox18.Text = row.Cells["採購回覆"].Value.ToString().Replace("\n", "\r\n");

        }

        //重新查詢資料後，找到對應的那筆資料並選取該列
        private void ReloadDataAndRestoreSelection_GV1(string ID)
        {
            // 查詢後嘗試還原選取列
            DataGridView SEARCH_DataGridView = new DataGridView();
            SEARCH_DataGridView = dataGridView1;

            foreach (DataGridViewRow row in SEARCH_DataGridView.Rows)
            {
                if (row.Cells["ID"].Value.ToString() == ID)
                {
                    row.Selected = true;
                    SEARCH_DataGridView.CurrentCell = row.Cells[0]; // 確保游標在那一列
                    SEARCH_DataGridView.FirstDisplayedScrollingRowIndex = row.Index;

                    // 手動呼叫 SelectionChanged 的處理邏輯
                    HandleRowSelectionChanged_GV1(row); // ← 把原本的 SelectionChanged 抽出來變一個方法
                    break;
                }
            }

            SEARCH_DataGridView.Refresh(); // 強制畫面刷新
        }

        public void SEARCH_GV2(string ISCLOSED, string OWNER, string PROJECTNAMES)
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            SETTEXT2();

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                DataGridViewRow row = dataGridView2.Rows[rowindex];
                HandleRowSelectionChanged_GV2(row);

            }
        }
        ///把原本的 SelectionChanged 抽出來變一個方法
        private void HandleRowSelectionChanged_GV2(DataGridViewRow row)
        {
            textBoxid2.Text = row.Cells["ID"].Value.ToString();
            textBox7.Text = row.Cells["專案編號"].Value.ToString().Replace("\n", "\r\n");
            textBox9.Text = row.Cells["項目名稱"].Value.ToString().Replace("\n", "\r\n");
            textBox10.Text = row.Cells["表單編號"].Value.ToString().Replace("\n", "\r\n");

            comboBox5.Text = row.Cells["分類"].Value.ToString();
            comboBox6.Text = row.Cells["專案負責人"].Value.ToString();
            comboBox7.Text = row.Cells["設計負責人"].Value.ToString();
            comboBox8.Text = row.Cells["專案階段"].Value.ToString();
            comboBox9.Text = row.Cells["是否結案"].Value.ToString();

        }

        //重新查詢資料後，找到對應的那筆資料並選取該列
        private void ReloadDataAndRestoreSelection_GV2(string ID)
        {
            // 查詢後嘗試還原選取列
            DataGridView SEARCH_DataGridView = new DataGridView();
            SEARCH_DataGridView = dataGridView2;

            foreach (DataGridViewRow row in SEARCH_DataGridView.Rows)
            {
                if (row.Cells["ID"].Value.ToString() == ID)
                {
                    row.Selected = true;
                    SEARCH_DataGridView.CurrentCell = row.Cells[0]; // 確保游標在那一列
                    SEARCH_DataGridView.FirstDisplayedScrollingRowIndex = row.Index;

                    // 手動呼叫 SelectionChanged 的處理邏輯
                    HandleRowSelectionChanged_GV2(row); // ← 把原本的 SelectionChanged 抽出來變一個方法
                    break;
                }
            }

            SEARCH_DataGridView.Refresh(); // 強制畫面刷新
        }
        public void UPDATE_TB_PROJECTS_PRODUCTS_COMMENTS(
            string ID,
            string STATUS,
            string TASTESREPLYS,
            string DESIGNREPLYS,
            string UPDATEDATES,
            string QCREPLYS,
            string PURREPLYS
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
                                    [DESIGNREPLYS]=@DESIGNREPLYS,
                                    [UPDATEDATES]=@UPDATEDATES,
                                    [QCREPLYS]=@QCREPLYS,
                                    [PURREPLYS]=@PURREPLYS
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
                cmd.Parameters.AddWithValue("@UPDATEDATES", UPDATEDATES);
                cmd.Parameters.AddWithValue("@QCREPLYS", QCREPLYS);
                cmd.Parameters.AddWithValue("@PURREPLYS", PURREPLYS);

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
            catch (Exception ex)
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATE_TB_PROJECTS_PRODUCTS_ALL(
             string ID,
             string NO ,
             string PROJECTNAMES,
             string KINDS,
             string OWNER,
             string DESIGNER,
             string STAGES,
             string ISCLOSED,
             string DOC_NBR,
             string UPDATEDATES
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    UPDATE  [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    SET 
                                    [NO]=@NO,
                                    [PROJECTNAMES]=@PROJECTNAMES,
                                    [KINDS]=@KINDS,
                                    [OWNER]=@OWNER,
                                    [DESIGNER]=@DESIGNER,
                                    [STAGES]=@STAGES,
                                    [ISCLOSED]=@ISCLOSED,
                                    [DOC_NBR]=@DOC_NBR,
                                    [UPDATEDATES]=@UPDATEDATES
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
                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@PROJECTNAMES", PROJECTNAMES);
                cmd.Parameters.AddWithValue("@KINDS", KINDS);
                cmd.Parameters.AddWithValue("@OWNER", OWNER);
                cmd.Parameters.AddWithValue("@DESIGNER", DESIGNER);
                cmd.Parameters.AddWithValue("@STAGES", STAGES);
                cmd.Parameters.AddWithValue("@ISCLOSED", ISCLOSED);
                cmd.Parameters.AddWithValue("@DOC_NBR", DOC_NBR);
                cmd.Parameters.AddWithValue("@UPDATEDATES", UPDATEDATES);



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
            catch (Exception ex)
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_TB_PROJECTS_PRODUCTS(
            string NO,
            string PROJECTNAMES,
            string KINDS,
            string OWNER,
            string DESIGNER,
            string STAGES,
            string ISCLOSED,
            string DOC_NBR,
            string STATUS,
            string TASTESREPLYS,
            string DESIGNREPLYS,
            string QCREPLYS
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    (
                                    NO,
                                    PROJECTNAMES,
                                    KINDS,
                                    OWNER,
                                    DESIGNER,
                                    STAGES,
                                    ISCLOSED,
                                    DOC_NBR,
                                    STATUS,
                                    TASTESREPLYS,
                                    DESIGNREPLYS,
                                    QCREPLYS
                                    )
                                    VALUES
                                    (
                                    @NO,
                                    @PROJECTNAMES,
                                    @KINDS,
                                    @OWNER,
                                    @DESIGNER,
                                    @STAGES,
                                    @ISCLOSED,
                                    @DOC_NBR,
                                    @STATUS,
                                    @TASTESREPLYS,
                                    @DESIGNREPLYS,
                                    @QCREPLYS
                                    )

                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值             
                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@PROJECTNAMES", PROJECTNAMES);
                cmd.Parameters.AddWithValue("@KINDS", KINDS);
                cmd.Parameters.AddWithValue("@OWNER", OWNER);
                cmd.Parameters.AddWithValue("@DESIGNER", DESIGNER);
                cmd.Parameters.AddWithValue("@STAGES", STAGES);
                cmd.Parameters.AddWithValue("@ISCLOSED", ISCLOSED);
                cmd.Parameters.AddWithValue("@DOC_NBR", DOC_NBR);
                cmd.Parameters.AddWithValue("@STATUS", STATUS);
                cmd.Parameters.AddWithValue("@TASTESREPLYS", TASTESREPLYS);
                cmd.Parameters.AddWithValue("@DESIGNREPLYS", DESIGNREPLYS);
                cmd.Parameters.AddWithValue("@QCREPLYS", QCREPLYS);

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
            catch (Exception ex)
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_TB_PROJECTS_PRODUCTS_HISTORYS_COMMENTS(
            string ID,
            string STATUS,
            string TASTESREPLYS,
            string DESIGNREPLYS,
            string QCREPLYS,
            string PURREPLYS
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS_HISTORYS]
                                    (
                                    [SID],                                  
                                    [STATUS],
                                    [TASTESREPLYS],
                                    [DESIGNREPLYS],
                                    [QCREPLYS],
                                    [PURREPLYS]
                                    )
                                    VALUES
                                    (
                                    @ID,                                   
                                    @STATUS,
                                    @TASTESREPLYS,
                                    @DESIGNREPLYS,
                                    @QCREPLYS,
                                    @PURREPLYS
                                    )

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
                cmd.Parameters.AddWithValue("@QCREPLYS", QCREPLYS);
                cmd.Parameters.AddWithValue("@PURREPLYS", PURREPLYS);


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
            catch (Exception ex)
            {
                //MessageBox.Show("錯誤：" + ex.Message);
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_TB_PROJECTS_PRODUCTS_HISTORYS_ALL(
                string ID,
                string NO,
                string PROJECTNAMES,
                string KINDS,
                string OWNER,
                string DESIGNER,
                string STAGES,
                string ISCLOSED,
                string DOC_NBR,
                string UPDATEDATES
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS_HISTORYS]
                                    (
                                    [SID],                               
                                    [NO],
                                    [PROJECTNAMES],
                                    [KINDS],
                                    [OWNER],
                                    [DESIGNER],
                                    [STAGES],
                                    [ISCLOSED],
                                    [DOC_NBR]
                                  
                                    )
                                    VALUES
                                    (
                                    @ID,                                   
                                    @NO,
                                    @PROJECTNAMES,
                                    @KINDS,
                                    @OWNER,
                                    @DESIGNER,
                                    @STAGES,
                                    @ISCLOSED,
                                    @DOC_NBR                                  
                                    )

                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值             
                cmd.Parameters.AddWithValue("@ID", ID);
                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@PROJECTNAMES", PROJECTNAMES);
                cmd.Parameters.AddWithValue("@KINDS", KINDS);
                cmd.Parameters.AddWithValue("@OWNER", OWNER);
                cmd.Parameters.AddWithValue("@DESIGNER", DESIGNER);
                cmd.Parameters.AddWithValue("@STAGES", STAGES);
                cmd.Parameters.AddWithValue("@ISCLOSED", ISCLOSED);
                cmd.Parameters.AddWithValue("@DOC_NBR", DOC_NBR);
              


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
            catch (Exception ex)
            {
                //MessageBox.Show("錯誤：" + ex.Message);
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDTE_TB_PROJECTS_PRODUCTS_HISTORYS(
            string ID
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    UPDATE [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS_HISTORYS]
                                    SET [TB_PROJECTS_PRODUCTS_HISTORYS].NO=[TB_PROJECTS_PRODUCTS].NO,
                                    [TB_PROJECTS_PRODUCTS_HISTORYS].PROJECTNAMES=[TB_PROJECTS_PRODUCTS].PROJECTNAMES
                                    FROM  [TKRESEARCH].[dbo].[TB_PROJECTS_PRODUCTS]
                                    WHERE [TB_PROJECTS_PRODUCTS].ID=[TB_PROJECTS_PRODUCTS_HISTORYS].SID
                                    AND ISNULL([TB_PROJECTS_PRODUCTS_HISTORYS].NO,'')=''
                                    AND [TB_PROJECTS_PRODUCTS_HISTORYS].SID=@ID
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值             
                cmd.Parameters.AddWithValue("@ID", ID);
               



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
            catch (Exception ex)
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

        public void SETTEXT2()
        {
            textBoxid2.Text = "";
            textBox7.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
          

        }
        public void SETTEXT_ADD()
        {
            textBox13.Text = 
                "1.樣本提供:" + "\r\n" +
                "2.成本提供(初算/正式):" + "\r\n" +
                "3.特殊原料、製程批量說明:";
            textBox14.Text =
                "1.試吃確認:" + "\r\n" +
                "2.報價確認:" + "\r\n" +
                "3.進度更新:";
            textBox15.Text =
                "1.圖面設計:" + "\r\n" +
                "2.上校稿:" + "\r\n" +
                "3.廠商確稿發包:";

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
            string QCREPLYS = textBox16.Text.Replace("\r\n", "\n");
            string UPDATEDATES = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string PURREPLYS = textBox18.Text.Replace("\r\n", "\n");

            UPDATE_TB_PROJECTS_PRODUCTS_COMMENTS(ID, STATUS, TASTESREPLYS, DESIGNREPLYS, UPDATEDATES, QCREPLYS, PURREPLYS);
            ADD_TB_PROJECTS_PRODUCTS_HISTORYS_COMMENTS(ID, STATUS, TASTESREPLYS, DESIGNREPLYS, QCREPLYS, PURREPLYS);
            UPDTE_TB_PROJECTS_PRODUCTS_HISTORYS(ID);

            string ISCLOSED = comboBox1.Text.Trim();
            string OWNER = comboBox2.Text.Trim();
            string PROJECTNAMES = textBox1.Text.Trim();

            SEARCH(ISCLOSED, OWNER, PROJECTNAMES);

            ReloadDataAndRestoreSelection_GV1(ID);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string ISCLOSED = comboBox3.Text.Trim();
            string OWNER = comboBox4.Text.Trim();
            string PROJECTNAMES = textBox6.Text.Trim();

            SEARCH_GV2(ISCLOSED, OWNER, PROJECTNAMES);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string ID = textBoxid2.Text.Trim();
            string NO = textBox7.Text.Trim();
            string PROJECTNAMES = textBox9.Text.Trim();
            string KINDS = comboBox5.Text.ToString();
            string OWNER = comboBox6.Text.ToString();            
            string DESIGNER = comboBox7.Text.ToString();
            string STAGES = comboBox8.Text.ToString();
            string ISCLOSED = comboBox9.Text.ToString();
            string DOC_NBR = textBox10.Text.Trim();
            string UPDATEDATES = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            UPDATE_TB_PROJECTS_PRODUCTS_ALL(
                ID,
                NO,
                PROJECTNAMES,
                KINDS,
                OWNER,
                DESIGNER,
                STAGES,
                ISCLOSED,
                DOC_NBR,
                UPDATEDATES
            );

            ADD_TB_PROJECTS_PRODUCTS_HISTORYS_ALL(
                ID,
                NO,
                PROJECTNAMES,
                KINDS,
                OWNER,
                DESIGNER,
                STAGES,
                ISCLOSED,
                DOC_NBR,
                UPDATEDATES);
            UPDTE_TB_PROJECTS_PRODUCTS_HISTORYS(ID);


            string SEARCH_ISCLOSED = comboBox3.Text.Trim();
            string SEARCH_OWNER = comboBox4.Text.Trim();
            string SEARCH_PROJECTNAMES = textBox6.Text.Trim();

            SEARCH_GV2(SEARCH_ISCLOSED, SEARCH_OWNER, SEARCH_PROJECTNAMES);
            ReloadDataAndRestoreSelection_GV2(ID);
        }


        #endregion

        private void button5_Click(object sender, EventArgs e)
        {
            string NO = textBox8.Text.Trim();
            string PROJECTNAMES = textBox11.Text.Trim();
            string KINDS = comboBox10.Text.ToString();
            string OWNER = comboBox11.Text.ToString();
            string DESIGNER = comboBox12.Text.ToString();
            string STAGES = comboBox13.Text.ToString();
            string ISCLOSED = comboBox14.Text.ToString();
            string DOC_NBR = textBox10.Text.Trim();
            string STATUS = textBox13.Text.Replace("\r\n", "\n");
            string TASTESREPLYS = textBox14.Text.Replace("\r\n", "\n");
            string DESIGNREPLYS = textBox15.Text.Replace("\r\n", "\n");
            string QCREPLYS = textBox17.Text.Replace("\r\n", "\n");

            ADD_TB_PROJECTS_PRODUCTS(               
                NO,
                PROJECTNAMES,
                KINDS,
                OWNER,
                DESIGNER,
                STAGES,
                ISCLOSED,
                DOC_NBR,
                STATUS,
                TASTESREPLYS,
                DESIGNREPLYS,
                QCREPLYS
                );
        }
    }
}
