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
    public partial class FrmRESEARCHDB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        DataSet ds1 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;
        int rowindex;
        int ROWSINDEX;
        int COLUMNSINDEX;

        string ID;
        byte[] BYTES1 = null;
        string CONTENTTYPES1 = null;
        string DOCNAMES1 = null;
        byte[] BYTES2 = null;
        string CONTENTTYPES2 = null;
        string DOCNAMES2 = null;

        byte[] BYTES31 = null;
        string CONTENTTYPES31 = null;
        string DOCNAMES31 = null;
        byte[] BYTES32 = null;
        string CONTENTTYPES32 = null;
        string DOCNAMES32 = null;
        byte[] BYTES33 = null;
        string CONTENTTYPES33 = null;
        string DOCNAMES33 = null;
        byte[] BYTES34 = null;
        string CONTENTTYPES34 = null;
        string DOCNAMES34 = null;
        byte[] BYTES35 = null;
        string CONTENTTYPES35 = null;
        string DOCNAMES35 = null;
        byte[] BYTES36 = null;
        string CONTENTTYPES36 = null;
        string DOCNAMES36 = null;

        byte[] BYTES41= null;
        string CONTENTTYPES41 = null;
        string DOCNAMES41 = null;
        byte[] BYTES42 = null;
        string CONTENTTYPES42 = null;
        string DOCNAMES42 = null;


        public FrmRESEARCHDB()
        {
            InitializeComponent();

            SEARCH(textBox1A.Text.Trim());
            SETdataGridView1();
            SEARCH2(textBox2A.Text.Trim());
            SETdataGridView2();
            SEARCH3(textBox3A.Text.Trim());
            SETdataGridView31();
            SETdataGridView32();
            SETdataGridView33();
            SETdataGridView34();
            SETdataGridView35();
            SETdataGridView36();
            SEARCH4(textBox4A.Text.Trim());
            SETdataGridView41();
            SETdataGridView42();

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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='原料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='原料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARAID";
            comboBox2.DisplayMember = "PARAID";
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='物料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "PARAID";
            comboBox3.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void SEARCH(string KEYS)
        {
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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                        [ID] 
                                        ,[DOCID] AS '文件編號'
                                        ,[COMMENTS]  AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                     
                                        ,[DOCNAMES]
                                        FROM [TKRESEARCH].[dbo].[TBDB1]
                                        WHERE [DOCID] LIKE '%{0}%'
                                        ORDER BY [ID] 

                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                         [ID] 
                                        ,[DOCID] AS '文件編號'
                                        ,[COMMENTS]  AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                      
                                        ,[DOCNAMES]
                                        FROM [TKRESEARCH].[dbo].[TBDB1]
                                        ORDER BY [ID] 
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds1.Tables["ds1"];

                    dataGridView1.AutoResizeColumns();



                }
                else
                {
                    dataGridView1.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void SEARCH2(string KEYS)
        {
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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [ID]
                                        ,[NAMES] AS '產品品名'
                                        ,[CHARS] AS '產品特色'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PRICES] AS '售價'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVEMETHODS] AS '保存方式'
                                        ,[PRIMES] AS '素別'
                                        ,[ALLERGENS] AS '過敏原'
                                        ,[OWNERS] AS '負責廠商'
                                        ,[EATES] AS '建議食用方式'
                                        ,[CHECKSUNITS] AS '驗證單位'
                                        ,[CHECKS] AS '驗證證書字號'
                                        ,[OTHERS] AS '其他口味'
                                        ,[COMMENTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES] AS '產品圖片'
                                        ,[CONTENTTYPES] 
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB2]
                                        WHERE [NAMES] LIKE '%{0}%'
                                        ORDER BY [ID] 

                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                        [ID]
                                        ,[NAMES] AS '產品品名'
                                        ,[CHARS] AS '產品特色'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PRICES] AS '售價'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVEMETHODS] AS '保存方式'
                                        ,[PRIMES] AS '素別'
                                        ,[ALLERGENS] AS '過敏原'
                                        ,[OWNERS] AS '負責廠商'
                                        ,[EATES] AS '建議食用方式'
                                        ,[CHECKSUNITS] AS '驗證單位'
                                        ,[CHECKS] AS '驗證證書字號'
                                        ,[OTHERS] AS '其他口味'
                                        ,[COMMENTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES] AS '產品圖片'
                                        ,[CONTENTTYPES] 
                                       
                                        FROM [TKRESEARCH].[dbo].[TBDB2]                                       
                                        ORDER BY [ID] 
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds1.Tables["ds1"];

                    dataGridView2.AutoResizeColumns();



                }
                else
                {
                    dataGridView2.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void SEARCH3(string KEYS)
        {
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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                       SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線過敏原'
                                        ,[PLACES] AS '產地'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品色澤'
                                        ,[TASTES] AS '產品風味'
                                        ,[LOTS] AS '產品批號'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[BASEONS] AS '基改/非基改'
                                        ,[COA] AS '檢附COA'
                                        ,[INCHECKRATES] AS '抽驗頻率'
                                        ,[RULES] AS '相關法規'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'  
                                        ,[DOCNAMES1] AS '三證'
                                        ,[DOCNAMES2] AS '產品成分'
                                        ,[DOCNAMES3] AS '基改/非基改'
                                        ,[DOCNAMES4] AS '檢驗報告'
                                        ,[DOCNAMES5] AS '營養標示'
                                        ,[DOCNAMES6] AS '產品圖片'                                     
                                        FROM [TKRESEARCH].[dbo].[TBDB3]
                                 
                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線過敏原'
                                        ,[PLACES] AS '產地'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品色澤'
                                        ,[TASTES] AS '產品風味'
                                        ,[LOTS] AS '產品批號'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[BASEONS] AS '基改/非基改'
                                        ,[COA] AS '檢附COA'
                                        ,[INCHECKRATES] AS '抽驗頻率'
                                        ,[RULES] AS '相關法規'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'  
                                        ,[DOCNAMES1] AS '三證'
                                        ,[DOCNAMES2] AS '產品成分'
                                        ,[DOCNAMES3] AS '基改/非基改'
                                        ,[DOCNAMES4] AS '檢驗報告'
                                        ,[DOCNAMES5] AS '營養標示'
                                        ,[DOCNAMES6] AS '產品圖片'                                             
                                        FROM [TKRESEARCH].[dbo].[TBDB3]
                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds1.Tables["ds1"];

                    dataGridView3.AutoResizeColumns();



                }
                else
                {
                    dataGridView3.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        //設定下載欄
        public void SETdataGridView1()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload";
            lnkDownload.HeaderText = "Download";
            lnkDownload.Text = "Download";

            dataGridView1.Columns.Insert(dataGridView1.ColumnCount, lnkDownload);
            dataGridView1.CellContentClick += new DataGridViewCellEventHandler(DataGridView1_CellClick);
        }
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload"))
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                            SELECT
                                             [ID] 
                                            ,[DOCID] AS '文件編號'
                                            ,[COMMENTS]  AS '備註'
                                            ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                            ,[DATAS]  
                                            ,[DOCNAMES]
                                            ,[CONTENTTYPES]
                                            FROM [TKRESEARCH].[dbo].[TBDB1]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS"];
                            contentType = sdr["CONTENTTYPES"].ToString();
                            fileName = sdr["DOCNAMES"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView2()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload31";
            lnkDownload.HeaderText = "Download";
            lnkDownload.Text = "Download";

            dataGridView2.Columns.Insert(dataGridView2.ColumnCount, lnkDownload);
            dataGridView2.CellContentClick += new DataGridViewCellEventHandler(DataGridView2_CellClick);
        }
        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload31"))
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                            SELECT 
                                            [ID]
                                            ,[NAMES]
                                            ,[CHARS]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PRICES]
                                            ,[SAVEDAYS]
                                            ,[SAVEMETHODS]
                                            ,[PRIMES]
                                            ,[ALLERGENS]
                                            ,[OWNERS]
                                            ,[EATES]
                                            ,[CHECKSUNITS]
                                            ,[CHECKS]
                                            ,[OTHERS]
                                            ,[COMMENTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES]
                                            ,[CONTENTTYPES]
                                            ,[DATAS]
                                            FROM [TKRESEARCH].[dbo].[TBDB2]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS"];
                            contentType = sdr["CONTENTTYPES"].ToString();
                            fileName = sdr["DOCNAMES"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView31()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload31";
            lnkDownload.HeaderText = "三證Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView31_CellClick);

        }
        private void DataGridView31_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload31"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS1"];
                            contentType = sdr["CONTENTTYPES1"].ToString();
                            fileName = sdr["DOCNAMES1"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView32()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload32";
            lnkDownload.HeaderText = "產品成分Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView32_CellClick);
            
        }
        private void DataGridView32_CellClick(object sender, DataGridViewCellEventArgs e)
        {   
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload32"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS2"];
                            contentType = sdr["CONTENTTYPES2"].ToString();
                            fileName = sdr["DOCNAMES2"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView33()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload33";
            lnkDownload.HeaderText = "基改/非基改Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView33_CellClick);

        }
        private void DataGridView33_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload33"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS3"];
                            contentType = sdr["CONTENTTYPES3"].ToString();
                            fileName = sdr["DOCNAMES3"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView34()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload34";
            lnkDownload.HeaderText = "檢驗報告Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView34_CellClick);

        }
        private void DataGridView34_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload34"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS4"];
                            contentType = sdr["CONTENTTYPES4"].ToString();
                            fileName = sdr["DOCNAMES4"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView35()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload35";
            lnkDownload.HeaderText = "營養標示Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView35_CellClick);

        }
        private void DataGridView35_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload35"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS5"];
                            contentType = sdr["CONTENTTYPES5"].ToString();
                            fileName = sdr["DOCNAMES5"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView36()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload36";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView36_CellClick);

        }
        private void DataGridView36_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload36"))
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS6"];
                            contentType = sdr["CONTENTTYPES6"].ToString();
                            fileName = sdr["DOCNAMES6"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView41()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload41";
            lnkDownload.HeaderText = "三證Download";
            lnkDownload.Text = "Download";

            dataGridView4.Columns.Insert(dataGridView4.ColumnCount, lnkDownload);
            dataGridView4.CellContentClick += new DataGridViewCellEventHandler(DataGridView41_CellClick);

        }
        private void DataGridView41_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload41"))
            {
                DataGridViewRow row = dataGridView4.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                           SELECT 
                                             [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[SPECS]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[COA]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB4]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS1"];
                            contentType = sdr["CONTENTTYPES1"].ToString();
                            fileName = sdr["DOCNAMES1"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }

        public void SETdataGridView42()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload42";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView4.Columns.Insert(dataGridView4.ColumnCount, lnkDownload);
            dataGridView4.CellContentClick += new DataGridViewCellEventHandler(DataGridView42_CellClick);

        }
        private void DataGridView42_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            if (e.RowIndex >= 0 && columnName.Equals("lnkDownload42"))
            {
                DataGridViewRow row = dataGridView4.Rows[e.RowIndex];

                int ID = Convert.ToInt16((row.Cells["ID"].Value));
                byte[] bytes;
                string fileName, contentType;

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                using (SqlConnection con = sqlConn)
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        SQL.AppendFormat(@"
                                          SELECT 
                                             [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[SPECS]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[COA]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB4]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Connection = con;
                        con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["DATAS2"];
                            contentType = sdr["CONTENTTYPES2"].ToString();
                            fileName = sdr["DOCNAMES2"].ToString();

                            Stream stream;
                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "All files (*.*)|*.*";
                            saveFileDialog.FilterIndex = 1;
                            saveFileDialog.RestoreDirectory = true;
                            saveFileDialog.FileName = fileName;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                stream = saveFileDialog.OpenFile();
                                stream.Write(bytes, 0, bytes.Length);
                                stream.Close();
                            }
                        }
                    }
                    con.Close();
                }
            }
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1B.Text = null;
            textBox1C.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (dataGridView1.CurrentCell.RowIndex > 0 || dataGridView1.CurrentCell.ColumnIndex > 0)
                {
                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;                }


                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1B.Text = row.Cells["ID"].Value.ToString();
                    textBox1C.Text = row.Cells["ID"].Value.ToString();
                    textBox14.Text = row.Cells["文件編號"].Value.ToString();
                    textBox15.Text = row.Cells["備註"].Value.ToString();


                }
                else
                {
                    textBox1B.Text = null;
                    textBox1C.Text = null;
                    textBox14.Text = null;
                    textBox15.Text = null;

                }
            }

        }
       

    

        public void OPEN1()
        {
            string FILETYPE = null;
            CONTENTTYPES1 = "";
            BYTES1 = null;
            DOCNAMES1 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;
                    DOCNAMES1 = Path.GetFileName(fileName);
                    textBox13.Text = fileName;

                    BYTES1 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES1 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES1 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES1 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES1 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES1 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES1 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES1 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES1 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES1 = "image/bmp";
                            break;
                    }

                
                }
            }
        }

        public void OPEN2()
        {
            string FILETYPE = null;
            CONTENTTYPES2 = "";
            BYTES2 = null;
            DOCNAMES2 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES2 = Path.GetFileName(fileName);
                    textBox216.Text = fileName;

                    BYTES2 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES2 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES2 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES2 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES2 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES2 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES2 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES2 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES2 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES2 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN31()
        {
            string FILETYPE = null;
            CONTENTTYPES31 = "";
            BYTES31 = null;
            DOCNAMES31 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES31 = Path.GetFileName(fileName);
                    textBox320.Text = fileName;

                    BYTES31 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES31 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES31 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES31 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES31 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES31 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES31 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES31 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES31 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES31 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN32()
        {
            string FILETYPE = null;
            CONTENTTYPES32 = "";
            BYTES32 = null;
            DOCNAMES32 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES32 = Path.GetFileName(fileName);
                    textBox321.Text = fileName;

                    BYTES32 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES32 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES32 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES32 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES32 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES32 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES32 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES32 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES32 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES32 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN33()
        {
            string FILETYPE = null;
            CONTENTTYPES33 = "";
            BYTES33 = null;
            DOCNAMES33 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES33 = Path.GetFileName(fileName);
                    textBox322.Text = fileName;

                    BYTES33 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES33 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES33 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES33 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES33 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES33 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES33 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES33 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES33 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES33 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN34()
        {
            string FILETYPE = null;
            CONTENTTYPES34 = "";
            BYTES34 = null;
            DOCNAMES34 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES34 = Path.GetFileName(fileName);
                    textBox323.Text = fileName;

                    BYTES34 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES34 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES34 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES34 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES34 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES34 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES34 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES34 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES34 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES34 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN35()
        {
            string FILETYPE = null;
            CONTENTTYPES35 = "";
            BYTES35 = null;
            DOCNAMES35 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES35 = Path.GetFileName(fileName);
                    textBox324.Text = fileName;

                    BYTES35 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES35 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES35 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES35 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES35 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES35 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES35 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES35 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES35 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES35 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN36()
        {
            string FILETYPE = null;
            CONTENTTYPES36 = "";
            BYTES36 = null;
            DOCNAMES36 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES36 = Path.GetFileName(fileName);
                    textBox325.Text = fileName;

                    BYTES36 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES36 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES36 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES36 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES36 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES36 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES36 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES36 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES36 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES36 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN41()
        {
            string FILETYPE = null;
            CONTENTTYPES41 = "";
            BYTES41 = null;
            DOCNAMES41 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES41 = Path.GetFileName(fileName);
                    textBox410.Text = fileName;

                    BYTES41 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES41 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES41 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES41 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES41 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES41 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES41 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES41 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES41 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES41 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void OPEN42()
        {
            string FILETYPE = null;
            CONTENTTYPES42 = "";
            BYTES42 = null;
            DOCNAMES42 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES42 = Path.GetFileName(fileName);
                    textBox411.Text = fileName;

                    BYTES42 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES42 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES42 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES42 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES42 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES42 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES42 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES42 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES42 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES42 = "image/bmp";
                            break;
                    }


                }
            }
        }

        public void ADD_TO_TBDB1(string DOCID,string COMMENTS, string DOCNAMES, string CONTENTTYPES,byte[] BYTES)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if(!string.IsNullOrEmpty(DOCID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB1]
                                        (
                                        [DOCID]
                                        ,[COMMENTS] 
                                        ,[DOCNAMES]
                                        ,[CONTENTTYPES]
                                        ,[DATAS]
                                        )
                                        VALUES
                                        (
                                        @DOCID
                                        ,@COMMENTS                                   
                                        ,@DOCNAMES
                                        ,@CONTENTTYPES
                                        ,@DATAS
                                        )
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@DOCID", DOCID);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                        cmd.Parameters.AddWithValue("@DOCNAMES", DOCNAMES);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES", CONTENTTYPES);
                        cmd.Parameters.AddWithValue("@DATAS", BYTES);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }
                
            }

          
        }
        public void  UPDATE_TO_TBDB1(string ID,string DOCID,string COMMENTS)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE[TKRESEARCH].[dbo].[TBDB1]
                                        SET DOCID=@DOCID,COMMENTS=@COMMENTS
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@DOCID", DOCID);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }

          
        }

        public void DELETE_TO_TBDB1(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB1]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                    
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }

          
        }

        public void ADD_TO_TBDB2(
                                string NAMES
                                , string CHARS
                                , string ORIS
                                , string SPECS
                                , string PRICES
                                , string SAVEDAYS
                                , string SAVEMETHODS
                                , string PRIMES
                                , string ALLERGENS
                                , string OWNERS
                                , string EATES
                                , string CHECKSUNITS
                                , string CHECKS
                                , string OTHERS
                                , string COMMENTS
                                , string DOCNAMES
                                , string CONTENTTYPES,
                                byte[] BYTES
                                )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                       
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB2]
                                        (
                                        [NAMES]
                                        ,[CHARS]
                                        ,[ORIS]
                                        ,[SPECS]
                                        ,[PRICES]
                                        ,[SAVEDAYS]
                                        ,[SAVEMETHODS]
                                        ,[PRIMES]
                                        ,[ALLERGENS]
                                        ,[OWNERS]
                                        ,[EATES]
                                        ,[CHECKSUNITS]
                                        ,[CHECKS]
                                        ,[OTHERS]
                                        ,[COMMENTS]
                                        ,[DOCNAMES]
                                        ,[CONTENTTYPES]
                                        ,[DATAS]
                                        )
                                        VALUES
                                        (
                                        @NAMES
                                        ,@CHARS
                                        ,@ORIS
                                        ,@SPECS
                                        ,@PRICES
                                        ,@SAVEDAYS
                                        ,@SAVEMETHODS
                                        ,@PRIMES
                                        ,@ALLERGENS
                                        ,@OWNERS
                                        ,@EATES
                                        ,@CHECKSUNITS
                                        ,@CHECKS
                                        ,@OTHERS
                                        ,@COMMENTS
                                        ,@DOCNAMES
                                        ,@CONTENTTYPES
                                        ,@DATAS
                                        )
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@CHARS", CHARS);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PRICES", PRICES);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVEMETHODS", SAVEMETHODS);
                        cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                        cmd.Parameters.AddWithValue("@ALLERGENS", ALLERGENS);
                        cmd.Parameters.AddWithValue("@OWNERS", OWNERS);
                        cmd.Parameters.AddWithValue("@EATES", EATES);
                        cmd.Parameters.AddWithValue("@CHECKSUNITS", CHECKSUNITS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@OTHERS", OTHERS);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                        cmd.Parameters.AddWithValue("@DOCNAMES", DOCNAMES);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES", CONTENTTYPES);
                        cmd.Parameters.AddWithValue("@DATAS", BYTES);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }

           
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {

            textBox2B.Text = null;
            textBox2C.Text = null;
            textBox221.Text = null;
            textBox222.Text = null;
            textBox223.Text = null;
            textBox224.Text = null;
            textBox225.Text = null;
            textBox226.Text = null;
            textBox227.Text = null;
            textBox228.Text = null;
            textBox229.Text = null;
            textBox230.Text = null;
            textBox231.Text = null;
            textBox232.Text = null;
            textBox233.Text = null;
            textBox234.Text = null;
            textBox235.Text = null;
            textBox241.Text = null;


            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;               

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox2B.Text = row.Cells["ID"].Value.ToString();
                    textBox2C.Text = row.Cells["ID"].Value.ToString();
                    textBox221.Text = row.Cells["產品品名"].Value.ToString();
                    textBox222.Text = row.Cells["產品特色"].Value.ToString();
                    textBox223.Text = row.Cells["產品成分"].Value.ToString();
                    textBox224.Text = row.Cells["產品規格"].Value.ToString();
                    textBox225.Text = row.Cells["售價"].Value.ToString();
                    textBox226.Text = row.Cells["保存期限"].Value.ToString();
                    textBox227.Text = row.Cells["保存方式"].Value.ToString();
                    textBox228.Text = row.Cells["素別"].Value.ToString();
                    textBox229.Text = row.Cells["過敏原"].Value.ToString();
                    textBox230.Text = row.Cells["負責廠商"].Value.ToString();
                    textBox231.Text = row.Cells["建議食用方式"].Value.ToString();
                    textBox232.Text = row.Cells["驗證單位"].Value.ToString();
                    textBox233.Text = row.Cells["驗證證書字號"].Value.ToString();
                    textBox234.Text = row.Cells["其他口味"].Value.ToString();
                    textBox235.Text = row.Cells["備註"].Value.ToString();
                    textBox241.Text = row.Cells["產品品名"].Value.ToString();

                }
                else
                {
                    textBox2B.Text = null;
                    textBox2C.Text = null;
                    textBox221.Text = null;
                    textBox222.Text = null;
                    textBox223.Text = null;
                    textBox224.Text = null;
                    textBox225.Text = null;
                    textBox226.Text = null;
                    textBox227.Text = null;
                    textBox228.Text = null;
                    textBox229.Text = null;
                    textBox230.Text = null;
                    textBox231.Text = null;
                    textBox232.Text = null;
                    textBox233.Text = null;
                    textBox234.Text = null;
                    textBox235.Text = null;
                    textBox241.Text = null;

                }
            }
        }

        public void UPDATE_TO_TBDB2(string ID
                                , string NAMES
                                , string CHARS
                                , string ORIS
                                , string SPECS
                                , string PRICES
                                , string SAVEDAYS
                                , string SAVEMETHODS
                                , string PRIMES
                                , string ALLERGENS
                                , string OWNERS
                                , string EATES
                                , string CHECKSUNITS
                                , string CHECKS
                                , string OTHERS
                                , string COMMENTS
                                )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB2]
                                        SET [NAMES]=@NAMES
                                        ,[CHARS]=@CHARS
                                        ,[ORIS]=@ORIS
                                        ,[SPECS]=@SPECS
                                        ,[PRICES]=@PRICES
                                        ,[SAVEDAYS]=@SAVEDAYS
                                        ,[SAVEMETHODS]=@SAVEMETHODS
                                        ,[PRIMES]=@PRIMES
                                        ,[ALLERGENS]=@ALLERGENS
                                        ,[OWNERS]=@OWNERS
                                        ,[EATES]=@EATES
                                        ,[CHECKSUNITS]=@CHECKSUNITS
                                        ,[CHECKS]=@CHECKS
                                        ,[OTHERS]=@OTHERS
                                        ,[COMMENTS]=@COMMENTS
                                 
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@CHARS", CHARS);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PRICES", PRICES);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVEMETHODS", SAVEMETHODS);
                        cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                        cmd.Parameters.AddWithValue("@ALLERGENS", ALLERGENS);
                        cmd.Parameters.AddWithValue("@OWNERS", OWNERS);
                        cmd.Parameters.AddWithValue("@EATES", EATES);
                        cmd.Parameters.AddWithValue("@CHECKSUNITS", CHECKSUNITS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@OTHERS", OTHERS);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

      


        public void ADD_TO_TBDB3(
                                string KINDS
                                , string SUPPLYS
                                , string NAMES
                                , string ORIS
                                , string SPECS
                                , string PROALLGENS
                                , string MANUALLGENS
                                , string PLACES
                                , string OUTS
                                , string COLORS
                                , string TASTES
                                , string LOTS
                                , string CHECKS
                                , string SAVEDAYS
                                , string SAVECONDITIONS
                                , string BASEONS
                                , string COA
                                , string INCHECKRATES
                                , string RULES
                                , string COMMEMTS
                                , string DOCNAMES1
                                , string CONTENTTYPES1
                                , byte[] DATAS1
                                , string DOCNAMES2
                                , string CONTENTTYPES2
                                , byte[] DATAS2
                                , string DOCNAMES3
                                , string CONTENTTYPES3
                                , byte[] DATAS3
                                , string DOCNAMES4
                                , string CONTENTTYPES4
                                , byte[] DATAS4
                                , string DOCNAMES5
                                , string CONTENTTYPES5
                                , byte[] DATAS5
                                , string DOCNAMES6
                                , string CONTENTTYPES6
                                , byte[] DATAS6
                                )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                       
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB3]
                                        (
                                         [KINDS]
                                        ,[SUPPLYS]
                                        ,[NAMES]
                                        ,[ORIS]
                                        ,[SPECS]
                                        ,[PROALLGENS]
                                        ,[MANUALLGENS]
                                        ,[PLACES]
                                        ,[OUTS]
                                        ,[COLORS]
                                        ,[TASTES]
                                        ,[LOTS]
                                        ,[CHECKS]
                                        ,[SAVEDAYS]
                                        ,[SAVECONDITIONS]
                                        ,[BASEONS]
                                        ,[COA]
                                        ,[INCHECKRATES]
                                        ,[RULES]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        ,[DOCNAMES4]
                                        ,[CONTENTTYPES4]
                                        ,[DATAS4]
                                        ,[DOCNAMES5]
                                        ,[CONTENTTYPES5]
                                        ,[DATAS5]
                                        ,[DOCNAMES6]
                                        ,[CONTENTTYPES6]
                                        ,[DATAS6]
                                       
                                        )
                                        VALUES
                                        (
                                          @KINDS
                                        ,@SUPPLYS
                                        ,@NAMES
                                        ,@ORIS
                                        ,@SPECS
                                        ,@PROALLGENS
                                        ,@MANUALLGENS
                                        ,@PLACES
                                        ,@OUTS
                                        ,@COLORS
                                        ,@TASTES
                                        ,@LOTS
                                        ,@CHECKS
                                        ,@SAVEDAYS
                                        ,@SAVECONDITIONS
                                        ,@BASEONS
                                        ,@COA
                                        ,@INCHECKRATES
                                        ,@RULES
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        ,@DOCNAMES3
                                        ,@CONTENTTYPES3
                                        ,@DATAS3
                                        ,@DOCNAMES4
                                        ,@CONTENTTYPES4
                                        ,@DATAS4
                                        ,@DOCNAMES5
                                        ,@CONTENTTYPES5
                                        ,@DATAS5
                                        ,@DOCNAMES6
                                        ,@CONTENTTYPES6
                                        ,@DATAS6
                                        )
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                        cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                        cmd.Parameters.AddWithValue("@PLACES", PLACES);
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);
                        cmd.Parameters.AddWithValue("@TASTES", TASTES);
                        cmd.Parameters.AddWithValue("@LOTS", LOTS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                        cmd.Parameters.AddWithValue("@BASEONS", BASEONS);
                        cmd.Parameters.AddWithValue("@COA", COA);
                        cmd.Parameters.AddWithValue("@INCHECKRATES", INCHECKRATES);
                        cmd.Parameters.AddWithValue("@RULES", RULES);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                        cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                        cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                        cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                        cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                        cmd.Parameters.AddWithValue("@DOCNAMES3", DOCNAMES3);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES3", CONTENTTYPES3);
                        cmd.Parameters.AddWithValue("@DATAS3", DATAS3);
                        cmd.Parameters.AddWithValue("@DOCNAMES4", DOCNAMES4);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES4", CONTENTTYPES4);
                        cmd.Parameters.AddWithValue("@DATAS4", DATAS4);
                        cmd.Parameters.AddWithValue("@DOCNAMES5", DOCNAMES5);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES5", CONTENTTYPES5);
                        cmd.Parameters.AddWithValue("@DATAS5", DATAS5);
                        cmd.Parameters.AddWithValue("@DOCNAMES6", DOCNAMES6);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES6", CONTENTTYPES6);
                        cmd.Parameters.AddWithValue("@DATAS6", DATAS6);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox331.Text = null;
            textBox332.Text = null;
            textBox333.Text = null;
            textBox334.Text = null;
            textBox335.Text = null;
            textBox336.Text = null;
            textBox337.Text = null;
            textBox338.Text = null;
            textBox339.Text = null;
            textBox340.Text = null;
            textBox341.Text = null;
            textBox342.Text = null;
            textBox343.Text = null;
            textBox344.Text = null;
            textBox345.Text = null;
            textBox346.Text = null;
            textBox347.Text = null;
            textBox348.Text = null;
            textBox349.Text = null;
            textBox350.Text = null;
            textBox3B.Text = null;
            textBox3C.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
              

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox3B.Text = row.Cells["ID"].Value.ToString();
                    textBox3C.Text = row.Cells["ID"].Value.ToString();

                    comboBox2.Text = row.Cells["分類"].Value.ToString();
                    textBox331.Text = row.Cells["供應商"].Value.ToString();
                    textBox332.Text = row.Cells["產品品名"].Value.ToString();
                    textBox333.Text = row.Cells["產品成分"].Value.ToString();
                    textBox334.Text = row.Cells["產品規格"].Value.ToString();
                    textBox335.Text = row.Cells["產品過敏原"].Value.ToString();
                    textBox336.Text = row.Cells["產線過敏原"].Value.ToString();
                    textBox337.Text = row.Cells["產地"].Value.ToString();
                    textBox338.Text = row.Cells["產品外觀"].Value.ToString();
                    textBox339.Text = row.Cells["產品色澤"].Value.ToString();
                    textBox340.Text = row.Cells["產品風味"].Value.ToString();
                    textBox341.Text = row.Cells["產品批號"].Value.ToString();
                    textBox342.Text = row.Cells["外包裝及驗收標準"].Value.ToString();
                    textBox343.Text = row.Cells["保存期限"].Value.ToString();
                    textBox344.Text = row.Cells["保存條件"].Value.ToString();
                    textBox345.Text = row.Cells["基改/非基改"].Value.ToString();
                    textBox346.Text = row.Cells["檢附COA"].Value.ToString();
                    textBox347.Text = row.Cells["抽驗頻率"].Value.ToString();
                    textBox348.Text = row.Cells["相關法規"].Value.ToString();
                    textBox349.Text = row.Cells["備註"].Value.ToString();

                    textBox350.Text = row.Cells["產品品名"].Value.ToString();
                    


                }
                else
                {
                    textBox331.Text = null;
                    textBox332.Text = null;
                    textBox333.Text = null;
                    textBox334.Text = null;
                    textBox335.Text = null;
                    textBox336.Text = null;
                    textBox337.Text = null;
                    textBox338.Text = null;
                    textBox339.Text = null;
                    textBox340.Text = null;
                    textBox341.Text = null;
                    textBox342.Text = null;
                    textBox343.Text = null;
                    textBox344.Text = null;
                    textBox345.Text = null;
                    textBox346.Text = null;
                    textBox347.Text = null;
                    textBox348.Text = null;
                    textBox349.Text = null;
                    textBox350.Text = null;
                    textBox3B.Text = null;
                    textBox3C.Text = null;
                }
            }

        }

        public void UPDATE_TO_TBDB3(
                                string ID
                               , string KINDS
                               , string SUPPLYS
                               , string NAMES
                               , string ORIS
                               , string SPECS
                               , string PROALLGENS
                               , string MANUALLGENS
                               , string PLACES
                               , string OUTS
                               , string COLORS
                               , string TASTES
                               , string LOTS
                               , string CHECKS
                               , string SAVEDAYS
                               , string SAVECONDITIONS
                               , string BASEONS
                               , string COA
                               , string INCHECKRATES
                               , string RULES
                               , string COMMEMTS
                              
                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB3]
                                        SET
                                        [KINDS]=@KINDS
                                        ,[SUPPLYS]=@SUPPLYS
                                        ,[NAMES]=@NAMES
                                        ,[ORIS]=@ORIS
                                        ,[SPECS]=@SPECS
                                        ,[PROALLGENS]=@PROALLGENS
                                        ,[MANUALLGENS]=@MANUALLGENS
                                        ,[PLACES]=@PLACES
                                        ,[OUTS]=@OUTS
                                        ,[COLORS]=@COLORS
                                        ,[TASTES]=@TASTES
                                        ,[LOTS]=@LOTS
                                        ,[CHECKS]=@CHECKS
                                        ,[SAVEDAYS]=@SAVEDAYS
                                        ,[SAVECONDITIONS]=@SAVECONDITIONS
                                        ,[BASEONS]=@BASEONS
                                        ,[COA]=@COA
                                        ,[INCHECKRATES]=@INCHECKRATES
                                        ,[RULES]=@RULES
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                        cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                        cmd.Parameters.AddWithValue("@PLACES", PLACES);
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);
                        cmd.Parameters.AddWithValue("@TASTES", TASTES);
                        cmd.Parameters.AddWithValue("@LOTS", LOTS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                        cmd.Parameters.AddWithValue("@BASEONS", BASEONS);
                        cmd.Parameters.AddWithValue("@COA", COA);
                        cmd.Parameters.AddWithValue("@INCHECKRATES", INCHECKRATES);
                        cmd.Parameters.AddWithValue("@RULES", RULES);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                       

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB2(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB2]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void DELETE_TO_TBDB3(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB3]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH4(string KEYS)
        {
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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                      SELECT  
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[SPECS] AS '產品規格'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品顏色'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[COA] AS '檢附COA'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                        
                                        ,[DOCNAMES2] AS '產品圖片'
                                        
                                        FROM [TKRESEARCH].[dbo].[TBDB4]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT  
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[SPECS] AS '產品規格'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品顏色'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[COA] AS '檢附COA'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                        
                                        ,[DOCNAMES2] AS '產品圖片'
                                        
                                        FROM [TKRESEARCH].[dbo].[TBDB4]
                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView4.DataSource = ds1.Tables["ds1"];

                    dataGridView4.AutoResizeColumns();



                }
                else
                {
                    dataGridView3.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB4(
                               string KINDS
                               , string SUPPLYS
                               , string NAMES                               
                               , string SPECS                             
                               , string OUTS
                               , string COLORS                            
                               , string CHECKS
                               , string SAVEDAYS                              
                               , string COA                              
                               , string COMMEMTS
                               , string DOCNAMES1
                               , string CONTENTTYPES1
                               , byte[] DATAS1
                               , string DOCNAMES2
                               , string CONTENTTYPES2
                               , byte[] DATAS2
                          
                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"                                      
                                        
                                        INSERT INTO  [TKRESEARCH].[dbo].[TBDB4]
                                        (
                                        [KINDS]
                                        ,[SUPPLYS]
                                        ,[NAMES]
                                        ,[SPECS]
                                        ,[OUTS]
                                        ,[COLORS]
                                        ,[CHECKS]
                                        ,[SAVEDAYS]
                                        ,[COA]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        )
                                        VALUES
                                        (
                                        @KINDS
                                        ,@SUPPLYS
                                        ,@NAMES
                                        ,@SPECS
                                        ,@OUTS
                                        ,@COLORS
                                        ,@CHECKS
                                        ,@SAVEDAYS
                                        ,@COA
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        )
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);                        
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);  
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@COA", COA);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                        cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                        cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                        cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                        cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                        cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                      
                        

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void UPDATE_TO_TBDB4(
                                string ID
                               , string KINDS
                               , string SUPPLYS
                               , string NAMES                            
                               , string SPECS                             
                               , string OUTS
                               , string COLORS                            
                               , string CHECKS
                               , string SAVEDAYS                              
                               , string COA                            
                               , string COMMEMTS

                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB4]
                                        SET
                                        [KINDS]=@KINDS
                                        ,[SUPPLYS]=@SUPPLYS
                                        ,[NAMES]=@NAMES                                       
                                        ,[SPECS]=@SPECS                                       
                                        ,[OUTS]=@OUTS
                                        ,[COLORS]=@COLORS                                        
                                        ,[CHECKS]=@CHECKS
                                        ,[SAVEDAYS]=@SAVEDAYS                                       
                                        ,[COA]=@COA                                      
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);                       
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);                       
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);                        
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);                        
                        cmd.Parameters.AddWithValue("@COA", COA);                        
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);


                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBox4B.Text = null;
            textBox4C.Text = null;
            textBox420.Text = null;
            textBox421.Text = null;
            textBox422.Text = null;
            textBox423.Text = null;
            textBox424.Text = null;
            textBox425.Text = null;
            textBox426.Text = null;
            textBox427.Text = null;
            textBox428.Text = null;
            textBox430.Text = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox4B.Text = row.Cells["ID"].Value.ToString();
                    textBox4C.Text = row.Cells["ID"].Value.ToString();

                    comboBox4.Text= row.Cells["分類"].Value.ToString();
                    textBox420.Text = row.Cells["供應商"].Value.ToString();
                    textBox421.Text = row.Cells["產品品名"].Value.ToString();
                    textBox422.Text = row.Cells["產品規格"].Value.ToString();
                    textBox423.Text = row.Cells["產品外觀"].Value.ToString();
                    textBox424.Text = row.Cells["產品顏色"].Value.ToString();
                    textBox425.Text = row.Cells["外包裝及驗收標準"].Value.ToString();
                    textBox426.Text = row.Cells["保存期限"].Value.ToString();
                    textBox427.Text = row.Cells["檢附COA"].Value.ToString();
                    textBox428.Text = row.Cells["備註"].Value.ToString();
                    textBox430.Text = row.Cells["產品品名"].Value.ToString();


                }
                else
                {
                    textBox4B.Text = null;
                    textBox4C.Text = null;
                    textBox420.Text = null;
                    textBox421.Text = null;
                    textBox422.Text = null;
                    textBox423.Text = null;
                    textBox424.Text = null;
                    textBox425.Text = null;
                    textBox426.Text = null;
                    textBox427.Text = null;
                    textBox428.Text = null;
                    textBox430.Text = null;

                }
            }

        }

        public void DELETE_TO_TBDB4(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB4]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH5(string KEYS)
        {
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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                      SELECT 
                                        [ID]
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                       
                                        ,[DOCNAMES2] AS '產品成分'
                                        
                                        ,[DOCNAMES3] AS '產品責任險'
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB5]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                         SELECT 
                                        [ID]
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                       
                                        ,[DOCNAMES2] AS '產品成分'
                                        
                                        ,[DOCNAMES3] AS '產品責任險'
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB5]

                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView5.DataSource = ds1.Tables["ds1"];

                    dataGridView5.AutoResizeColumns();



                }
                else
                {
                    dataGridView5.DataSource = null;

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
            SEARCH(textBox1A.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OPEN1();

        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADD_TO_TBDB1(textBox11.Text, textBox12.Text, DOCNAMES1, CONTENTTYPES1, BYTES1);
            SEARCH(textBox1A.Text.Trim());
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATE_TO_TBDB1(textBox1B.Text,textBox14.Text,textBox15.Text);
            SEARCH(textBox1A.Text.Trim());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TO_TBDB1(textBox1C.Text);
                SEARCH(textBox1A.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH2(textBox2A.Text.Trim());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OPEN2();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string NAMES = textBox201.Text;
            string CHARS = textBox202.Text;
            string ORIS = textBox203.Text;
            string SPECS = textBox204.Text;
            string PRICES = textBox205.Text;
            string SAVEDAYS = textBox206.Text;
            string SAVEMETHODS = textBox207.Text;
            string PRIMES = textBox208.Text; 
            string ALLERGENS = textBox209.Text;
            string OWNERS = textBox210.Text;
            string EATES = textBox211.Text;
            string CHECKSUNITS = textBox212.Text;
            string CHECKS = textBox213.Text;
            string OTHERS = textBox214.Text;
            string COMMENTS = textBox215.Text;


            ADD_TO_TBDB2( NAMES
                                ,  CHARS
                                ,  ORIS
                                ,  SPECS
                                ,  PRICES
                                ,  SAVEDAYS
                                ,  SAVEMETHODS
                                ,  PRIMES
                                ,  ALLERGENS
                                ,  OWNERS
                                ,  EATES
                                ,  CHECKSUNITS
                                ,  CHECKS
                                ,  OTHERS
                                ,  COMMENTS
                                ,  DOCNAMES2
                                ,  CONTENTTYPES2
                                ,  BYTES2
                                );
            SEARCH2(textBox2A.Text.Trim());
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string ID = textBox2B.Text;
            string NAMES = textBox221.Text;
            string CHARS = textBox222.Text;
            string ORIS = textBox223.Text;
            string SPECS = textBox224.Text;
            string PRICES = textBox225.Text;
            string SAVEDAYS = textBox226.Text;
            string SAVEMETHODS = textBox227.Text;
            string PRIMES = textBox228.Text;
            string ALLERGENS = textBox229.Text;
            string OWNERS = textBox230.Text;
            string EATES = textBox231.Text;
            string CHECKSUNITS = textBox232.Text;
            string CHECKS = textBox233.Text;
            string OTHERS = textBox234.Text;
            string COMMENTS = textBox235.Text;

            UPDATE_TO_TBDB2(ID, NAMES, CHARS, ORIS, SPECS, PRICES, SAVEDAYS, SAVEMETHODS, PRIMES, ALLERGENS, OWNERS, EATES, CHECKSUNITS, CHECKS, OTHERS, COMMENTS);
            SEARCH2(textBox2A.Text.Trim());
        }
        private void button10_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TO_TBDB2(textBox2C.Text);
                SEARCH2(textBox2A.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SEARCH3(textBox3A.Text);
        }
        private void button13_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox1.Text.ToString();
            string SUPPLYS = textBox301.Text.ToString();
            string NAMES = textBox302.Text.ToString();
            string ORIS = textBox303.Text.ToString();
            string SPECS = textBox304.Text.ToString();
            string PROALLGENS = textBox305.Text.ToString();
            string MANUALLGENS = textBox306.Text.ToString();
            string PLACES = textBox307.Text.ToString();
            string OUTS = textBox308.Text.ToString();
            string COLORS = textBox309.Text.ToString();
            string TASTES = textBox310.Text.ToString();
            string LOTS = textBox311.Text.ToString();
            string CHECKS = textBox312.Text.ToString();
            string SAVEDAYS = textBox313.Text.ToString();
            string SAVECONDITIONS = textBox314.Text.ToString();
            string BASEONS = textBox315.Text.ToString();
            string COA = textBox315.Text.ToString();
            string INCHECKRATES = textBox317.Text.ToString();
            string RULES = textBox318.Text.ToString();
            string COMMEMTS = textBox319.Text.ToString();
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1= new byte[] {1};
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
            string DOCNAMES3 = "";
            string CONTENTTYPES3 = "";
            byte[] DATAS3 = new byte[] { 1 };
            string DOCNAMES4 = "";
            string CONTENTTYPES4 = "";
            byte[] DATAS4 = new byte[] { 1 };
            string DOCNAMES5 = "";
            string CONTENTTYPES5 = "";
            byte[] DATAS5 = new byte[] { 1 };
            string DOCNAMES6 = "";
            string CONTENTTYPES6 = "";
            byte[] DATAS6 = new byte[] { 1 };

            if (!string.IsNullOrEmpty(DOCNAMES31))
            {
                DOCNAMES1 = DOCNAMES31;
                CONTENTTYPES1 = CONTENTTYPES31;
                DATAS1 = BYTES31;
            }
            if (!string.IsNullOrEmpty(DOCNAMES32))
            {
                DOCNAMES2 = DOCNAMES32;
                CONTENTTYPES2 = CONTENTTYPES32;
                DATAS2 = BYTES32;
            }
            if (!string.IsNullOrEmpty(DOCNAMES33))
            {
                DOCNAMES3 = DOCNAMES33;
                CONTENTTYPES3 = CONTENTTYPES33;
                DATAS3 = BYTES33;
            }
            if (!string.IsNullOrEmpty(DOCNAMES34))
            {
                DOCNAMES4 = DOCNAMES34;
                CONTENTTYPES4 = CONTENTTYPES34;
                DATAS4 = BYTES34;
            }
            if (!string.IsNullOrEmpty(DOCNAMES35))
            {
                DOCNAMES5 = DOCNAMES35;
                CONTENTTYPES5 = CONTENTTYPES35;
                DATAS5 = BYTES35;
            }
            if (!string.IsNullOrEmpty(DOCNAMES36))
            {
                DOCNAMES6 = DOCNAMES36;
                CONTENTTYPES6 = CONTENTTYPES36;
                DATAS6 = BYTES36;
            }


      

            ADD_TO_TBDB3(
                        KINDS
                        , SUPPLYS
                        , NAMES
                        , ORIS
                        , SPECS
                        , PROALLGENS
                        , MANUALLGENS
                        , PLACES
                        , OUTS
                        , COLORS
                        , TASTES
                        , LOTS
                        , CHECKS
                        , SAVEDAYS
                        , SAVECONDITIONS
                        , BASEONS
                        , COA
                        , INCHECKRATES
                        , RULES
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        , DOCNAMES3
                        , CONTENTTYPES3
                        , DATAS3
                        , DOCNAMES4
                        , CONTENTTYPES4
                        , DATAS4
                        , DOCNAMES5
                        , CONTENTTYPES5
                        , DATAS5
                        , DOCNAMES6
                        , CONTENTTYPES6
                        , DATAS6
                        );

            SEARCH3(textBox3A.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OPEN31();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            OPEN32();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            OPEN33();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            OPEN34();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            OPEN35();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            OPEN36();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string ID = textBox3B.Text;
            string KINDS = comboBox2.Text;

            string SUPPLYS = textBox331.Text;
            string NAMES = textBox332.Text;
            string ORIS = textBox333.Text;
            string SPECS = textBox334.Text;
            string PROALLGENS = textBox335.Text;
            string MANUALLGENS = textBox336.Text;
            string PLACES = textBox337.Text;
            string OUTS = textBox338.Text;
            string COLORS = textBox339.Text;
            string TASTES = textBox340.Text;
            string LOTS = textBox341.Text;
            string CHECKS = textBox342.Text;
            string SAVEDAYS = textBox343.Text;
            string SAVECONDITIONS = textBox344.Text;
            string BASEONS = textBox345.Text;
            string COA = textBox346.Text;
            string INCHECKRATES = textBox347.Text;
            string RULES = textBox348.Text;
            string COMMEMTS = textBox349.Text;

            UPDATE_TO_TBDB3(
                                ID
                              , KINDS
                              , SUPPLYS
                              , NAMES
                              , ORIS
                              , SPECS
                              , PROALLGENS
                              , MANUALLGENS
                              , PLACES
                              , OUTS
                              , COLORS
                              , TASTES
                              , LOTS
                              , CHECKS
                              , SAVEDAYS
                              , SAVECONDITIONS
                              , BASEONS
                              , COA
                              , INCHECKRATES
                              , RULES
                              , COMMEMTS

                              );

            SEARCH3(textBox3A.Text.Trim());
        }
        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TO_TBDB3(textBox3C.Text);
                SEARCH3(textBox3A.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button21_Click(object sender, EventArgs e)
        {
            SEARCH4(textBox4A.Text.Trim());
        }

        private void button22_Click(object sender, EventArgs e)
        {
            OPEN41();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            OPEN42();
        }
        private void button28_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox3.Text.ToString();
            string SUPPLYS = textBox401.Text.ToString();
            string NAMES = textBox402.Text.ToString();
            string SPECS = textBox403.Text.ToString();           
            string OUTS = textBox404.Text.ToString();
            string COLORS = textBox405.Text.ToString();           
            string CHECKS = textBox406.Text.ToString();
            string SAVEDAYS = textBox407.Text.ToString();          
            string COMMEMTS = textBox408.Text.ToString();
            string COA = textBox409.Text.ToString();
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
           

            if (!string.IsNullOrEmpty(DOCNAMES41))
            {
                DOCNAMES1 = DOCNAMES41;
                CONTENTTYPES1 = CONTENTTYPES41;
                DATAS1 = BYTES41;
            }
            if (!string.IsNullOrEmpty(DOCNAMES42))
            {
                DOCNAMES2 = DOCNAMES42;
                CONTENTTYPES2 = CONTENTTYPES42;
                DATAS2 = BYTES42;
            }
           




            ADD_TO_TBDB4(
                        KINDS
                        , SUPPLYS
                        , NAMES                     
                        , SPECS                     
                        , OUTS
                        , COLORS                      
                        , CHECKS
                        , SAVEDAYS                       
                        , COA                      
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                       
                        );

            SEARCH4(textBox4A.Text);
        }
        private void button25_Click(object sender, EventArgs e)
        {
            string ID = textBox4B.Text;
            string KINDS = comboBox4.Text;
            string SUPPLYS = textBox420.Text;
            string NAMES = textBox421.Text;
            string SPECS = textBox422.Text;
            string OUTS = textBox423.Text;
            string COLORS = textBox424.Text;
            string CHECKS = textBox425.Text;
            string SAVEDAYS = textBox426.Text;
            string COA = textBox427.Text;
            string COMMEMTS = textBox428.Text;

            UPDATE_TO_TBDB4(
                                ID
                               , KINDS
                               , SUPPLYS
                               , NAMES
                               , SPECS
                               , OUTS
                               , COLORS
                               , CHECKS
                               , SAVEDAYS
                               , COA
                               , COMMEMTS
                               );

         SEARCH4(textBox4A.Text);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TO_TBDB4(textBox4C.Text);
                SEARCH4(textBox4A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            SEARCH5(textBox5A.Text);
        }
        #endregion


    }
}
