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

        public FrmRESEARCHDB()
        {
            InitializeComponent();

            SEARCH(textBox1A.Text.Trim());
            SETdataGridView1();
            SEARCH2(textBox2A.Text.Trim());
            SETdataGridView2();
        }

        #region FUNCTION

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
                                        ,[DATAS] 
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
                                        ,[DATAS] 
                                        FROM [TKRESEARCH].[dbo].[TBDB2]
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
            lnkDownload.Name = "lnkDownload";
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

            SEARCH(textBox1A.Text.Trim());
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

            SEARCH(textBox1A.Text.Trim());
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

            SEARCH(textBox1A.Text.Trim());
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
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATE_TO_TBDB1(textBox1B.Text,textBox14.Text,textBox15.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_TO_TBDB1(textBox1C.Text);
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
        #endregion


    }
}
