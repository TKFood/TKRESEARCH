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

        public FrmRESEARCHDB()
        {
            InitializeComponent();

            SEARCH(textBox1A.Text.Trim());
            SETdataGridView1();
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

            if (e.RowIndex >= 0)
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
                        cmd.Parameters.AddWithValue("@DOCID", Path.GetFileName(DOCID));
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
        #endregion


    }
}
