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
    public partial class FrmDB1 : Form
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

        public FrmDB1()
        {
            InitializeComponent();

           
        }


        #region FUNCTION
        public void SEARCH(string MB001)
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

                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"  
                                   
                                    ", MB001);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                    SELECT 
                                    [id]
                                    ,[NAME]
                                  

                                    FROM [TKRESEARCH].[dbo].[tblFiles]
                                    ORDER BY [NAME]
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

                    SETdataGridView1();

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
            if (dataGridView1.Columns.Count>=3)
            {
                dataGridView1.Columns.RemoveAt(3);
            }
           

            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload";
            lnkDownload.HeaderText = "Download";
            lnkDownload.Text = "Download";
          
            dataGridView1.Columns.Insert(2, lnkDownload);
            dataGridView1.CellContentClick += new DataGridViewCellEventHandler(DataGridView1_CellClick);
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                int id = Convert.ToInt16((row.Cells[0].Value));
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
                                         [id]
                                        ,[NAME]
                                        ,[CONTENTTYPE]
                                        ,[DATA]
                                        FROM [TKRESEARCH].[dbo].[tblFiles]
                                        where id=@id
                                            ");
                        cmd.CommandText = SQL.ToString();
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            sdr.Read();
                            bytes = (byte[])sdr["Data"];
                            contentType = sdr["ContentType"].ToString();
                            fileName = sdr["Name"].ToString();

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

        private void UploadFile()
        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;
                    Stream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                    BinaryReader br = new BinaryReader(fs); //reads the binary files  
                    Byte[] bytes = br.ReadBytes((Int32)fs.Length); //counting the file length into bytes  
                    //byte[] bytes = File.ReadAllBytes(fileName);
                    string contentType = "";
                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".doc":
                            contentType = "application/msword";
                            break;
                        case ".xls":
                            contentType = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            contentType = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            contentType = "application/pdf";
                            break;
                        case ".jpg":
                            contentType = "image/jpeg";
                            break;
                        case ".png":
                            contentType = "image/png";
                            break;
                        case ".gif":
                            contentType = "image/gif";
                            break;
                        case ".bmp":
                            contentType = "image/bmp";
                            break;
                    }

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
                        string sql = "INSERT INTO [TKRESEARCH].[dbo].[tblFiles] VALUES(@Name, @ContentType, @Data)";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@Name", Path.GetFileName(fileName));
                            cmd.Parameters.AddWithValue("@ContentType", contentType);
                            cmd.Parameters.AddWithValue("@Data", bytes);
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                    SEARCH(textBox1.Text.Trim());
                }
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox1.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UploadFile();
        }


        #endregion


    }
}
