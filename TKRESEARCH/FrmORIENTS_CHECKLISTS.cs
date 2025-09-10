using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    public partial class FrmORIENTS_CHECKLISTS : Form
    {
        string destFolder = @"\\192.168.1.109\prog更新\TKRESEARCH\IMAGES\FrmORIENTS_CHECKLISTS";

        private SqlConnection conn;
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        string talbename = null;
        int rownum = 0;
        int result;
        string currentID = null;

        public FrmORIENTS_CHECKLISTS()
        {
            InitializeComponent();
        }
        private void FrmORIENTS_CHECKLISTS_Load(object sender, EventArgs e)
        {

        }

        #region FUNCTION
        public void SEARCH(string PRODUCTNAME)
        {
            SqlDataAdapter adapter=new SqlDataAdapter();
            DataSet ds;
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            dataGridView1.DataSource = null;
         

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

                //if (!string.IsNullOrEmpty(ISCLOSED))
                //{
                //    QUERYS.AppendFormat(@"AND ISCLOSED='{0}'", ISCLOSED);
                //}

                //if (!string.IsNullOrEmpty(MB001))
                //{
                //    QUERYS2.AppendFormat(@"AND (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')", MB001);
                //}
                //else
                //{
                //    QUERYS2.AppendFormat(@"");
                //}

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                        ID,
                                        CATEGORY AS '分類',
                                        UPDATETIME AS '更新時間',
                                        SUPPLIER AS '供應商',
                                        PRODUCTNAME AS '品名',
                                        INGREDIENT_CN AS '成分(中文)',
                                        INGREDIENT_EN AS '成分(英文)',
                                        PRODUCT_ALLERGEN AS '產品過敏原',
                                        LINE_ALLERGEN AS '產線過敏原',
                                        ORIGIN AS '產地',
                                        PACKAGE_SPEC AS '外包裝及驗收標準',
                                        PRODUCT_APPEARANCE AS '產品外觀',
                                        COLOR AS '色澤',
                                        FLAVOR AS '風味',
                                        BATCHNO AS '產品批號',
                                        UNIT_WEIGHT AS '單位重量',
                                        SHELFLIFE AS '保存期限',
                                        STORAGE_CONDITION AS '保存條件',
                                        GMO_STATUS AS '基改/非基改',
                                        HAS_COA AS '檢附COA',
                                        INSPECTION_FREQUENCY AS '抽驗頻率',
                                        REMARK AS '備註',
                                        BRIX AS '糖度(Brix)',
                                        ICON AS '圖片路徑'
                                    FROM [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                                    WHERE 1=1
                                    ORDER BY CATEGORY,SUPPLIER,PRODUCTNAME
                                    ", QUERYS.ToString(), QUERYS2.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds = new DataSet(); // 這樣就不需要再 Clear()
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                   
                    //// 先移除舊的 ImageColumn
                    //if (dataGridView1.Columns.Contains("ICON_IMAGE"))
                    //    dataGridView1.Columns.Remove("ICON_IMAGE");

                    //// 新增 ImageColumn（顯示縮圖）
                    //DataGridViewImageColumn imgCol = new DataGridViewImageColumn();
                    //imgCol.Name = "ICON_IMAGE";
                    //imgCol.HeaderText = "圖片";
                    //imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
                    //dataGridView1.Columns.Add(imgCol);

                    //// 填入縮圖
                    //foreach (DataGridViewRow row in dataGridView1.Rows)
                    //{
                    //    if (row.IsNewRow) continue;

                    //    string path = row.Cells["圖片路徑"].Value?.ToString();
                    //    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    //    {
                    //        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                    //        {
                    //            row.Cells["ICON_IMAGE"].Value = Image.FromStream(stream);
                    //        }
                    //    }
                    //}

                    dataGridView1.DataSource = ds.Tables["ds"];
                    dataGridView1.AutoResizeColumns();                   
                }

            }
            catch (Exception EX)
            {

            }
            finally
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                string id = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
               
                textBox2.Text = id;

                // 取圖片路徑
                string iconPath = dataGridView1.CurrentRow.Cells["圖片路徑"].Value?.ToString();

                if (!string.IsNullOrEmpty(iconPath) && File.Exists(iconPath))
                {
                    using (var stream = new FileStream(iconPath, FileMode.Open, FileAccess.Read))
                    {
                        pictureBox1.Image = Image.FromStream(stream);
                    }
                    pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else
                {
                    pictureBox1.Image = null;
                }
            }
        }

        private void SaveImagePathToDB(string ID, string ICON)
        {
            try
            {
                StringBuilder sbSql = new StringBuilder();
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
                                    UPDATE  [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                                    SET 
                                    ICON=@ICON                                    
                                    WHERE ID=@ID
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值
                cmd.Parameters.AddWithValue("@ID", ID);
                cmd.Parameters.AddWithValue("@ICON", ICON);
               

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
        public void SET_TEXTBOX_NULL()
        {
            textBox2.Text = "";

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH("");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string id = textBox2.Text;
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "圖片檔 (*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {

                    string origFileName = Path.GetFileName(ofd.FileName);
                    string ext = Path.GetExtension(origFileName); // 取得副檔名

                    // 生成新的檔名，例如: id + yyyyMMddHHmmss                   
                    string newFileName = string.Format("{0}_{1}{2}",
                                        id,
                                        DateTime.Now.ToString("yyyyMMddHHmmss"),
                                        ext);

                    string destFolder = @"\\192.168.1.109\prog更新\TKRESEARCH\IMAGES\FrmORIENTS_CHECKLISTS";
                    string destPath = Path.Combine(destFolder, newFileName);

                    try
                    {
                        if (!Directory.Exists(destFolder))
                        {
                            Directory.CreateDirectory(destFolder);
                        }

                        // 複製圖片到共用資料夾 (存在則覆蓋)
                        File.Copy(ofd.FileName, destPath, true);

                        // 顯示圖片
                        pictureBox1.ImageLocation = destPath;

                        // 存到資料庫 (這裡示範用完整路徑存)
                        SaveImagePathToDB(id, destPath);

                        MessageBox.Show("圖片已存檔並寫入資料庫！");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("存檔失敗: " + ex.Message);
                    }
                }
            }
        }
        #endregion


    }
}
