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
        string btnSATUS = null;
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
            comboBox1_load();

            // 綁定事件
            toolStripButton1.Click += BtnAdd_Click;
            toolStripButton2.Click += BtnDelete_Click;
            toolStripButton3.Click += BtnEdit_Click;
            toolStripButton4.Click += BtnSave_Click;
        }

        #region FUNCTION
        public void LoadComboBox(ComboBox comboBox, string sql, string displayMember, string valueMember)
        {
            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                // 解密帳號密碼
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(sql, sqlConn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        comboBox.DataSource = dt;
                        comboBox.DisplayMember = displayMember;  // 顯示的欄位
                        comboBox.ValueMember = valueMember;      // 實際值欄位
                    }
                    else
                    {
                        comboBox.DataSource = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入 ComboBox 發生錯誤：" + ex.Message);
            }
        }
        public void comboBox1_load()
        {
            string sql = @"
                        SELECT '全部' AS 'PARANAME'
                        UNION ALL
                        SELECT 
                        [PARANAME]
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='TB_ORIENTS_CHECKLISTS_CATEGORY'
                        ";

            LoadComboBox(comboBox1, sql, "PARANAME", "PARANAME");
        }

        public void SEARCH(string CATEGORY, string PRODUCTNAME, string SUPPLIER)
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

                //QUERYS
                if (CATEGORY.Equals("全部"))
                {
                    QUERYS.AppendFormat(@"");
                }
                else if(!string.IsNullOrEmpty(CATEGORY) && !CATEGORY.Equals("全部"))
                {
                    QUERYS.AppendFormat(@" AND CATEGORY='{0}' ", CATEGORY);
                }
                else
                {
                    QUERYS.AppendFormat(@"");
                }
                //QUERYS2
                if (!string.IsNullOrEmpty(PRODUCTNAME))
                {
                    QUERYS2.AppendFormat(@" AND PRODUCTNAME LIKE '%{0}%' ", PRODUCTNAME);
                }
                else
                {
                    QUERYS2.AppendFormat(@"");
                }
                //QUERYS3
                if (!string.IsNullOrEmpty(SUPPLIER))
                {
                    QUERYS3.AppendFormat(@"  AND SUPPLIER LIKE '%{0}%' ", SUPPLIER);
                }
                else
                {
                    QUERYS3.AppendFormat(@"");
                }


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
                                    {0}
                                    {1}
                                    {2}
                                    ORDER BY CAST(LEFT(CATEGORY, CHARINDEX('.', CATEGORY) - 1) AS INT),SUPPLIER,PRODUCTNAME
                                    ", QUERYS.ToString(), QUERYS2.ToString(), QUERYS3.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds = new DataSet(); // 這樣就不需要再 Clear()
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    DataTable dt = ds.Tables["ds"];
                    dataGridView1.DataSource = dt;

                    // 先移除舊的 ImageColumn
                    if (dataGridView1.Columns.Contains("ICON_IMAGE"))
                        dataGridView1.Columns.Remove("ICON_IMAGE");

                    // 新增 ImageColumn（顯示縮圖）
                    DataGridViewImageColumn imgCol = new DataGridViewImageColumn();
                    imgCol.Name = "ICON_IMAGE";
                    imgCol.HeaderText = "圖片";
                    imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
                    dataGridView1.Columns.Add(imgCol);

                    // 填入縮圖
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow) continue;

                        string path = row.Cells["圖片路徑"].Value?.ToString();
                        if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        {
                            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                            {
                                row.Cells["ICON_IMAGE"].Value = Image.FromStream(stream);
                            }
                        }
                    }

                    //dataGridView1.Columns["分類"].Width = 100;
                    //dataGridView1.Columns["更新時間"].Width = 100;
                    //dataGridView1.Columns["供應商"].Width = 100;
                    // 4. 自動將所有欄位寬度設為 100
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        col.Width = 100;
                    }

                    //dataGridView1.AutoResizeColumns();                   
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
                string CATEGORY = dataGridView1.CurrentRow.Cells["分類"].Value.ToString();
                string SUPPLIER = dataGridView1.CurrentRow.Cells["供應商"].Value.ToString();
                string PRODUCTNAME = dataGridView1.CurrentRow.Cells["品名"].Value.ToString();
                string INGREDIENT_CN = dataGridView1.CurrentRow.Cells["成分(中文)"].Value.ToString();
                string INGREDIENT_EN = dataGridView1.CurrentRow.Cells["成分(英文)"].Value.ToString();
                string PRODUCT_ALLERGEN = dataGridView1.CurrentRow.Cells["產品過敏原"].Value.ToString();
                string LINE_ALLERGEN = dataGridView1.CurrentRow.Cells["產線過敏原"].Value.ToString();
                string ORIGIN = dataGridView1.CurrentRow.Cells["產地"].Value.ToString();
                string PACKAGE_SPEC = dataGridView1.CurrentRow.Cells["外包裝及驗收標準"].Value.ToString();
                string PRODUCT_APPEARANCE = dataGridView1.CurrentRow.Cells["產品外觀"].Value.ToString();
                string COLOR = dataGridView1.CurrentRow.Cells["色澤"].Value.ToString();
                string FLAVOR = dataGridView1.CurrentRow.Cells["風味"].Value.ToString();
                string BATCHNO = dataGridView1.CurrentRow.Cells["產品批號"].Value.ToString();
                string UNIT_WEIGHT = dataGridView1.CurrentRow.Cells["單位重量"].Value.ToString();
                string SHELFLIFE = dataGridView1.CurrentRow.Cells["保存期限"].Value.ToString();
                string STORAGE_CONDITION = dataGridView1.CurrentRow.Cells["保存條件"].Value.ToString();
                string GMO_STATUS = dataGridView1.CurrentRow.Cells["基改/非基改"].Value.ToString();
                string HAS_COA = dataGridView1.CurrentRow.Cells["檢附COA"].Value.ToString();
                string INSPECTION_FREQUENCY = dataGridView1.CurrentRow.Cells["抽驗頻率"].Value.ToString();
                string BRIX = dataGridView1.CurrentRow.Cells["糖度(Brix)"].Value.ToString();
                string REMARK = dataGridView1.CurrentRow.Cells["備註"].Value.ToString();


                textBox2.Text = id;
                textBox2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                textBox13.Text = "";
                textBox14.Text = "";
                textBox15.Text = "";
                textBox16.Text = "";
                textBox17.Text = "";
                textBox18.Text = "";
                textBox19.Text = "";
                textBox20.Text = "";
                textBox21.Text = "";
                textBox22.Text = "";
                textBox23.Text = "";

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
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
    

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string CATEGORY = comboBox1.Text.ToString();
            string PRODUCTNAME = textBox1.Text.Trim();
            string SUPPLIER = textBox3.Text.Trim();

            SEARCH(CATEGORY, PRODUCTNAME, SUPPLIER);
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


        private void BtnAdd_Click(object sender, EventArgs e)
        {
            btnSATUS = "ADD";

            //MessageBox.Show("新增功能");
            // TODO: 開啟新增模式，清空輸入欄位
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            btnSATUS = "DELETE";

            //MessageBox.Show("刪除功能");
            // TODO: 確認後刪除選中的資料
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            btnSATUS = "EDIT";

            //MessageBox.Show("修改功能");
            // TODO: 載入選中資料，進入編輯模式
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
           
            if(btnSATUS == null)
            {
                MessageBox.Show("沒有按 功能");
                return;  // 提早結束，不會往下執行
            }
            if (btnSATUS.Equals("ADD") )
            {
                MessageBox.Show("ADD");
            }
            else if(btnSATUS.Equals("DELETE"))
            {
                MessageBox.Show("DELETE");
            }
            else if(btnSATUS.Equals("EDIT"))
            {
                MessageBox.Show("EDIT");
            }

            btnSATUS = null;

            // TODO: 將目前資料存入資料庫
        }
        #endregion


    }
}
