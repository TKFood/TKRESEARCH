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
        string currentId = "";

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
            comboBox2_load();
           
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

        public void comboBox2_load()
        {
            string sql = @"
                        SELECT 
                        [PARANAME] AS 'PARANAME'
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='TB_ORIENTS_CHECKLISTS_CATEGORY'
                        ";

            LoadComboBox(comboBox2, sql, "PARANAME", "PARANAME");
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
                    QUERYS2.AppendFormat(@" AND (PRODUCTNAME LIKE '%{0}%' OR MB001 LIKE '%{0}%' )", PRODUCTNAME);
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
                                        MB001 AS '品號',                                      
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
                string MB001 = dataGridView1.CurrentRow.Cells["品號"].Value.ToString();
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
                textBox24.Text = MB001;
                textBox4.Text = SUPPLIER;
                textBox5.Text = PRODUCTNAME;
                textBox6.Text = INGREDIENT_CN;
                textBox7.Text = INGREDIENT_EN;
                textBox8.Text = PRODUCT_ALLERGEN;
                textBox9.Text = LINE_ALLERGEN;
                textBox10.Text = UNIT_WEIGHT;
                textBox11.Text = ORIGIN;
                textBox12.Text = PACKAGE_SPEC;
                textBox13.Text = PRODUCT_APPEARANCE;
                textBox14.Text = COLOR;
                textBox15.Text = FLAVOR;
                textBox16.Text = BATCHNO;
                textBox17.Text = SHELFLIFE;
                textBox18.Text = STORAGE_CONDITION;
                textBox19.Text = GMO_STATUS;
                textBox20.Text = HAS_COA;
                textBox21.Text = INSPECTION_FREQUENCY;
                textBox22.Text = BRIX;
                textBox23.Text = REMARK;
                comboBox2.Text = CATEGORY;

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

        public void RefreshData(string currentId)
        {            
            // 3. 回到原本那筆資料
            if (!string.IsNullOrEmpty(currentId))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["ID"].Value.ToString() == currentId)
                    {
                        row.Selected = true;
                        dataGridView1.CurrentCell = row.Cells[0];

                        dataGridView1_SelectionChanged(dataGridView1, EventArgs.Empty);
                   
                        break;
                    }
                }
            }
        }

        public void ADD_TB_ORIENTS_CHECKLISTS(
            string ID,
            string MB001,
            string CATEGORY,
            string SUPPLIER,
            string PRODUCTNAME,
            string INGREDIENT_CN,
            string INGREDIENT_EN,
            string PRODUCT_ALLERGEN,
            string LINE_ALLERGEN,
            string ORIGIN,
            string PACKAGE_SPEC,
            string PRODUCT_APPEARANCE,
            string COLOR,
            string FLAVOR,
            string BATCHNO,
            string UNIT_WEIGHT,
            string SHELFLIFE,
            string STORAGE_CONDITION,
            string GMO_STATUS,
            string HAS_COA,
            string INSPECTION_FREQUENCY,
            string REMARK,
            string BRIX
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

                string sql = @"
                            INSERT INTO [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                            ([MB001],[CATEGORY],[SUPPLIER],[PRODUCTNAME],[INGREDIENT_CN],[INGREDIENT_EN],
                             [PRODUCT_ALLERGEN],[LINE_ALLERGEN],[ORIGIN],[PACKAGE_SPEC],[PRODUCT_APPEARANCE],
                             [COLOR],[FLAVOR],[BATCHNO],[UNIT_WEIGHT],[SHELFLIFE],[STORAGE_CONDITION],
                             [GMO_STATUS],[HAS_COA],[INSPECTION_FREQUENCY],[REMARK],[BRIX],[UPDATETIME])
                            VALUES
                            (@MB001,@CATEGORY,@SUPPLIER,@PRODUCTNAME,@INGREDIENT_CN,@INGREDIENT_EN,
                             @PRODUCT_ALLERGEN,@LINE_ALLERGEN,@ORIGIN,@PACKAGE_SPEC,@PRODUCT_APPEARANCE,
                             @COLOR,@FLAVOR,@BATCHNO,@UNIT_WEIGHT,@SHELFLIFE,@STORAGE_CONDITION,
                             @GMO_STATUS,@HAS_COA,@INSPECTION_FREQUENCY,@REMARK,@BRIX,@UPDATETIME)
                            ";

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@MB001", MB001);
                    cmd.Parameters.AddWithValue("@CATEGORY", CATEGORY);
                    cmd.Parameters.AddWithValue("@SUPPLIER", SUPPLIER);
                    cmd.Parameters.AddWithValue("@PRODUCTNAME", PRODUCTNAME);
                    cmd.Parameters.AddWithValue("@INGREDIENT_CN", INGREDIENT_CN);
                    cmd.Parameters.AddWithValue("@INGREDIENT_EN", INGREDIENT_EN);
                    cmd.Parameters.AddWithValue("@PRODUCT_ALLERGEN", PRODUCT_ALLERGEN);
                    cmd.Parameters.AddWithValue("@LINE_ALLERGEN", LINE_ALLERGEN);
                    cmd.Parameters.AddWithValue("@ORIGIN", ORIGIN);
                    cmd.Parameters.AddWithValue("@PACKAGE_SPEC", PACKAGE_SPEC);
                    cmd.Parameters.AddWithValue("@PRODUCT_APPEARANCE", PRODUCT_APPEARANCE);
                    cmd.Parameters.AddWithValue("@COLOR", COLOR);
                    cmd.Parameters.AddWithValue("@FLAVOR", FLAVOR);
                    cmd.Parameters.AddWithValue("@BATCHNO", BATCHNO);
                    cmd.Parameters.AddWithValue("@UNIT_WEIGHT", UNIT_WEIGHT);
                    cmd.Parameters.AddWithValue("@SHELFLIFE", SHELFLIFE);
                    cmd.Parameters.AddWithValue("@STORAGE_CONDITION", STORAGE_CONDITION);
                    cmd.Parameters.AddWithValue("@GMO_STATUS", GMO_STATUS);
                    cmd.Parameters.AddWithValue("@HAS_COA", HAS_COA);
                    cmd.Parameters.AddWithValue("@INSPECTION_FREQUENCY", INSPECTION_FREQUENCY);
                    cmd.Parameters.AddWithValue("@REMARK", REMARK);
                    cmd.Parameters.AddWithValue("@BRIX", BRIX);
                    cmd.Parameters.AddWithValue("@UPDATETIME", DateTime.Now.ToString("yyyy/MM/dd"));

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }

            }
            catch(Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }
            
        }
        public void UPDATE_TB_ORIENTS_CHECKLISTS(
            string ID,
            string MB001,
            string CATEGORY,
            string SUPPLIER,
            string PRODUCTNAME,
            string INGREDIENT_CN,
            string INGREDIENT_EN,
            string PRODUCT_ALLERGEN,
            string LINE_ALLERGEN,
            string ORIGIN,
            string PACKAGE_SPEC,
            string PRODUCT_APPEARANCE,
            string COLOR,
            string FLAVOR,
            string BATCHNO,
            string UNIT_WEIGHT,
            string SHELFLIFE,
            string STORAGE_CONDITION,
            string GMO_STATUS,
            string HAS_COA,
            string INSPECTION_FREQUENCY,
            string REMARK,
            string BRIX

            )
        {
            try
            {
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                string sql = @"
                            UPDATE [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                            SET
                            [MB001]=@MB001,
                            [CATEGORY]=@CATEGORY,[SUPPLIER]=@SUPPLIER,[PRODUCTNAME]=@PRODUCTNAME,
                            [INGREDIENT_CN]=@INGREDIENT_CN,[INGREDIENT_EN]=@INGREDIENT_EN,
                            [PRODUCT_ALLERGEN]=@PRODUCT_ALLERGEN,[LINE_ALLERGEN]=@LINE_ALLERGEN,
                            [ORIGIN]=@ORIGIN,[PACKAGE_SPEC]=@PACKAGE_SPEC,[PRODUCT_APPEARANCE]=@PRODUCT_APPEARANCE,
                            [COLOR]=@COLOR,[FLAVOR]=@FLAVOR,[BATCHNO]=@BATCHNO,[UNIT_WEIGHT]=@UNIT_WEIGHT,
                            [SHELFLIFE]=@SHELFLIFE,[STORAGE_CONDITION]=@STORAGE_CONDITION,[GMO_STATUS]=@GMO_STATUS,
                            [HAS_COA]=@HAS_COA,[INSPECTION_FREQUENCY]=@INSPECTION_FREQUENCY,[REMARK]=@REMARK,
                            [BRIX]=@BRIX,[UPDATETIME]=@UPDATETIME
                            WHERE [ID]=@ID"
                            ;

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", ID);
                    cmd.Parameters.AddWithValue("@MB001", MB001);
                    cmd.Parameters.AddWithValue("@CATEGORY", CATEGORY);
                    cmd.Parameters.AddWithValue("@SUPPLIER", SUPPLIER);
                    cmd.Parameters.AddWithValue("@PRODUCTNAME", PRODUCTNAME);
                    cmd.Parameters.AddWithValue("@INGREDIENT_CN", INGREDIENT_CN);
                    cmd.Parameters.AddWithValue("@INGREDIENT_EN", INGREDIENT_EN);
                    cmd.Parameters.AddWithValue("@PRODUCT_ALLERGEN", PRODUCT_ALLERGEN);
                    cmd.Parameters.AddWithValue("@LINE_ALLERGEN", LINE_ALLERGEN);
                    cmd.Parameters.AddWithValue("@ORIGIN", ORIGIN);
                    cmd.Parameters.AddWithValue("@PACKAGE_SPEC", PACKAGE_SPEC);
                    cmd.Parameters.AddWithValue("@PRODUCT_APPEARANCE", PRODUCT_APPEARANCE);
                    cmd.Parameters.AddWithValue("@COLOR", COLOR);
                    cmd.Parameters.AddWithValue("@FLAVOR", FLAVOR);
                    cmd.Parameters.AddWithValue("@BATCHNO", BATCHNO);
                    cmd.Parameters.AddWithValue("@UNIT_WEIGHT", UNIT_WEIGHT);
                    cmd.Parameters.AddWithValue("@SHELFLIFE", SHELFLIFE);
                    cmd.Parameters.AddWithValue("@STORAGE_CONDITION", STORAGE_CONDITION);
                    cmd.Parameters.AddWithValue("@GMO_STATUS", GMO_STATUS);
                    cmd.Parameters.AddWithValue("@HAS_COA", HAS_COA);
                    cmd.Parameters.AddWithValue("@INSPECTION_FREQUENCY", INSPECTION_FREQUENCY);
                    cmd.Parameters.AddWithValue("@REMARK", REMARK);
                    cmd.Parameters.AddWithValue("@BRIX", BRIX);
                    cmd.Parameters.AddWithValue("@UPDATETIME", DateTime.Now.ToString("yyyy/MM/dd"));

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }
            
        }
        public void DELETE_TB_ORIENTS_CHECKLISTS(string ID)
        {
            try
            {
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
                string sql = "DELETE FROM [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS] WHERE [ID]=@ID";

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", ID);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch(Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }
           
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            // 1. 檢查輸入是否為空（如果允許空值則跳過）
            if (string.IsNullOrEmpty(textBox24.Text.Trim()))
            {
                return;
            }

            // 2. 執行耗時的資料庫查詢
            DataTable DT = FIND_TB_ORIENTS_CHECKLISTS_REPEATS(textBox24.Text.Trim());

            // 3. 判斷重複
            if (DT != null && DT.Rows.Count >= 2)
            {
                // 發現重複時，彈出警告
                MessageBox.Show("品號:" + textBox24.Text.Trim() + " 有重複 " + DT.Rows.Count + "筆資料, 請修改!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // 【可選】如果強制要求修改，可以設定 e.Cancel = true
                // e.Cancel = true; // 這會阻止使用者離開 textBox24
            }

        }

        public DataTable FIND_TB_ORIENTS_CHECKLISTS_REPEATS(string MB001)
        {
            // 初始化連線資訊和解密（保留您的邏輯）
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            // 資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            // 【SQL 修正】：使用視窗函數確保只回傳出現次數 > 1 的記錄
            string sqlQuery = @"
                                SELECT 
                                    [ID], [MB001], [CATEGORY], [SUPPLIER], [PRODUCTNAME], [INGREDIENT_CN], 
                                    [INGREDIENT_EN], [PRODUCT_ALLERGEN], [LINE_ALLERGEN], [ORIGIN], 
                                    [PACKAGE_SPEC], [PRODUCT_APPEARANCE], [COLOR], [FLAVOR], [BATCHNO], 
                                    [UNIT_WEIGHT], [SHELFLIFE], [STORAGE_CONDITION], [GMO_STATUS], 
                                    [HAS_COA], [INSPECTION_FREQUENCY], [REMARK], [BRIX], [ICON], 
                                    [UPDATETIME]
                                FROM 
                                    (
                                        SELECT 
                                            *,
                                            -- 使用 PARTITION BY 計算每個 MB001 出現的次數
                                            COUNT([MB001]) OVER (PARTITION BY [MB001]) AS TotalCount 
                                        FROM 
                                            [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                                        -- 在子查詢中，先限制 MB001 的範圍，優化性能
                                        WHERE 
                                            [MB001] = @MB001  
                                    ) AS Subquery
                                WHERE 
                                    TotalCount >= 1  -- 篩選出所有重複（出現次數大於 1）的記錄
                                ORDER BY
                                    [MB001], [ID];
                            ";

            DataTable dtResult = new DataTable();

            try
            {
                // 1. 建立連線物件 (使用解密後的連線字串)
                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    // 2. 建立指令物件 (SqlCommand)
                    using (SqlCommand cmd = new SqlCommand(sqlQuery, conn))
                    {
                        // 3. 參數化處理：新增 SQL 參數
                        // 這是正確新增參數的地方，且使用了 SQL 語句中的 @MB001
                        cmd.Parameters.AddWithValue("@MB001", MB001);

                        // 4. 建立資料配接器，並將指令物件傳入
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            // 5. 執行查詢並將結果填入 dtResult
                            adapter.Fill(dtResult);
                        }
                    } // cmd.Dispose()
                } // conn.Close() and conn.Dispose()

                return dtResult;
            }
            catch (Exception ex)
            {
                // 處理任何連線或查詢錯誤
                // 建議保留 MessageBox 或其他日誌記錄
                // MessageBox.Show("查詢資料庫錯誤: " + ex.Message, "資料庫錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 在錯誤發生時傳回 null 或一個空的 DataTable
                return null;
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
            textBox24.Text = "";
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


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            btnSATUS = "ADD";

            string ID = textBox2.Text;
            string MB001 = textBox24.Text;
            string CATEGORY = comboBox2.Text.ToString();
            string SUPPLIER = textBox4.Text;
            string PRODUCTNAME = textBox5.Text;
            string INGREDIENT_CN = textBox6.Text;
            string INGREDIENT_EN = textBox7.Text;
            string PRODUCT_ALLERGEN = textBox8.Text;
            string LINE_ALLERGEN = textBox9.Text;
            string ORIGIN = textBox11.Text;
            string PACKAGE_SPEC = textBox12.Text;
            string PRODUCT_APPEARANCE = textBox13.Text;
            string COLOR = textBox14.Text;
            string FLAVOR = textBox15.Text;
            string BATCHNO = textBox16.Text;
            string UNIT_WEIGHT = textBox10.Text;
            string SHELFLIFE = textBox17.Text;
            string STORAGE_CONDITION = textBox18.Text;
            string GMO_STATUS = textBox19.Text;
            string HAS_COA = textBox20.Text;
            string INSPECTION_FREQUENCY = textBox21.Text;
            string BRIX = textBox22.Text;
            string REMARK = textBox23.Text;

            ADD_TB_ORIENTS_CHECKLISTS(
                ID,
                MB001,
                CATEGORY,
                SUPPLIER,
                PRODUCTNAME,
                INGREDIENT_CN,
                INGREDIENT_EN,
                PRODUCT_ALLERGEN,
                LINE_ALLERGEN,
                ORIGIN,
                PACKAGE_SPEC,
                PRODUCT_APPEARANCE,
                COLOR,
                FLAVOR,
                BATCHNO,
                UNIT_WEIGHT,
                SHELFLIFE,
                STORAGE_CONDITION,
                GMO_STATUS,
                HAS_COA,
                INSPECTION_FREQUENCY,
                REMARK,
                BRIX
            );
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            btnSATUS = "EDIT";
            currentId = textBox2.Text;

            string ID = textBox2.Text;
            string MB001 = textBox24.Text;
            string CATEGORY = comboBox2.Text.ToString();
            string SUPPLIER = textBox4.Text;
            string PRODUCTNAME = textBox5.Text;
            string INGREDIENT_CN = textBox6.Text;
            string INGREDIENT_EN = textBox7.Text;
            string PRODUCT_ALLERGEN = textBox8.Text;
            string LINE_ALLERGEN = textBox9.Text;
            string ORIGIN = textBox11.Text;
            string PACKAGE_SPEC = textBox12.Text;
            string PRODUCT_APPEARANCE = textBox13.Text;
            string COLOR = textBox14.Text;
            string FLAVOR = textBox15.Text;
            string BATCHNO = textBox16.Text;
            string UNIT_WEIGHT = textBox10.Text;
            string SHELFLIFE = textBox17.Text;
            string STORAGE_CONDITION = textBox18.Text;
            string GMO_STATUS = textBox19.Text;
            string HAS_COA = textBox20.Text;
            string INSPECTION_FREQUENCY = textBox21.Text;
            string BRIX = textBox22.Text;
            string REMARK = textBox23.Text;

            UPDATE_TB_ORIENTS_CHECKLISTS(
                ID,
                MB001,
                CATEGORY,
                SUPPLIER,
                PRODUCTNAME,
                INGREDIENT_CN,
                INGREDIENT_EN,
                PRODUCT_ALLERGEN,
                LINE_ALLERGEN,
                ORIGIN,
                PACKAGE_SPEC,
                PRODUCT_APPEARANCE,
                COLOR,
                FLAVOR,
                BATCHNO,
                UNIT_WEIGHT,
                SHELFLIFE,
                STORAGE_CONDITION,
                GMO_STATUS,
                HAS_COA,
                INSPECTION_FREQUENCY,
                REMARK,
                BRIX
            );

            string SEARCH_CATEGORY = comboBox1.Text.ToString();
            string SEARCH_PRODUCTNAME = textBox1.Text.Trim();
            string SEARCH_SUPPLIER = textBox3.Text.Trim();

            SEARCH(SEARCH_CATEGORY, SEARCH_PRODUCTNAME, SEARCH_SUPPLIER);
            MessageBox.Show("資料已修改成功！");

            //回到原本那筆資料
            RefreshData(currentId);
            btnSATUS = null;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                {
                    string id = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                    string DEL_PRODUCTNAME = dataGridView1.CurrentRow.Cells["品名"].Value.ToString();
                    // 顯示確認訊息
                    var result = MessageBox.Show(
                        $"確定要刪除這筆資料嗎？\n  品名 = {DEL_PRODUCTNAME}",
                        "刪除確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.Yes)
                    {
                        DELETE_TB_ORIENTS_CHECKLISTS(id);

                        string CATEGORY = comboBox1.Text.ToString();
                        string PRODUCTNAME = textBox1.Text.Trim();
                        string SUPPLIER = textBox3.Text.Trim();

                        SEARCH(CATEGORY, PRODUCTNAME, SUPPLIER);
                        MessageBox.Show("資料已刪除成功！");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);
            }
        }

        #endregion

     
    }
}
