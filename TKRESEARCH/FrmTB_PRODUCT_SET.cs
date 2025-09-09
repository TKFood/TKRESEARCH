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
    public partial class FrmTB_PRODUCT_SET : Form
    {
        private SqlDataAdapter adapter_TB_PRODUCT_SET_M;
        private DataSet ds_TB_PRODUCT_SET_M;
        private SqlDataAdapter adapter_TB_PRODUCT_SET_D;
        private DataSet ds_TB_PRODUCT_SET_D;

        private SqlConnection conn;
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;      
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;      
        string talbename = null;
        int rownum = 0;
        int result;
        string currentMID = null;

        public FrmTB_PRODUCT_SET()
        {
            InitializeComponent();

          
        }
        private void FrmTB_PRODUCT_SET_Load(object sender, EventArgs e)
        {
            comboBox1_load();
            comboBox2_load();
            comboBox3_load();
            comboBox4_load();
            DATAGRIDSET();
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
                        SELECT 
                            [ID],
                            [KIND],
                            [PARAID],
                            [PARANAME]
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='FrmTB_PRODUCT_SET'
                        ORDER BY [PARAID]";

            LoadComboBox(comboBox1, sql, "PARANAME", "PARAID");
        }
        public void comboBox2_load()
        {
            string sql = @"
                        SELECT 
                            [ID],
                            [KIND],
                            [PARAID],
                            [PARANAME]
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='FrmTB_PRODUCT_SET'
                        ORDER BY [PARAID]";

            LoadComboBox(comboBox2, sql, "PARANAME", "PARAID");
        }
        public void comboBox3_load()
        {
            string sql = @"
                        SELECT 
                            [ID],
                            [KIND],
                            [PARAID],
                            [PARANAME]
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='FrmTB_PRODUCT_SET_KINDS'
                        ORDER BY [PARAID]";

            LoadComboBox(comboBox3, sql, "PARANAME", "PARAID");
        }
        public void comboBox4_load()
        {
            string sql = @"
                        SELECT 
                            [ID],
                            [KIND],
                            [PARAID],
                            [PARANAME]
                        FROM [TKRESEARCH].[dbo].[TBPARA]
                        WHERE [KIND]='FrmTB_PRODUCT_SET'
                        ORDER BY [PARAID]";

            LoadComboBox(comboBox4, sql, "PARANAME", "PARAID");
        }

        public void DATAGRIDSET()
        {
            // DataGridView1 屬性設定
            dataGridView1.AllowUserToAddRows = true;   // 允許新增
            dataGridView1.AllowUserToDeleteRows = true; // 允許刪除
            dataGridView1.ReadOnly = false;             // 可編輯
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            // dataGridView2 屬性設定
            dataGridView2.AllowUserToAddRows = true;   // 允許新增
            dataGridView2.AllowUserToDeleteRows = true; // 允許刪除
            dataGridView2.ReadOnly = false;             // 可編輯
            dataGridView2.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }
        public void SEARCH(string ISCLOSED,string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();           
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;

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

                if(!string.IsNullOrEmpty(ISCLOSED))
                {
                    QUERYS.AppendFormat(@"AND ISCLOSED='{0}'", ISCLOSED);
                }

                if (!string.IsNullOrEmpty(MB001))
                {
                    QUERYS2.AppendFormat(@"AND (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')", MB001);
                }
                else
                {
                    QUERYS2.AppendFormat(@"");
                }

                sbSql.Clear();              

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MID]
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ", QUERYS.ToString(), QUERYS2.ToString());

                adapter_TB_PRODUCT_SET_M = new SqlDataAdapter(@"" + sbSql, sqlConn);

                // -------- 手動設定 InsertCommand --------
                adapter_TB_PRODUCT_SET_M.InsertCommand = new SqlCommand(@"
                    INSERT INTO [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M] 
                    (MB001, MB002) 
                    VALUES (@MB001, @MB002);
                    SELECT SCOPE_IDENTITY();", sqlConn);
                adapter_TB_PRODUCT_SET_M.InsertCommand.Parameters.Add("@MB001", SqlDbType.NVarChar, 0, "品號");
                adapter_TB_PRODUCT_SET_M.InsertCommand.Parameters.Add("@MB002", SqlDbType.NVarChar, 0, "品名");

                // -------- 手動設定 UpdateCommand --------
                adapter_TB_PRODUCT_SET_M.UpdateCommand = new SqlCommand(@"
                    UPDATE [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M] 
                    SET MB001=@MB001, MB002=@MB002 
                    WHERE MID=@MID", sqlConn);
                adapter_TB_PRODUCT_SET_M.UpdateCommand.Parameters.Add("@MB001", SqlDbType.NVarChar, 0, "品號");
                adapter_TB_PRODUCT_SET_M.UpdateCommand.Parameters.Add("@MB002", SqlDbType.NVarChar, 0, "品名");
                adapter_TB_PRODUCT_SET_M.UpdateCommand.Parameters.Add("@MID", SqlDbType.Int, 0, "MID");

                // -------- 手動設定 DeleteCommand --------
                adapter_TB_PRODUCT_SET_M.DeleteCommand = new SqlCommand(@"
                    DELETE FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M] 
                    WHERE MID=@MID", sqlConn);
                adapter_TB_PRODUCT_SET_M.DeleteCommand.Parameters.Add("@MID", SqlDbType.Int, 0, "MID");


                sqlCmdBuilder = new SqlCommandBuilder(adapter_TB_PRODUCT_SET_M);
                sqlConn.Open();
                ds_TB_PRODUCT_SET_M = new DataSet(); // 這樣就不需要再 Clear()
                adapter_TB_PRODUCT_SET_M.Fill(ds_TB_PRODUCT_SET_M, "ds_TB_PRODUCT_SET_M");
                sqlConn.Close();

                if (ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"];
                    //dataGridView1.AutoResizeColumns();
                    // 指定固定寬度
                    dataGridView1.Columns["品號"].Width = 200;
                    dataGridView1.Columns["品名"].Width = 200;

                    // 查詢後，還原到剛剛那筆dataGridView1
                    // 同時指向明細
                    if (!string.IsNullOrEmpty(currentMID))
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells["MID"].Value.ToString() == currentMID)
                            {
                                row.Selected = true;
                                dataGridView1.CurrentCell = row.Cells["品號"];

                                // 直接呼叫明細查詢，更新明細 DataGridView2
                                SEARCH_TB_PRODUCT_SET_D(currentMID);

                                break;
                            }
                        }
                    }
                }

            }
            catch (Exception EX)
            {

            }
            finally
            {

            }
        }

        public void SEARCH_TB_PRODUCT_SET_D(string MID)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            dataGridView2.DataSource = null;
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

                sbSql.AppendFormat(@"  
                                SELECT 
                                [MID]
                                ,[SERNO] AS '序號'
                                ,[MB001] AS '品號'
                                ,[MB002] AS '品名'
                                ,[AMOUNTS] AS '用量'
                                ,[UNITS] AS '單位'
                                FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D]
                                WHERE MID='{0}' ", MID);

                // 在建立 adapter_TB_PRODUCT_SET_D 之後，設定三個 command
                adapter_TB_PRODUCT_SET_D = new SqlDataAdapter(sbSql.ToString(), sqlConn);

                // Insert
                adapter_TB_PRODUCT_SET_D.InsertCommand = new SqlCommand(
                    @"INSERT INTO [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D] 
                    (MID, SERNO, MB001, MB002, AMOUNTS, UNITS) 
                     VALUES (@MID, @SERNO, @MB001, @MB002, @AMOUNTS, @UNITS)", sqlConn);
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@MID", SqlDbType.Int).SourceColumn = "MID";
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@SERNO", SqlDbType.NVarChar, 20).SourceColumn = "序號";
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@MB001", SqlDbType.NVarChar, 50).SourceColumn = "品號";
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@MB002", SqlDbType.NVarChar, 200).SourceColumn = "品名";
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@AMOUNTS", SqlDbType.NVarChar, 200).SourceColumn = "用量";
                adapter_TB_PRODUCT_SET_D.InsertCommand.Parameters.Add("@UNITS", SqlDbType.NVarChar, 20).SourceColumn = "單位";

                // Update — 注意把 WHERE 的 key 參數設為 Original
                adapter_TB_PRODUCT_SET_D.UpdateCommand = new SqlCommand(
                    @"UPDATE [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D]
                      SET MB001 = @MB001,
                          MB002 = @MB002,
                          AMOUNTS = @AMOUNTS,
                          UNITS = @UNITS
                      WHERE MID = @MID AND SERNO = @SERNO", sqlConn);

                adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@MB001", SqlDbType.NVarChar, 50).SourceColumn = "品號";
                adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@MB002", SqlDbType.NVarChar, 200).SourceColumn = "品名";
                adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@AMOUNTS", SqlDbType.NVarChar, 200).SourceColumn = "用量";
                adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@UNITS", SqlDbType.NVarChar, 20).SourceColumn = "單位";

                // Key 參數 (用 Original 版本以匹配更新前的 key)
                var pMID = adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@MID", SqlDbType.Int);
                pMID.SourceColumn = "MID";
                pMID.SourceVersion = DataRowVersion.Original;

                var pSERNO = adapter_TB_PRODUCT_SET_D.UpdateCommand.Parameters.Add("@SERNO", SqlDbType.NVarChar, 20);
                pSERNO.SourceColumn = "序號";
                pSERNO.SourceVersion = DataRowVersion.Original;

                // Delete
                adapter_TB_PRODUCT_SET_D.DeleteCommand = new SqlCommand(
                    @"DELETE FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D] WHERE MID = @MID AND SERNO = @SERNO", sqlConn);
                var dMID = adapter_TB_PRODUCT_SET_D.DeleteCommand.Parameters.Add("@MID", SqlDbType.Int);
                dMID.SourceColumn = "MID";
                dMID.SourceVersion = DataRowVersion.Original;
                var dSERNO = adapter_TB_PRODUCT_SET_D.DeleteCommand.Parameters.Add("@SERNO", SqlDbType.NVarChar, 20);
                dSERNO.SourceColumn = "序號";
                dSERNO.SourceVersion = DataRowVersion.Original; ;

                ds_TB_PRODUCT_SET_D = new DataSet();
                adapter_TB_PRODUCT_SET_D.Fill(ds_TB_PRODUCT_SET_D, "ds_TB_PRODUCT_SET_D");

                sqlCmdBuilder = new SqlCommandBuilder(adapter_TB_PRODUCT_SET_D);
                sqlConn.Open();
                ds_TB_PRODUCT_SET_D = new DataSet(); // 這樣就不需要再 Clear()
                adapter_TB_PRODUCT_SET_D.Fill(ds_TB_PRODUCT_SET_D, "ds_TB_PRODUCT_SET_D");
                sqlConn.Close();

                if (ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"];
                    //dataGridView2.AutoResizeColumns();
                    //指定固定寬度
                    dataGridView2.Columns["序號"].Width = 100;
                    dataGridView2.Columns["品號"].Width = 200;
                    dataGridView2.Columns["品名"].Width = 200;
                    dataGridView2.Columns["用量"].Width = 100;
                    dataGridView2.Columns["單位"].Width = 100;
                }

                // ✅ 即使沒有資料，也要建立 DataTable 結構並綁定              
                if (ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"].Rows.Count == 0)
                {
                    // 建立一個空的 Row，不是真的資料，只是確保 DataGridView 可以編輯
                    ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"].Rows.Clear();
                    dataGridView2.DataSource = ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"];
                    //dataGridView2.AutoResizeColumns();
                   
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
                string mid = dataGridView1.CurrentRow.Cells["MID"].Value.ToString();
                SEARCH_TB_PRODUCT_SET_D(mid);

                string MB001 = dataGridView1.CurrentRow.Cells["品號"].Value.ToString();
                string MB002 = dataGridView1.CurrentRow.Cells["品名"].Value.ToString();
                textBox2.Text = mid;
                textBox3.Text = MB001;
                textBox4.Text = MB002;
            }
        }
        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            // 這裡假設主檔 DataGridView 是 dataGridView1，MID 在目前選取列
            if (dataGridView1.CurrentRow == null) return;

            string masterMID = dataGridView1.CurrentRow.Cells["MID"].Value.ToString();
            // 每次有新列被加進來，就重新編流水號
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].IsNewRow) continue; // 跳過新增中的空白列

                int serno = (i + 1) * 10; // 1筆=10, 2筆=20, ...
                dataGridView2.Rows[i].Cells["序號"].Value = serno.ToString("D4"); // 4碼補零

                // 設定跟主檔一樣的 MID
                dataGridView2.Rows[i].Cells["MID"].Value = masterMID;
            }
        }

        public void SET_dataGridView1_MID()
        {
            if (dataGridView1.CurrentRow != null)
            {
                // 先記住目前選到的 MID
                currentMID = dataGridView1.CurrentRow.Cells["MID"].Value.ToString();              
                //MessageBox.Show(currentMID);
            }
        }

        public void SEARCH_TB_PRODUCT_SET_M(string MID)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlDataAdapter adapter;
            DataSet ds;

            dataGridView3.DataSource = null;

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

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MID]
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[ISCLOSED] AS '結案碼'
                                    ,[DEP] AS '填表部門'
                                    ,[KINDS] AS '新品舊品更改'
                                    ,[UNITS] AS '計量單位'
                                    ,[BOXS] AS '入/箱'
                                    ,[BEFORESIZES] AS '烘焙前尺寸(長/寬/厚度)'
                                    ,[AFTERSIZES] AS '烘焙後尺寸(長/寬/厚度)'
                                    ,[BEFOREWEIGHTS] AS '烘焙前重量(克/厚度)'
                                    ,[AFTERWEIGHTS] AS '烘焙後重量(克/厚度)'
                                    ,[MOQS] AS '標準批量(1桶產量)'
                                    ,[MOQMINS] AS '最低製造量'
                                    ,[PROCESS] AS '製作程序說明或附件'                                  
                                    ,[COMMENTS] AS '備註'
                                    FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                    WHERE [MID]='{0}'
                                    ", MID);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);


                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds = new DataSet(); // 這樣就不需要再 Clear()
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds.Tables["ds"];
                    //dataGridView1.AutoResizeColumns();
                    // 指定固定寬度
                    dataGridView3.Columns["品號"].Width = 200;
                    dataGridView3.Columns["品名"].Width = 200;
                   
                }

            }
            catch (Exception EX)
            {

            }
            finally
            {

            }

        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX_NULL_TAB2();
            {
                if (dataGridView3.CurrentRow != null)
                {
                    string mid = dataGridView3.CurrentRow.Cells["MID"].Value.ToString();

                    string MB001 = dataGridView3.CurrentRow.Cells["品號"].Value.ToString();
                    string MB002 = dataGridView3.CurrentRow.Cells["品名"].Value.ToString();
                    string ISCLOSED = dataGridView3.CurrentRow.Cells["結案碼"].Value.ToString();
                    string DEP = dataGridView3.CurrentRow.Cells["填表部門"].Value.ToString();
                    string KINDS = dataGridView3.CurrentRow.Cells["新品舊品更改"].Value.ToString();
                    string UNITS = dataGridView3.CurrentRow.Cells["計量單位"].Value.ToString();
                    string BOXS = dataGridView3.CurrentRow.Cells["入/箱"].Value.ToString();
                    string BEFORESIZES = dataGridView3.CurrentRow.Cells["烘焙前尺寸(長/寬/厚度)"].Value.ToString();
                    string AFTERSIZES = dataGridView3.CurrentRow.Cells["烘焙後尺寸(長/寬/厚度)"].Value.ToString();
                    string BEFOREWEIGHTS = dataGridView3.CurrentRow.Cells["烘焙前重量(克/厚度)"].Value.ToString();
                    string AFTERWEIGHTS = dataGridView3.CurrentRow.Cells["烘焙後重量(克/厚度)"].Value.ToString();
                    string MOQS = dataGridView3.CurrentRow.Cells["標準批量(1桶產量)"].Value.ToString();
                    string MOQMINS = dataGridView3.CurrentRow.Cells["最低製造量"].Value.ToString();
                    string PROCESS = dataGridView3.CurrentRow.Cells["製作程序說明或附件"].Value.ToString();
                    string COMMENTS = dataGridView3.CurrentRow.Cells["備註"].Value.ToString();

                    textBox5.Text = MB001;
                    textBox6.Text = MB002;
                    textBox7.Text = DEP;
                    textBox8.Text = UNITS;
                    textBox9.Text = BOXS;
                    textBox10.Text = BEFORESIZES;
                    textBox11.Text = AFTERWEIGHTS;
                    textBox12.Text = BEFOREWEIGHTS;
                    textBox13.Text = AFTERWEIGHTS;
                    textBox14.Text = MOQS;
                    textBox15.Text = MOQMINS;
                    textBox16.Text = PROCESS;
                    textBox17.Text = COMMENTS;
                    comboBox2.Text = ISCLOSED;
                    comboBox3.Text = KINDS;
                }
            }
        }

        public void UPDATE_TB_PRODUCT_SET_M(
            string MID,
            string MB001,
            string MB002,
            string ISCLOSED,
            string DEP,
            string KINDS,
            string UNITS,
            string BOXS,
            string BEFORESIZES,
            string AFTERSIZES,
            string BEFOREWEIGHTS,
            string AFTERWEIGHTS,
            string MOQS,
            string MOQMINS,
            string PROCESS,
            string COMMENTS
            )
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                    SET 
                                    MB001=@MB001
                                    ,MB002=@MB002
                                    ,ISCLOSED=@ISCLOSED
                                    ,DEP=@DEP
                                    ,KINDS=@KINDS
                                    ,UNITS=@UNITS
                                    ,BOXS=@BOXS
                                    ,BEFORESIZES=@BEFORESIZES
                                    ,AFTERSIZES=@AFTERSIZES
                                    ,BEFOREWEIGHTS=@BEFOREWEIGHTS
                                    ,AFTERWEIGHTS=@AFTERWEIGHTS
                                    ,MOQS=@MOQS
                                    ,MOQMINS=@MOQMINS
                                    ,PROCESS=@PROCESS
                                    ,COMMENTS=@COMMENTS
                                    WHERE MID=@MID
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                //使用 cmd.Parameters.Clear() 清除之前的参数，确保在每次执行时没有冲突
                cmd.Parameters.Clear();
                // 使用參數化查詢，並對每個參數進行賦值
                cmd.Parameters.AddWithValue("@MID", MID);
                cmd.Parameters.AddWithValue("@MB001", MB001);
                cmd.Parameters.AddWithValue("@MB002", MB002);
                cmd.Parameters.AddWithValue("@ISCLOSED", ISCLOSED);
                cmd.Parameters.AddWithValue("@DEP", DEP);
                cmd.Parameters.AddWithValue("@KINDS", KINDS);
                cmd.Parameters.AddWithValue("@UNITS", UNITS);
                cmd.Parameters.AddWithValue("@BOXS", BOXS);
                cmd.Parameters.AddWithValue("@BEFORESIZES", BEFORESIZES);
                cmd.Parameters.AddWithValue("@AFTERSIZES", AFTERSIZES);
                cmd.Parameters.AddWithValue("@BEFOREWEIGHTS", BEFOREWEIGHTS);
                cmd.Parameters.AddWithValue("@AFTERWEIGHTS", AFTERWEIGHTS);
                cmd.Parameters.AddWithValue("@MOQS", MOQS);
                cmd.Parameters.AddWithValue("@MOQMINS", MOQMINS);
                cmd.Parameters.AddWithValue("@PROCESS", PROCESS);
                cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

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

        public void SEARCH_GV4(string ISCLOSED, string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlDataAdapter adapter;
            DataSet ds;

            dataGridView4.DataSource = null;
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

                if (!string.IsNullOrEmpty(ISCLOSED))
                {
                    QUERYS.AppendFormat(@"AND ISCLOSED='{0}'", ISCLOSED);
                }

                if (!string.IsNullOrEmpty(MB001))
                {
                    QUERYS2.AppendFormat(@"AND (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')", MB001);
                }
                else
                {
                    QUERYS2.AppendFormat(@"");
                }

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MID]
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ", QUERYS.ToString(), QUERYS2.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);


                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds = new DataSet(); // 這樣就不需要再 Clear()
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView4.DataSource = ds.Tables["ds"];
                    //dataGridView1.AutoResizeColumns();
                    // 指定固定寬度
                    dataGridView4.Columns["品號"].Width = 200;
                    dataGridView4.Columns["品名"].Width = 200;
                    
                }

            }
            catch (Exception EX)
            {

            }
            finally
            {

            }
        }

        public void SETFASTREPORT(string MID)
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password); 
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1(MID);
            Report report1 = new Report();
            report1.Load(@"REPORT\半成品設定表.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string MID)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(MID))
            {
                SB.AppendFormat(@"
                                SELECT 
                                [TB_PRODUCT_SET_M].[MID]
                                ,[TB_PRODUCT_SET_M].[MB001] AS '品號'
                                ,[TB_PRODUCT_SET_M].[MB002] AS '品名'
                                ,[TB_PRODUCT_SET_M].[ISCLOSED] AS '結案碼'
                                ,[TB_PRODUCT_SET_M].[DEP] AS '填表部門'
                                ,[TB_PRODUCT_SET_M].[KINDS] AS '新品舊品更改 '
                                ,[TB_PRODUCT_SET_M].[UNITS] AS '計量單位'
                                ,[TB_PRODUCT_SET_M].[BOXS] AS '入/箱'
                                ,[TB_PRODUCT_SET_M].[BEFORESIZES] AS '烘焙前尺寸(長/寬/厚度)'
                                ,[TB_PRODUCT_SET_M].[AFTERSIZES] AS '烘焙後尺寸(長/寬/厚度)'
                                ,[TB_PRODUCT_SET_M].[BEFOREWEIGHTS] AS '烘焙前重量(克/厚度)'
                                ,[TB_PRODUCT_SET_M].[AFTERWEIGHTS] AS '烘焙後重量(克/厚度)'
                                ,[TB_PRODUCT_SET_M].[MOQS] AS '標準批量(1桶產量)'
                                ,[TB_PRODUCT_SET_M].[MOQMINS] AS '最低製造量'
                                ,[TB_PRODUCT_SET_M].[PROCESS] AS '製作程序說明或附件'
                                ,[TB_PRODUCT_SET_M].[COMMENTS] AS '備註'
                                ,[TB_PRODUCT_SET_D].[SERNO] AS '序號'
                                ,[TB_PRODUCT_SET_D].[MB001] AS '明細品號'
                                ,[TB_PRODUCT_SET_D].[MB002] AS '明細品名'
                                ,[TB_PRODUCT_SET_D].[AMOUNTS] AS '明細用量'
                                ,[TB_PRODUCT_SET_D].[UNITS] AS '明細單位'
                                ,[TB_PRODUCT_SET_D].[COMMENTS] AS '明細備註'
                                FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                LEFT JOIN [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D] ON [TB_PRODUCT_SET_M].MID=[TB_PRODUCT_SET_D].MID
                                WHERE [TB_PRODUCT_SET_M].MID='{0}'
                                ORDER BY [TB_PRODUCT_SET_M].[MID],[TB_PRODUCT_SET_D].[SERNO]
                                ", MID);
            }
            return SB;

        }

        public void SET_TEXTBOX_NULL()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
        }
        public void SET_TEXTBOX_NULL_TAB2()
        {
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
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(comboBox1.Text.ToString(),textBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                adapter_TB_PRODUCT_SET_M.Update(ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"]); // 更新回資料庫
                adapter_TB_PRODUCT_SET_D.Update(ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"]); // 更新回資料庫

                SET_dataGridView1_MID();
                SEARCH(comboBox1.Text.ToString(), textBox1.Text.Trim());

                MessageBox.Show("資料已儲存成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存失敗：" + ex.Message);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                {
                    string mid = dataGridView1.CurrentRow.Cells["MID"].Value.ToString();
                    string mb001 = dataGridView1.CurrentRow.Cells["品號"].Value.ToString();
                    string mb002 = dataGridView1.CurrentRow.Cells["品名"].Value.ToString();

                    // 顯示確認訊息
                    var result = MessageBox.Show(
                        $"確定要刪除這筆資料嗎？\n 品號 = {mb001}, 品名 = {mb002}",
                        "刪除確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.Yes)
                    {
                        //int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["MID"].Value);

                        // 先刪除明細
                        foreach (DataRow dr in ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"].Select($"MID='{mid}'"))
                        {
                            dr.Delete();
                        }
                        adapter_TB_PRODUCT_SET_D.Update(ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"]);


                        // 從 DataGridView 刪掉 → DataRow 標記為 Deleted
                        dataGridView1.Rows.Remove(dataGridView1.CurrentRow);

                        // 一次把異動同步回資料庫
                        adapter_TB_PRODUCT_SET_M.Update(ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"]);

                        SET_dataGridView1_MID();
                        SEARCH(comboBox1.Text.ToString(), textBox1.Text.Trim());
                        MessageBox.Show("資料已刪除成功！");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.CurrentRow != null)
                {
                    string mid = dataGridView2.CurrentRow.Cells["MID"].Value.ToString();
                    string SERNO = dataGridView2.CurrentRow.Cells["序號"].Value.ToString();
                    string mb001 = dataGridView2.CurrentRow.Cells["品號"].Value.ToString();
                    string mb002 = dataGridView2.CurrentRow.Cells["品名"].Value.ToString();

                    // 顯示確認訊息
                    var result = MessageBox.Show(
                        $"確定要刪除這筆資料嗎？\n 品號 = {mb001}, 品名 = {mb002}",
                        "刪除確認",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.Yes)
                    {
                        //int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["MID"].Value);

                        // 從 DataGridView 刪掉 → DataRow 標記為 Deleted
                        dataGridView2.Rows.Remove(dataGridView2.CurrentRow);

                        // 一次把異動同步回資料庫
                        adapter_TB_PRODUCT_SET_D.Update(ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"]);

                        SET_dataGridView1_MID();
                        SEARCH(comboBox1.Text.ToString(), textBox1.Text.Trim());
                        MessageBox.Show("資料已刪除成功！");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCH_TB_PRODUCT_SET_M(textBox2.Text.Trim());
        }


        private void button6_Click(object sender, EventArgs e)
        {
            string MID = textBox2.Text;
            string MB001 = textBox5.Text;
            string MB002 = textBox6.Text;
            string ISCLOSED = comboBox2.Text.ToString();
            string DEP = textBox7.Text;
            string KINDS = comboBox3.Text.ToString();
            string UNITS = textBox8.Text;
            string BOXS = textBox9.Text;
            string BEFORESIZES = textBox10.Text;
            string AFTERSIZES = textBox11.Text;
            string BEFOREWEIGHTS = textBox12.Text;
            string AFTERWEIGHTS = textBox13.Text;
            string MOQS = textBox14.Text;
            string MOQMINS = textBox15.Text;
            string PROCESS = textBox16.Text;
            string COMMENTS = textBox17.Text;

            UPDATE_TB_PRODUCT_SET_M(
            MID,
            MB001,
            MB002,
            ISCLOSED,
            DEP,
            KINDS,
            UNITS,
            BOXS,
            BEFORESIZES,
            AFTERSIZES,
            BEFOREWEIGHTS,
            AFTERWEIGHTS,
            MOQS,
            MOQMINS,
            PROCESS,
            COMMENTS
            );

            SEARCH_TB_PRODUCT_SET_M(textBox2.Text.Trim());
            MessageBox.Show("完成");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SEARCH_GV4(comboBox4.Text.ToString(), textBox18.Text.Trim());
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                string mid = dataGridView4.CurrentRow.Cells["MID"].Value.ToString();
                SETFASTREPORT(mid);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {

        }

        #endregion


    }
}
