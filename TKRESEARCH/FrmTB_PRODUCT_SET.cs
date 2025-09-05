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

        public FrmTB_PRODUCT_SET()
        {
            InitializeComponent();

            DATAGRIDSET();
        }

        #region FUNCTION
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
        public void SEARCH()
        {
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


                sbSql.Clear();              

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MID]
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_M]
                                    ");

                adapter_TB_PRODUCT_SET_M = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter_TB_PRODUCT_SET_M);
                sqlConn.Open();
                ds_TB_PRODUCT_SET_M = new DataSet(); // 這樣就不需要再 Clear()
                adapter_TB_PRODUCT_SET_M.Fill(ds_TB_PRODUCT_SET_M, "ds_TB_PRODUCT_SET_M");
                sqlConn.Close();

                if (ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"];
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

                StringBuilder QUERYS = new StringBuilder();
                StringBuilder QUERYS2 = new StringBuilder();
                StringBuilder QUERYS3 = new StringBuilder();


                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MID]
                                    ,[SERNO] AS '序號'
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    FROM [TKRESEARCH].[dbo].[TB_PRODUCT_SET_D]
                                    WHERE MID='{0}'
                                    ", MID);

                adapter_TB_PRODUCT_SET_D = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter_TB_PRODUCT_SET_D);
                sqlConn.Open();
                ds_TB_PRODUCT_SET_D = new DataSet(); // 這樣就不需要再 Clear()
                adapter_TB_PRODUCT_SET_D.Fill(ds_TB_PRODUCT_SET_D, "ds_TB_PRODUCT_SET_D");
                sqlConn.Close();

                if (ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"];
                    dataGridView2.AutoResizeColumns();
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
            if (dataGridView1.CurrentRow != null)
            {
                string mid = dataGridView1.CurrentRow.Cells["MID"].Value.ToString();
                SEARCH_TB_PRODUCT_SET_D(mid);
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                adapter_TB_PRODUCT_SET_M.Update(ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"]); // 更新回資料庫
                adapter_TB_PRODUCT_SET_D.Update(ds_TB_PRODUCT_SET_D.Tables["ds_TB_PRODUCT_SET_D"]); // 更新回資料庫

                SEARCH();
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

                        // 從 DataGridView 刪掉 → DataRow 標記為 Deleted
                        dataGridView1.Rows.Remove(dataGridView1.CurrentRow);

                        // 一次把異動同步回資料庫
                        adapter_TB_PRODUCT_SET_M.Update(ds_TB_PRODUCT_SET_M.Tables["ds_TB_PRODUCT_SET_M"]);

                        SEARCH();
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
