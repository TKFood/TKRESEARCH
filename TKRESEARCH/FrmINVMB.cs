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
    public partial class FrmINVMB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();


        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;
        int rowindex;
        int ROWSINDEX;
        int COLUMNSINDEX;

        string ID;

        public FrmINVMB()
        {
            InitializeComponent();
        }


        #region FUNCTION

        public void SEARCH(string MB001)
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



                sbSql.Clear();

                if(!string.IsNullOrEmpty(MB001))
                {
                    sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MB001] AS '品號'
                                    ,[NAME] AS '品名'
                                    ,[UNIT] AS '單位'
                                    ,[SUPPLIER] AS '供應商'
                                    ,[ORIGIN] AS '產地'
                                    ,[UNITWEIGHT] AS '單位重量'
                                    ,[SAVELIFE] AS '保存期限'
                                    ,[SAVESONDITIONS] AS '保存條件'
                                    ,[METARIAL] AS '材質'
                                    ,(SELECT ISNULL(SUM([INOUT]*[NUMS]),0) FROM [TKRESEARCH].[dbo].[INVLA] WHERE [INVLA].MB001=[INVMB].MB001) AS '總數量'
                                    FROM [TKRESEARCH].[dbo].[INVMB]
                                    WHERE ( MB001 LIKE '%{0}%' OR  NAME LIKE '%{0}%')
                                    ORDER BY MB001
                                    ", MB001);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                    SELECT 
                                    [MB001] AS '品號'
                                    ,[NAME] AS '品名'
                                    ,[UNIT] AS '單位'
                                    ,[SUPPLIER] AS '供應商'
                                    ,[ORIGIN] AS '產地'
                                    ,[UNITWEIGHT] AS '單位重量'
                                    ,[SAVELIFE] AS '保存期限'
                                    ,[SAVESONDITIONS] AS '保存條件'
                                    ,[METARIAL] AS '材質'
                                    ,(SELECT ISNULL(SUM([INOUT]*[NUMS]),0) FROM [TKRESEARCH].[dbo].[INVLA] WHERE [INVLA].MB001=[INVMB].MB001) AS '總數量'
                                    FROM [TKRESEARCH].[dbo].[INVMB]
                                    ORDER BY MB001
                                    ");
                }
                

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

                        dataGridView1.AutoResizeColumns();

                        if (ROWSINDEX > 0 || COLUMNSINDEX > 0)
                        {
                            dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                            DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];
                            ID = row.Cells["ID"].Value.ToString();

                          



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
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox17.Text = null;
            textBox99.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (dataGridView1.CurrentCell.RowIndex > 0 || dataGridView1.CurrentCell.ColumnIndex > 0)
                {
                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;
                }

          
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["品號"].Value.ToString();
                    textBox2.Text = row.Cells["品名"].Value.ToString();
                    textBox3.Text = row.Cells["單位"].Value.ToString();
                    textBox7.Text = row.Cells["供應商"].Value.ToString();
                    textBox8.Text = row.Cells["產地"].Value.ToString();
                    textBox9.Text = row.Cells["單位重量"].Value.ToString();
                    textBox10.Text = row.Cells["保存期限"].Value.ToString();
                    textBox11.Text = row.Cells["保存條件"].Value.ToString();
                    textBox17.Text = row.Cells["材質"].Value.ToString();
                    textBox99.Text = row.Cells["總數量"].Value.ToString();


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox17.Text = null;
                    textBox99.Text = null;
                }
            }

           
        }

        public void DELINVMB(string MB001)
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    DELETE [TKRESEARCH].[dbo].[INVMB]
                                    WHERE [MB001]='{0}'
                                    ", MB001);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }


            SEARCH(textBox999.Text.Trim());
        }

        public void UPDATEINVMB(string MB001, string NAME, string UNIT, string SUPPLIER, string ORIGIN, string UNITWEIGHT, string SAVELIFE, string SAVESONDITIONS,string METARIAL)
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    UPDATE [TKRESEARCH].[dbo].[INVMB]
                                    SET [NAME]='{1}',[UNIT]='{2}',[SUPPLIER]='{3}',[ORIGIN]='{4}',[UNITWEIGHT]='{5}',[SAVELIFE]='{6}',[SAVESONDITIONS]='{7}',[METARIAL]='{8}'
                                    WHERE [MB001]='{0}'
                                    ", MB001, NAME, UNIT, SUPPLIER, ORIGIN, UNITWEIGHT, SAVELIFE, SAVESONDITIONS, METARIAL);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }


            SEARCH(textBox999.Text.Trim());
        }

        public void INSERTINVMB(string MB001, string NAME, string UNIT, string SUPPLIER, string ORIGIN, string UNITWEIGHT, string SAVELIFE, string SAVESONDITIONS,string METARIAL)
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                   INSERT  [TKRESEARCH].[dbo].[INVMB]
                                    ([MB001],[NAME],[UNIT],[SUPPLIER],[ORIGIN],[UNITWEIGHT],[SAVELIFE],[SAVESONDITIONS],[METARIAL])
                                    VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')
                                    ", MB001, NAME, UNIT, SUPPLIER, ORIGIN, UNITWEIGHT, SAVELIFE, SAVESONDITIONS, METARIAL);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }


            SEARCH(textBox999.Text.Trim());
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\研發品號.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));

            //report1.Preview = previewControl1;
            //report1.Show();

            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.OoXML.Excel2007Export REPORTExcelxport = new FastReport.Export.OoXML.Excel2007Export();
            // show the export options dialog and do the export

            //桌面路徑 
            string DESKTOPNAME = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\";
            //匯出檔名
            string FLESEXPORTNAME = "研發品號" + DateTime.Now.ToString("yyyyMMddHHss") + ".xlsx";
            //匯出到桌面
            report1.Export(REPORTExcelxport, DESKTOPNAME + FLESEXPORTNAME);

            //C#開啟Excel文件，要裝excel
            System.Diagnostics.Process.Start(DESKTOPNAME + FLESEXPORTNAME);

            //if (REPORTExcelxport.ShowDialog())
            //{
            //    report1.Export(REPORTExcelxport, desktop + "result.xlsx");
            //}



        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 
                            [MB001] AS '品號'
                            ,[NAME] AS '品名'
                            ,[UNIT] AS '單位'
                            ,[SUPPLIER] AS '供應商'
                            ,[ORIGIN] AS '產地'
                            ,[UNITWEIGHT] AS '單位重量'
                            ,[SAVELIFE] AS '保存期限'
                            ,[SAVESONDITIONS] AS '保存條件'
                            ,[METARIAL] AS '材質'
                            ,(SELECT ISNULL(SUM([INOUT]*[NUMS]),0) FROM [TKRESEARCH].[dbo].[INVLA] WHERE [INVLA].MB001=[INVMB].MB001) AS '總數量'
                            FROM [TKRESEARCH].[dbo].[INVMB]
                            ORDER BY MB001

                            ");


            return SB;

        }


        public string SEARCHINVLA(string MB001)
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



                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT 品號,品名,單位,SUM(數量) 數量
                                    FROM (
                                    SELECT 品號,品名,單位,批號
                                    ,(SELECT ISNULL(SUM([INOUT]*[NUMS]),0) FROM [TKRESEARCH].[dbo].[INVLA] WHERE [INVLA].MB001=品號 AND [INVLA].[LOT]=批號 )  AS '數量'
                                    FROM(
                                    SELECT 
                                    [INVMB].[MB001] AS '品號'
                                    ,[INVMB].[NAME] AS '品名'
                                    ,[INVMB].[UNIT] AS '單位'
                                    ,ISNULL([INVLA].[LOT],'') AS '批號'
                                    FROM [TKRESEARCH].[dbo].[INVMB]
                                    LEFT JOIN [TKRESEARCH].[dbo].[INVLA] ON [INVLA].MB001=[INVMB].MB001
                                    GROUP BY [INVMB].[MB001],[INVMB].[NAME],[INVMB].[UNIT],ISNULL([INVLA].[LOT],'')
                                    ) AS TEMP
                                    ) AS TEMP 
                                    WHERE 1=1
                                    AND 品號='{0}'
                                    GROUP BY 品號,品名,單位
                                    ORDER BY 品號
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"].Rows[0]["數量"].ToString();
                }
                else
                {
                    return "0";
                }

            }
            catch
            {
                return "0";
            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox999.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UPDATEINVMB(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), textBox7.Text.Trim(), textBox8.Text.Trim(), textBox9.Text.Trim(), textBox10.Text.Trim(), textBox11.Text.Trim(), textBox17.Text.Trim());
        }

        private void button3_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELINVMB(textBox1.Text.Trim());

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            INSERTINVMB(textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox12.Text.Trim(), textBox13.Text.Trim(), textBox14.Text.Trim(), textBox15.Text.Trim(), textBox16.Text.Trim(), textBox18.Text.Trim());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion


    }
}
