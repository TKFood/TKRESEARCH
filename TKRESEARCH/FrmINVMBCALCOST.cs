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
    public partial class FrmINVMBCALCOST : Form
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

 


        DataSet ds = new DataSet();
      
        int rownum = 0;
        int result;
    

        string ID;


        public FrmINVMBCALCOST()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
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
            Sequel.AppendFormat(@"SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='CALCOSTPRODS'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));
           
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
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
            Sequel.AppendFormat(@"SELECT 
                                 [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKRESEARCH].[dbo].[TBPARA]
                                WHERE [KIND]='CALCOSTPRODS'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARANAME";
            comboBox2.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void SEARCH(string PRODNAMES,string ISCLOESED)
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

                if (!string.IsNullOrEmpty(PRODNAMES))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT  
                                        [PRODNAMES] AS '產品名稱'
                                        ,[SPECS] AS '規格及重量'
                                        ,[COSTRAW] AS '原料成本'
                                        ,[COSTMATERIL] AS '物料成本'
                                        ,[COSTART] AS '人工成本(製造+內外包)'
                                        ,[COSTMANU] AS '製造費用' 
                                        ,[COSTTOTALS] AS '單位成本合計'
                                        ,CONVERT(NVARCHAR,[CREATEDATE],112) AS '建立日期'
                                        ,[ISCLOESED] AS '結案'
                                        ,[ID]

                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                        WHERE [PRODNAMES] LIKE '%{0}%'
                                        AND [ISCLOESED] IN ('{1}')
                                        ORDER BY CONVERT(NVARCHAR,[CREATEDATE],112)
                                        ", PRODNAMES, ISCLOESED);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT  
                                        [PRODNAMES] AS '產品名稱'
                                        ,[SPECS] AS '規格及重量'
                                        ,[COSTRAW] AS '原料成本'
                                        ,[COSTMATERIL] AS '物料成本'
                                        ,[COSTART] AS '人工成本(製造+內外包)'
                                        ,[COSTMANU] AS '製造費用' 
                                        ,[COSTTOTALS] AS '單位成本合計'
                                        ,CONVERT(NVARCHAR,[CREATEDATE],112) AS '建立日期'
                                        ,[ISCLOESED] AS '結案'
                                        ,[ID]

                                        FROM [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                        WHERE [ISCLOESED] IN ('{0}')

                                        ORDER BY CONVERT(NVARCHAR,[CREATEDATE],112)
                                   
                                    ", ISCLOESED);
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
            textBoxID.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;               

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["產品名稱"].Value.ToString();
                    textBox2.Text = row.Cells["規格及重量"].Value.ToString();
                    textBoxID.Text = row.Cells["ID"].Value.ToString();
                   


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBoxID.Text = null;
                }
            }
        }

        public void ADDCALCOSTPRODS(string PRODNAMES,string SPECS)
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
                                    INSERT INTO  [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                    ([PRODNAMES],[SPECS])
                                    VALUES
                                    ('{0}','{1}')

                                    ", PRODNAMES, SPECS);

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


           
        }

        public void UPDATECALCOSTPRODS(string PRODNAMES, string SPECS,string ID,string ISCLOESED)
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
                                   UPDATE [TKRESEARCH].[dbo].[CALCOSTPRODS]
                                    SET [PRODNAMES]='{1}',[SPECS]='{2}',[ISCLOESED]='{3}'
                                    WHERE [ID]='{0}'

                                    ", ID, PRODNAMES, SPECS, ISCLOESED);

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



        }

        public void SETTEXTBOX1()
        {
            textBox3.Text = null;
            textBox4.Text = null;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox999.Text.Trim(),comboBox1.Text.ToString());
        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADDCALCOSTPRODS(textBox3.Text, textBox4.Text);
            SETTEXTBOX1();

            SEARCH(textBox999.Text.Trim(), comboBox1.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UPDATECALCOSTPRODS(textBox1.Text, textBox2.Text,textBoxID.Text,comboBox2.Text);

            SEARCH(textBox999.Text.Trim(), comboBox1.Text.ToString());
        }

        #endregion

     
    }
}
