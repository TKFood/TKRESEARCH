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
    public partial class FrmDEVINV : Form
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

        public FrmDEVINV()
        {
            InitializeComponent();
        }

       

        public void SEARCH()
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
                                    ");

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
                    textBox7.Text = row.Cells["品號"].Value.ToString();
                    textBox8.Text = row.Cells["品名"].Value.ToString();
                    textBox9.Text = row.Cells["單位"].Value.ToString();


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox7.Text = null;
                    textBox8.Text = null;
                    textBox9.Text = null;

                }
            }
        }

        public void INSERTINVLA(string INOUT
                                , string MB001
                                , string NAME
                                , string UNIT
                                , string NUMS
                                , string LOT
                                , string CMMENTS
                                , string USERNAME
                                )
        {
            DateTime dt1 = DateTime.Now;
            string ID = GETMAXID(dt1.ToString("yyyyMMdd"), dt1);

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
                                    INSERT INTO  [TKRESEARCH].[dbo].[INVLA]
                                    (
                                    [ID]
                                    ,[INOUT]
                                    ,[MB001]
                                    ,[NAME]
                                    ,[UNIT]
                                    ,[NUMS]
                                    ,[LOT]
                                    ,[CMMENTS]
                                    ,[USERNAME]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,{1}
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,{5}
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    )
                                    ", ID, INOUT, MB001, NAME, UNIT, NUMS, LOT, CMMENTS, USERNAME);

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

           
            SEARCH();

        }

        public string GETMAXID(string DATES,DateTime dt1)
        {
            string ID;
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();   
                            
                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(MAX(ID),'00000000000') AS ID
                                    FROM [TKRESEARCH].[dbo].[INVLA]
                                    WHERE ID LIKE '{0}%'
                                    ", DATES);

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        ID = SETID(ds4.Tables["TEMPds4"].Rows[0]["ID"].ToString(), dt1);
                        return ID;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
           

        }
        public string SETID(string ID,DateTime dt1)
        {           

            if (ID.Equals("00000000000"))
            {
                return dt1.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(ID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt1.ToString("yyyyMMdd") + temp.ToString();
            }

           
        }

        public void SETNULL()
        {
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
        }

        #region FUNCTION


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string USERNAME=shareArea.UserName;

            INSERTINVLA("1"
                        , textBox1.Text.Trim()
                        , textBox2.Text.Trim()
                        , textBox3.Text.Trim()
                        , textBox4.Text.Trim()
                        , textBox5.Text.Trim()
                        , textBox6.Text.Trim()
                        , USERNAME
                        );
            SETNULL();

            //MessageBox.Show(USERNAME);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string USERNAME = shareArea.UserName;

            INSERTINVLA("-1"
                        , textBox7.Text.Trim()
                        , textBox8.Text.Trim()
                        , textBox9.Text.Trim()
                        , textBox10.Text.Trim()
                        , textBox11.Text.Trim()
                        , textBox12.Text.Trim()
                        , USERNAME
                        );
            SETNULL();
        }

        #endregion


    }
}
