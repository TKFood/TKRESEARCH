﻿using System;
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
using System.Net.Mail;
using TKITDLL;
using System.Globalization;
using System.Collections;

namespace TKRESEARCH
{
    public partial class FrmTB_DEV_NEWLISTS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string talbename = null;
        int rownum = 0;
        int result;
   


        public FrmTB_DEV_NEWLISTS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string yyyyMM, string NAMES)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            string YY = yyyyMM.Substring(2, 2);
            string MM = yyyyMM.Substring(4, 2);
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

                StringBuilder SQLquery1 = new StringBuilder();
                StringBuilder SQLquery2 = new StringBuilder();

                if (!string.IsNullOrEmpty(NAMES))
                {
                    SQLquery1.AppendFormat(@" AND [NAMES] LIKE '%{0}%' ", NAMES);
                }
                else
                {
                    SQLquery1.AppendFormat(@" ");

                    if(!string.IsNullOrEmpty(yyyyMM))
                    {
                        SQLquery2.AppendFormat(@" AND [NO] LIKE '%{0}%'", YY+"-"+ MM);
                    }
                }
               

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [NO] AS '編號'
                                    ,[NAMES] AS '商品'
                                    ,[SPECS] AS '規格'
                                    ,[COMMENTS] AS '需求'
                                    ,[INGREDIENTS] AS '成份'
                                    ,[COSTS] AS '成本'
                                    ,[MOQS] AS 'MOQ'
                                    ,[MANUPRODS] AS '一天產能量'
                                    ,CONVERT(NVARCHAR,[CARESTEDATES],112) AS '建立日期'
                                    ,[ID]

                                    FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ORDER BY [NO]
                                    ", SQLquery1.ToString(), SQLquery2.ToString());

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


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMM"),textBox1.Text.Trim());
        }

        #endregion
}
}
