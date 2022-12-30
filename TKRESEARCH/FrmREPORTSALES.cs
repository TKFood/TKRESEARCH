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
using TKITDLL;

namespace TKRESEARCH
{
    public partial class FrmREPORTSALES : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        public FrmREPORTSALES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATE,string EDATES,string MB001)
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

            SQL1 = SETSQL(SDATE, EDATES, MB001);
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售資料.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATE, string EDATES, string MB001)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 

                            SELECT  年月,TG004 AS '客代',MA002 AS '客戶',MR1MR003 AS '分類1',MR2MR003 AS '分類2',LA001 AS '品號',MB002 AS '品名',MB003 AS '規格',SUM(LA011) AS '銷售數量',SUM(TH037) AS '銷售金額'
                            FROM 
                            (
                            SELECT TG004,SUBSTRING(TG003,1,6) AS '年月',MA002,MR1.MR003 MR1MR003,MR2.MR003 MR2MR003,MA015,TG003,LA001,MB002,MB003,LA011,TH037
                            FROM 
                            (
                            SELECT TG003,TG004,TH001,TH002,TH003,LA001,LA011,TH037
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA
                            WHERE TG001=TH001 AND TG002=TH002
                            AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                            AND TG023='Y'
                            AND (LA001 LIKE '4%' OR LA001 LIKE '5%')
                            AND TG003>='{0}' AND TG003<='{1}'
                            UNION ALL
                            SELECT TB001,TB002,'','','',TB010,TB019,TB031
                            FROM [TK].dbo.POSTB
                            WHERE  (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND TB001>='{0}' AND TB001<='{1}'
                            )  AS TEMP 
                            LEFT JOIN [TK].dbo.COPMA ON MA001=TG004
                            LEFT JOIN [TK].dbo.CMSMR MR1 ON MA017=MR1.MR002 AND MR1.MR001=1 
                            LEFT JOIN [TK].dbo.CMSMR MR2 ON MA076=MR2.MR002 AND MR2.MR001=2 
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA001
                            ) AS TEMP2
                            WHERE  (LA001 LIKE '%{2}%' OR MB002 LIKE '%{2}%')
                            GROUP BY 年月,TG004,MA002,MR1MR003,MR2MR003,LA001,MB002,MB003
                            ", SDATE, EDATES, MB001);

            return SB;

        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker1.Value.ToString("yyyyMMdd"), textBox1.Text);
        }
        #endregion

    }
}
