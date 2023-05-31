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
    public partial class FrmREPORTCOST : Form
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

        public FrmREPORTCOST()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT2(string SDATE, string MB001, string MB003, string MB002)
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

            SQL1 = SETSQL2(SDATE, MB001, MB003, MB002);
            Report report1 = new Report();
            report1.Load(@"REPORT\平均成本.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string YM, string MB001, string MB003, string MB002)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY = new StringBuilder();

            //查詢條件
            if (!string.IsNullOrEmpty(MB001))
            {
                SQUERY.AppendFormat(@" AND TA001 LIKE '%{0}%' ", MB001);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }

            if (!string.IsNullOrEmpty(MB002))
            {
                SQUERY.AppendFormat(@" AND MB002 LIKE '%{0}%' ", MB002);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(MB003))
            {
                SQUERY.AppendFormat(@" AND MB003 LIKE '%{0}%' ", MB003);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }

            if (!string.IsNullOrEmpty(YM) && !string.IsNullOrEmpty(SQUERY.ToString()))
            {
                SB.AppendFormat(@"
                                    SELECT *
                                    FROM 
                                    (
                                    SELECT TA002 AS '年月',TA001 AS '品號',MB002 AS '品名',MB003 AS '規格',生產入庫數,ME005 在製約量_材料,本階人工成本,本階製造費用,ME007 材料成本,ME008 人工成本,ME009 製造費用,ME010 加工費用
                                    ,CONVERT(DECIMAL(16,2),((ME007+ME008+ME009+ME010)/(生產入庫數+ME005))) 單位成本, CONVERT(DECIMAL(16,2),((ME007)/(生產入庫數+ME005))) 單位材料成本, CONVERT(DECIMAL(16,2),((ME008)/(生產入庫數+ME005))) 單位人工成本,CONVERT(DECIMAL(16,2),((ME009)/(生產入庫數+ME005))) 單位製造成本,CONVERT(DECIMAL(16,2),((ME010)/(生產入庫數+ME005))) 單位加工成本
                                    ,MB068
                                    ,(CASE WHEN MB068 IN ('09') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均包裝人工成本
                                    ,(CASE WHEN MB068 IN ('09') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均包裝製造費用
                                    ,(CASE WHEN MB068 IN ('03') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均小線人工成本
                                    ,(CASE WHEN MB068 IN ('03') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均小線製造費用
                                    ,(CASE WHEN MB068 IN ('02') THEN 本階人工成本/(生產入庫數+ME005) ELSE 0 END ) 平均大線人工成本
                                    ,(CASE WHEN MB068 IN ('02') THEN 本階製造費用/(生產入庫數+ME005) ELSE 0 END ) 平均大線製造費用
                                    ,MB047
                                    FROM 
                                    (
                                    SELECT TA002,TA001,SUM(TA012) '生產入庫數',SUM(TA016-TA019) AS '本階人工成本',SUM(TA017-TA020) AS '本階製造費用'
                                    FROM [TK].dbo.CSTTA
                                    WHERE TA002 LIKE '{0}%'
                                    GROUP BY TA002,TA001
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                                    WHERE 1=1
                                    {1}

                                    AND (生產入庫數+ME005)>0
                                   
                                    ) AS TEMP2
                                    ORDER BY  品號,年月

 

                                    ", YM, SQUERY.ToString());
            }


            return SB;

        }

        public void SETFASTREPORT3(string SDATE, string MB001, string MB002, string MB003)
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

            SQL1 = SETSQL3(SDATE, MB001, MB002, MB003);
            Report report1 = new Report();
            report1.Load(@"REPORT\品號平均年度單位成本.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string YM, string MB001, string MB002, string MB003)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(MB001))
            {
                //SQUERY.AppendFormat(@"   AND 成品品號 LIKE '%{0}%'", MB001);
                SQUERY.AppendFormat(@"
                                    AND 成品品號 IN
                                    (
                                    SELECT TA001 
                                    FROM 
                                    (
                                    SELECT TA002,TA001,SUM(TA012) '生產入庫數'
                                    FROM [TK].dbo.CSTTA
                                    WHERE TA002 LIKE '{0}%'
                                    GROUP BY TA002,TA001
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                                    WHERE 1=1
                                    AND MB001 LIKE '%{1}%'

                                    AND (生產入庫數+ME005)>0
                                    GROUP BY TA001
                                    )
                                    ", YM, MB001);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }

            if (!string.IsNullOrEmpty(MB002))
            {
                //SQUERY.AppendFormat(@"  AND 成品品名 LIKE '%{0}%'", MB002);

                SQUERY.AppendFormat(@"
                                    AND 成品品號 IN
                                    (
                                    SELECT TA001 
                                    FROM 
                                    (
                                    SELECT TA002,TA001,SUM(TA012) '生產入庫數'
                                    FROM [TK].dbo.CSTTA
                                    WHERE TA002 LIKE '{0}%'
                                    GROUP BY TA002,TA001
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                                    WHERE 1=1
                                    AND MB002 LIKE '%{1}%'

                                    AND (生產入庫數+ME005)>0
                                    GROUP BY TA001
                                    )
                                    ", YM, MB002);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(MB003))
            {
                //SQUERY.AppendFormat(@"  AND 成品品名 LIKE '%{0}%'", MB002);

                SQUERY.AppendFormat(@"
                                    AND 成品品號 IN
                                    (
                                    SELECT TA001 
                                    FROM 
                                    (
                                    SELECT TA002,TA001,SUM(TA012) '生產入庫數'
                                    FROM [TK].dbo.CSTTA
                                    WHERE TA002 LIKE '{0}%'
                                    GROUP BY TA002,TA001
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                                    WHERE 1=1
                                    AND MB003 LIKE '%{1}%'

                                    AND (生產入庫數+ME005)>0
                                    GROUP BY TA001
                                    )
                                    ", YM, MB003);
            }
            else
            {
                SQUERY.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(YM) && !string.IsNullOrEmpty(SQUERY.ToString()))
            {
                SB.AppendFormat(@"


                                    SELECT *
                                    ,CONVERT(NVARCHAR,CONVERT(DECIMAL(16,4),(CASE WHEN 總成品平均成本>0 THEN 分攤成本/總成品平均成本 ELSE 0 END))*100)+'%' AS '各百分比' 
                                    ,CONVERT(DECIMAL(16,2),分攤成本) AS 分攤成本
                                    FROM 
                                    (
                                    SELECT '{0}'AS '年度',MC001 AS '成品品號',MB1.MB002  AS '成品品名' ,MC004,MD003  AS '使用品號',MB2.MB002  AS '使用品名',MD006,MD007
                                    ,總成品平均成本
                                    ,材料平均成本
                                    ,人工平均成本
                                    ,製造平均成本
                                    ,加工平均成本
                                    ,各採購單位成本
                                    ,總採購單位成本
                                    ,總半成品重
                                    ,(CASE WHEN 總成品平均成本>0 THEN (CASE WHEN (MB2.MB001 LIKE '3%' OR MB2.MB001 LIKE '4%')THEN ((材料平均成本-總採購單位成本)*MD006/MD007/總半成品重) ELSE 各採購單位成本*MD006/MD007/MC004 END) ELSE 0 END) AS '分攤成本' 
                                    ,(CASE WHEN MD003 LIKE '1%' THEN '1原料'  WHEN MD003 LIKE '2%' THEN '2物料' WHEN (MD003 LIKE '3%' OR MD003 LIKE '4%') THEN '3半成品'END ) AS '分類'
                                    FROM
                                    (
                                    SELECT MC001,MC004,MD003,MD006,MD007
                                    ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                                    ,ISNULL((SELECT AVG((ME007)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '材料平均成本'
                                    ,ISNULL((SELECT AVG((ME008)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '人工平均成本'
                                    ,ISNULL((SELECT AVG((ME009)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '製造平均成本'
                                    ,ISNULL((SELECT AVG((ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MD001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '加工平均成本'
                                    ,(CASE WHEN ( MB2.MB001 LIKE '1%' OR MB2.MB001 LIKE '2%') AND MB2.MB064>0 AND MB2.MB065 >0 THEN MB2.MB065/MB2.MB064*MD006/MD007/MC004 ELSE MB2.MB050*MD006/MD007/MC004 END ) AS '各採購單位成本'
                                    ,(SELECT SUM (CASE WHEN  ( MB001 LIKE '1%' OR MB001 LIKE '2%') AND MB064>0 AND MB065 >0 THEN MB065/MB064*MD006/MD007/MC004 ELSE MB050*MD006/MD007/MC004 END) FROM [TK].dbo.BOMMC MC, [TK].dbo.BOMMD MD ,[TK].dbo.INVMB MB WHERE  MC.MC001=MD.MD001 AND MD.MD003=MB.MB001 AND MC.MC001=BOMMC.MC001)   AS '總採購單位成本'
                                    ,ISNULL((SELECT SUM (MD006/MD007) FROM [TK].dbo.BOMMC MC, [TK].dbo.BOMMD MD ,[TK].dbo.INVMB MB WHERE  MC.MC001=MD.MD001 AND MD.MD003=MB.MB001 AND MC.MC001=BOMMC.MC001 AND (MB.MB001 LIKE '3%' OR MB.MB001 LIKE '4%')),0)  AS '總半成品重'
                                    FROM [TK].dbo.BOMMC
                                    LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=BOMMC.MC001
                                    , [TK].dbo.BOMMD
                                    LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=BOMMD.MD003
                                    WHERE MC001=MD001
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=TEMP.MC001
                                    LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=TEMP.MD003
                                    UNION ALL
                                    SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                                    ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,ISNULL((SELECT AVG((ME008)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '成本'
                                    ,'4人工' AS '分類'
                                    FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                                    WHERE  MC001=MB001
                                    AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 
                                    UNION ALL
                                    SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                                    ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,ISNULL((SELECT AVG((ME009)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '成本'
                                    ,'5製造' AS '分類'
                                    FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                                    WHERE  MC001=MB001
                                    AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 
                                    UNION ALL
                                    SELECT '{0}',MC001 AS '成品品號',MB002  AS '成品品名',0 ,''  AS '使用品號','' AS '使用品名',0,0
                                    ,ISNULL((SELECT AVG((ME007+ME008+ME009+ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '總成品平均成本'
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,0
                                    ,ISNULL((SELECT AVG((ME010)/(ME003+ME005+ME004)) FROM [TK].dbo.CSTME WHERE  ME001=MC001 AND (ME003+ME005+ME004)>0 AND (ME007+ME008+ME009+ME010)>0 AND ME002 LIKE '{0}%'),0) AS '成本'
                                    ,'6加工' AS '分類'
                                    FROM [TK].dbo.BOMMC,[TK].dbo.INVMB
                                    WHERE  MC001=MB001
                                    AND (MC001 LIKE '3%' OR MC001 LIKE '4%' OR MC001 LIKE '5%') 

                                  
                                    ) AS TEMP2
                                    WHERE 1=1
                                    {1}
                                    ORDER BY 成品品號,分類,使用品號

                                    ", YM, SQUERY.ToString());
            }

            return SB;

        }

        public void SETFASTREPORT4(string SDATE, string MB001, string MB002, string MB003)
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

            SQL1 = SETSQL4(SDATE, MB001, MB002, MB003);
            Report report1 = new Report();
            report1.Load(@"REPORT\BOM表成本.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL4(string YM, string MB001, string MB002, string MB003)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SQUERY1 = new StringBuilder();
            StringBuilder SQUERY2 = new StringBuilder();
            StringBuilder SQUERY3 = new StringBuilder();

            if (!string.IsNullOrEmpty(MB001))
            {
                //SQUERY.AppendFormat(@"   AND 成品品號 LIKE '%{0}%'", MB001);
                SQUERY1.AppendFormat(@"
                                    AND MC001 LIKE '%{0}%'  
                                    ",  MB001);
            }
            else
            {
                SQUERY1.AppendFormat(@"");
            }

            if (!string.IsNullOrEmpty(MB002))
            {
                //SQUERY.AppendFormat(@"  AND 成品品名 LIKE '%{0}%'", MB002);

                SQUERY2.AppendFormat(@"
                                    AND  MB1.MB002 LIKE '%{0}%'
                                    ", MB002);
            }
            else
            {
                SQUERY2.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(MB003))
            {
                //SQUERY.AppendFormat(@"  AND 成品品名 LIKE '%{0}%'", MB002);

                SQUERY3.AppendFormat(@"
                                    AND  MB1.MB003 LIKE '%{0}%'
                                    ",  MB003);
            }
            else
            {
                SQUERY3.AppendFormat(@"");
            }
            if (!string.IsNullOrEmpty(YM))
            {
                SB.AppendFormat(@"
                                SELECT 
                                MC001 AS '成品品號',MB1.MB002 AS '成品品名',MD003 AS '組件品號',MB2.MB002 AS '組件品名',MB2.MB004 AS '組件單位',CONVERT(decimal(16,4),MB2.MB050) AS '最近進價',MB2.MB102  AS '進價是否含稅',MC004 AS '標準批量',MD006 AS '組成用量',MD007 AS '底數',MD008 AS '損秏率'
                                ,(SELECT TOP 1 '最近進貨日:'+TG003+' 廠商:'+TG005+' '+MA002 FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.PURMA WHERE TG001=TH001 AND TG002=TH002 AND TG005=MA001 AND TH004=MD003 ORDER BY TG003 DESC) AS 'MA002'
                                ,(CASE WHEN MD003 LIKE '1%' OR MD003 LIKE '2%' THEN(CONVERT(decimal(16,4),MB2.MB050*MD006/MD007*(1+MD008)/MC004)) ELSE 0 END) AS '分攤單位進貨成本'
                                ,CONVERT(decimal(16,2),(CASE WHEN MD003 NOT LIKE '1%' THEN 
                                (CASE WHEN MD003 NOT LIKE '2%' THEN 
                                ((SELECT AVG(LB010) FROM [TK].dbo.INVLB WHERE LB001=MD003 AND LB002 LIKE '{0}%' GROUP BY LB001)*MD006/MD007*(1+MD008)/MC004) 
                                ELSE 0 END)
                                ELSE 0 END)) AS '非採購單位成本'

                                ,CONVERT(decimal(16,2),(SELECT SUM(MB050*MD006/MD007*(1+MD008)/MC004) FROM [TK].dbo.BOMMC MC,[TK].dbo.BOMMD MD,[TK].dbo.INVMB MB WHERE MC.MC001=MD.MD001 AND MB.MB001=MD.MD003 AND MD.MD001=BOMMC.MC001 ))  AS '成品單位進貨成本'
                                ,CONVERT(decimal(16,2),(SELECT AVG(LB010) LB010
                                FROM [TK].dbo.INVLB
                                WHERE LB001=MC001
                                AND LB002 LIKE '{0}%'
                                GROUP BY LB001)) AS '單位成本-材料'
                                FROM [TK].dbo.BOMMC
                                LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=MC001
                                ,[TK].dbo.BOMMD
                                LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=MD003
                                WHERE MC001=MD001
                                {1}
                                {2}
                                {3}
                                ORDER BY MC001,MD003 


                                    ", YM, SQUERY1.ToString(), SQUERY2.ToString(), SQUERY3.ToString());
            }

            return SB;

        }



        #endregion

        #region BUTTON

        private void button61_Click(object sender, EventArgs e)
        {
            tabControl2.SelectedTab = tabControl2.TabPages["tabPage12"];
            SETFASTREPORT2(dateTimePicker1.Value.ToString("yyyy"), textBox123.Text, textBox124.Text, textBox125.Text);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            tabControl2.SelectedTab = tabControl2.TabPages["tabPage13"];
            SETFASTREPORT3(dateTimePicker1.Value.ToString("yyyy"), textBox123.Text, textBox125.Text, textBox124.Text);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker2.Value.ToString("yyyy"), textBox1.Text, textBox2.Text, textBox3.Text);
        }
        #endregion


    }
}
