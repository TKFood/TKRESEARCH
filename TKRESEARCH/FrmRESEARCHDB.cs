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
    public partial class FrmRESEARCHDB : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        DataSet ds1 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;
        int rowindex;
        int ROWSINDEX;
        int COLUMNSINDEX;

        string ID;
        byte[] BYTES1 = null;
        string CONTENTTYPES1 = null;
        string DOCNAMES1 = null;
        byte[] BYTES2 = null;
        string CONTENTTYPES2 = null;
        string DOCNAMES2 = null;

        byte[] BYTES31 = null;
        string CONTENTTYPES31 = null;
        string DOCNAMES31 = null;
        byte[] BYTES32 = null;
        string CONTENTTYPES32 = null;
        string DOCNAMES32 = null;
        byte[] BYTES33 = null;
        string CONTENTTYPES33 = null;
        string DOCNAMES33 = null;
        byte[] BYTES34 = null;
        string CONTENTTYPES34 = null;
        string DOCNAMES34 = null;
        byte[] BYTES35 = null;
        string CONTENTTYPES35 = null;
        string DOCNAMES35 = null;
        byte[] BYTES36 = null;
        string CONTENTTYPES36 = null;
        string DOCNAMES36 = null;

        byte[] BYTES41= null;
        string CONTENTTYPES41 = null;
        string DOCNAMES41 = null;
        byte[] BYTES42 = null;
        string CONTENTTYPES42 = null;
        string DOCNAMES42 = null;

        byte[] BYTES51 = null;
        string CONTENTTYPES51 = null;
        string DOCNAMES51 = null;
        byte[] BYTES52 = null;
        string CONTENTTYPES52 = null;
        string DOCNAMES52 = null;
        byte[] BYTES53 = null;
        string CONTENTTYPES53 = null;
        string DOCNAMES53 = null;

        byte[] BYTES61 = null;
        string CONTENTTYPES61 = null;
        string DOCNAMES61 = null;
        byte[] BYTES62 = null;
        string CONTENTTYPES62 = null;
        string DOCNAMES62 = null;

        byte[] BYTES641 = null;
        string CONTENTTYPES641 = null;
        string DOCNAMES641 = null;
        byte[] BYTES642 = null;
        string CONTENTTYPES642 = null;
        string DOCNAMES642 = null;


        byte[] BYTES71 = null;
        string CONTENTTYPES71 = null;
        string DOCNAMES71 = null;
        byte[] BYTES72 = null;
        string CONTENTTYPES72 = null;
        string DOCNAMES72 = null;
        byte[] BYTES73 = null;
        string CONTENTTYPES73 = null;
        string DOCNAMES73 = null;

        byte[] BYTES81 = null;
        string CONTENTTYPES81 = null;
        string DOCNAMES81 = null;
        byte[] BYTES82 = null;
        string CONTENTTYPES82 = null;
        string DOCNAMES82 = null;
        byte[] BYTES83 = null;
        string CONTENTTYPES83 = null;
        string DOCNAMES83 = null;

        byte[] BYTES91 = null;
        string CONTENTTYPES91 = null;
        string DOCNAMES91 = null;

        long FILESIZE=0;

        public FrmRESEARCHDB()
        {
            InitializeComponent();

            FILESIZE = FIND_FILESIZE();

            SEARCH(textBox1A.Text.Trim());
            SETdataGridView1();
            SEARCH2(textBox2A.Text.Trim());
            SETdataGridView2();
            SEARCH3(textBox3A.Text.Trim());
            SETdataGridView31();
            SETdataGridView32();
            SETdataGridView33();
            SETdataGridView34();
            SETdataGridView35();
            SETdataGridView36();
            SEARCH4(textBox4A.Text.Trim());
            SETdataGridView41();
            SETdataGridView42();
            SEARCH5(textBox5A.Text);
            SETdataGridView51();
            SETdataGridView52();
            SETdataGridView53();
            SEARCH6(textBox6A.Text, comboBox12.Text);
            SETdataGridView61();
            SETdataGridView62();
            SEARCH7(textBox7A.Text);
            SETdataGridView71();
            SETdataGridView72();
            SETdataGridView73();
            SEARCH8(textBox8A.Text);
            SETdataGridView81();
            SETdataGridView82();
            SETdataGridView83();
            SEARCH9(textBox9A.Text);
            SETdataGridView91();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();
            comboBox7load();
            comboBox8load();
            comboBox9load();
            comboBox10load();
            comboBox11load();
            comboBox12load();

        }

        #region FUNCTION
        public long FIND_FILESIZE()
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

                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='上傳檔案最大'  ORDER BY ID ");
                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("PARAID", typeof(string));
                dt.Columns.Add("PARANAME", typeof(string));
                da.Fill(dt);
                sqlConn.Close();

                if(dt.Rows.Count>0)
                {
                    return Convert.ToInt64(dt.Rows[0]["PARAID"].ToString());
                }
                else
                {
                    return 0;
                }
            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
           
        }
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='原料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='原料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARAID";
            comboBox2.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void comboBox3load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='物料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "PARAID";
            comboBox3.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void comboBox4load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='物料'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "PARAID";
            comboBox4.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox5load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='成品'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "PARAID";
            comboBox5.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox6load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='成品'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "PARAID";
            comboBox6.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox7load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='成品'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "PARAID";
            comboBox7.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox8load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND]='成品'  ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "PARAID";
            comboBox8.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox9load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND] IN ('成品','原料','物料' ) ORDER BY [KIND],ID  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox9.DataSource = dt.DefaultView;
            comboBox9.ValueMember = "PARAID";
            comboBox9.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void comboBox10load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[KIND],[PARAID],[PARANAME]  FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND] IN ('成品','原料','物料' ) ORDER BY [KIND],ID  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox10.DataSource = dt.DefaultView;
            comboBox10.ValueMember = "PARAID";
            comboBox10.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void comboBox11load()
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
            Sequel.AppendFormat(@" SELECT [KIND] FROM [TKRESEARCH].[dbo].[TBPARA] WHERE [KIND] IN ('原料','成品','物料') GROUP BY [KIND] ORDER BY [KIND] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KIND", typeof(string));
         
            da.Fill(dt);
            comboBox11.DataSource = dt.DefaultView;
            comboBox11.ValueMember = "KIND";
            comboBox11.DisplayMember = "KIND";
            sqlConn.Close();


        }
        public void comboBox12load()
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
            Sequel.AppendFormat(@" SELECT  'N' [PARAID] UNION ALL SELECT  'Y' [PARAID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
         
            da.Fill(dt);
            comboBox12.DataSource = dt.DefaultView;
            comboBox12.ValueMember = "PARAID";
            comboBox12.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void SEARCH(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                        [ID] 
                                        ,[DOCID] AS '文件編號'
                                        ,[COMMENTS]  AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                     
                                        ,[DOCNAMES]
                                        FROM [TKRESEARCH].[dbo].[TBDB1]
                                        WHERE [DOCID] LIKE '%{0}%'
                                        ORDER BY [ID] 

                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                         [ID] 
                                        ,[DOCID] AS '文件編號'
                                        ,[COMMENTS]  AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                      
                                        ,[DOCNAMES]
                                        FROM [TKRESEARCH].[dbo].[TBDB1]
                                        ORDER BY [ID] 
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds1.Tables["ds1"];

                    dataGridView1.AutoResizeColumns();



                }
                else
                {
                    dataGridView1.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void SEARCH2(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [ID]
                                        ,[NAMES] AS '產品品名'
                                        ,[CHARS] AS '產品特色'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PRICES] AS '售價'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVEMETHODS] AS '保存方式'
                                        ,[PRIMES] AS '素別'
                                        ,[ALLERGENS] AS '過敏原'
                                        ,[OWNERS] AS '負責廠商'
                                        ,[EATES] AS '建議食用方式'
                                        ,[CHECKSUNITS] AS '驗證單位'
                                        ,[CHECKS] AS '驗證證書字號'
                                        ,[OTHERS] AS '其他口味'
                                        ,[COMMENTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES] AS '產品圖片' 
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB2]
                                        WHERE [NAMES] LIKE '%{0}%'
                                        ORDER BY [ID] 

                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                        [ID]
                                        ,[NAMES] AS '產品品名'
                                        ,[CHARS] AS '產品特色'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PRICES] AS '售價'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVEMETHODS] AS '保存方式'
                                        ,[PRIMES] AS '素別'
                                        ,[ALLERGENS] AS '過敏原'
                                        ,[OWNERS] AS '負責廠商'
                                        ,[EATES] AS '建議食用方式'
                                        ,[CHECKSUNITS] AS '驗證單位'
                                        ,[CHECKS] AS '驗證證書字號'
                                        ,[OTHERS] AS '其他口味'
                                        ,[COMMENTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES] AS '產品圖片'
                                       
                                        FROM [TKRESEARCH].[dbo].[TBDB2]                                       
                                        ORDER BY [ID] 
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds1.Tables["ds1"];

                    dataGridView2.AutoResizeColumns();



                }
                else
                {
                    dataGridView2.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void SEARCH3(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                       SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線過敏原'
                                        ,[PLACES] AS '產地'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品色澤'
                                        ,[TASTES] AS '產品風味'
                                        ,[LOTS] AS '產品批號'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[BASEONS] AS '基改/非基改'
                                        ,[COA] AS '檢附COA'
                                        ,[INCHECKRATES] AS '抽驗頻率'
                                        ,[RULES] AS '相關法規'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'  
                                        ,[DOCNAMES1] AS '三證'
                                        ,[DOCNAMES2] AS '產品成分'
                                        ,[DOCNAMES3] AS '基改/非基改'
                                        ,[DOCNAMES4] AS '檢驗報告'
                                        ,[DOCNAMES5] AS '營養標示'
                                        ,[DOCNAMES6] AS '產品圖片'                                     
                                        FROM [TKRESEARCH].[dbo].[TBDB3]
                                 
                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[SPECS] AS '產品規格'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線過敏原'
                                        ,[PLACES] AS '產地'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品色澤'
                                        ,[TASTES] AS '產品風味'
                                        ,[LOTS] AS '產品批號'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[BASEONS] AS '基改/非基改'
                                        ,[COA] AS '檢附COA'
                                        ,[INCHECKRATES] AS '抽驗頻率'
                                        ,[RULES] AS '相關法規'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'  
                                        ,[DOCNAMES1] AS '三證'
                                        ,[DOCNAMES2] AS '產品成分'
                                        ,[DOCNAMES3] AS '基改/非基改'
                                        ,[DOCNAMES4] AS '檢驗報告'
                                        ,[DOCNAMES5] AS '營養標示'
                                        ,[DOCNAMES6] AS '產品圖片'                                             
                                        FROM [TKRESEARCH].[dbo].[TBDB3]
                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds1.Tables["ds1"];

                    dataGridView3.AutoResizeColumns();



                }
                else
                {
                    dataGridView3.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        //設定下載欄
        public void SETdataGridView1()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload";
            lnkDownload.HeaderText = "Download";
            lnkDownload.Text = "Download";

            dataGridView1.Columns.Insert(dataGridView1.ColumnCount, lnkDownload);
            dataGridView1.CellContentClick += new DataGridViewCellEventHandler(DataGridView1_CellClick);
        }
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            DataGridView dgv = (DataGridView)sender;
            string columnName = dgv.Columns[e.ColumnIndex].Name;

            try
            {
                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload"))
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                            SELECT
                                             [ID] 
                                            ,[DOCID] AS '文件編號'
                                            ,[COMMENTS]  AS '備註'
                                            ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                            ,[DATAS]  
                                            ,[DOCNAMES]
                                            ,[CONTENTTYPES]
                                            FROM [TKRESEARCH].[dbo].[TBDB1]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS"];
                                contentType = sdr["CONTENTTYPES"].ToString();
                                fileName = sdr["DOCNAMES"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView2()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload31";
            lnkDownload.HeaderText = "Download";
            lnkDownload.Text = "Download";

            dataGridView2.Columns.Insert(dataGridView2.ColumnCount, lnkDownload);
            dataGridView2.CellContentClick += new DataGridViewCellEventHandler(DataGridView2_CellClick);
        }
        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload31"))
                {
                    DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                            SELECT 
                                            [ID]
                                            ,[NAMES]
                                            ,[CHARS]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PRICES]
                                            ,[SAVEDAYS]
                                            ,[SAVEMETHODS]
                                            ,[PRIMES]
                                            ,[ALLERGENS]
                                            ,[OWNERS]
                                            ,[EATES]
                                            ,[CHECKSUNITS]
                                            ,[CHECKS]
                                            ,[OTHERS]
                                            ,[COMMENTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES]
                                            ,[CONTENTTYPES]
                                            ,[DATAS]
                                            FROM [TKRESEARCH].[dbo].[TBDB2]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS"];
                                contentType = sdr["CONTENTTYPES"].ToString();
                                fileName = sdr["DOCNAMES"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

           
        }

        public void SETdataGridView31()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload31";
            lnkDownload.HeaderText = "三證Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView31_CellClick);

        }
        private void DataGridView31_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload31"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }
            
        }

        public void SETdataGridView32()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload32";
            lnkDownload.HeaderText = "產品成分Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView32_CellClick);
            
        }
        private void DataGridView32_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload32"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

           
        }

        public void SETdataGridView33()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload33";
            lnkDownload.HeaderText = "基改/非基改Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView33_CellClick);

        }
        private void DataGridView33_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload33"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS3"];
                                contentType = sdr["CONTENTTYPES3"].ToString();
                                fileName = sdr["DOCNAMES3"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView34()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload34";
            lnkDownload.HeaderText = "檢驗報告Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView34_CellClick);

        }
        private void DataGridView34_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload34"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS4"];
                                contentType = sdr["CONTENTTYPES4"].ToString();
                                fileName = sdr["DOCNAMES4"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView35()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload35";
            lnkDownload.HeaderText = "營養標示Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView35_CellClick);

        }
        private void DataGridView35_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload35"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS5"];
                                contentType = sdr["CONTENTTYPES5"].ToString();
                                fileName = sdr["DOCNAMES5"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }


            
        }

        public void SETdataGridView36()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload36";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView3.Columns.Insert(dataGridView3.ColumnCount, lnkDownload);
            dataGridView3.CellContentClick += new DataGridViewCellEventHandler(DataGridView36_CellClick);

        }
        private void DataGridView36_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload36"))
                {
                    DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[SPECS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PLACES]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[LOTS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[BASEONS]
                                            ,[COA]
                                            ,[INCHECKRATES]
                                            ,[RULES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            ,[DOCNAMES4]
                                            ,[CONTENTTYPES4]
                                            ,[DATAS4]
                                            ,[DOCNAMES5]
                                            ,[CONTENTTYPES5]
                                            ,[DATAS5]
                                            ,[DOCNAMES6]
                                            ,[CONTENTTYPES6]
                                            ,[DATAS6]
                                            FROM [TKRESEARCH].[dbo].[TBDB3]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS6"];
                                contentType = sdr["CONTENTTYPES6"].ToString();
                                fileName = sdr["DOCNAMES6"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView41()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload41";
            lnkDownload.HeaderText = "三證Download";
            lnkDownload.Text = "Download";

            dataGridView4.Columns.Insert(dataGridView4.ColumnCount, lnkDownload);
            dataGridView4.CellContentClick += new DataGridViewCellEventHandler(DataGridView41_CellClick);

        }
        private void DataGridView41_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload41"))
                {
                    DataGridViewRow row = dataGridView4.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT 
                                             [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[SPECS]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[COA]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB4]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

           
        }

        public void SETdataGridView42()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload42";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView4.Columns.Insert(dataGridView4.ColumnCount, lnkDownload);
            dataGridView4.CellContentClick += new DataGridViewCellEventHandler(DataGridView42_CellClick);

        }
        private void DataGridView42_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload42"))
                {
                    DataGridViewRow row = dataGridView4.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                          SELECT 
                                             [ID]
                                            ,[KINDS]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[SPECS]
                                            ,[OUTS]
                                            ,[COLORS]
                                            ,[CHECKS]
                                            ,[SAVEDAYS]
                                            ,[COA]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB4]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }


            
        }

        public void SETdataGridView51()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload51";
            lnkDownload.HeaderText = "三證Download";
            lnkDownload.Text = "Download";

            dataGridView5.Columns.Insert(dataGridView5.ColumnCount, lnkDownload);
            dataGridView5.CellContentClick += new DataGridViewCellEventHandler(DataGridView51_CellClick);

        }
        private void DataGridView51_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload51"))
                {
                    DataGridViewRow row = dataGridView5.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                          SELECT
                                            [ID]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB5]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView52()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload52";
            lnkDownload.HeaderText = "產品成分Download";
            lnkDownload.Text = "Download";

            dataGridView5.Columns.Insert(dataGridView5.ColumnCount, lnkDownload);
            dataGridView5.CellContentClick += new DataGridViewCellEventHandler(DataGridView52_CellClick);

        }
        private void DataGridView52_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload52"))
                {
                    DataGridViewRow row = dataGridView5.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT
                                            [ID]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB5]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }


            
        }

        public void SETdataGridView53()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload53";
            lnkDownload.HeaderText = "產品責任險Download";
            lnkDownload.Text = "Download";

            dataGridView5.Columns.Insert(dataGridView5.ColumnCount, lnkDownload);
            dataGridView5.CellContentClick += new DataGridViewCellEventHandler(DataGridView53_CellClick);

        }
        private void DataGridView53_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload53"))
                {
                    DataGridViewRow row = dataGridView5.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                           SELECT
                                            [ID]
                                            ,[SUPPLYS]
                                            ,[NAMES]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB5]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS3"];
                                contentType = sdr["CONTENTTYPES3"].ToString();
                                fileName = sdr["DOCNAMES3"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView61()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload61";
            lnkDownload.HeaderText = "營養標示Download";
            lnkDownload.Text = "Download";

            dataGridView6.Columns.Insert(dataGridView6.ColumnCount, lnkDownload);
            dataGridView6.CellContentClick += new DataGridViewCellEventHandler(DataGridView61_CellClick);

        }
        private void DataGridView61_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload61"))
                {
                    DataGridViewRow row = dataGridView6.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                          SELECT 
                                            [ID]
                                            ,[KINDS]
                                            ,[IANUMERS]
                                            ,[REGISTERNO]
                                            ,[MANUNAMES]
                                            ,[ADDRESS]
                                            ,[CHECKS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[MANUS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PRIMES]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[CHARS]
                                            ,[PACKAGES]
                                            ,[WEIGHTS]
                                            ,[SPECS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB6]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

           
        }

        public void SETdataGridView62()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload62";
            lnkDownload.HeaderText = "產品成分Download";
            lnkDownload.Text = "Download";

            dataGridView6.Columns.Insert(dataGridView6.ColumnCount, lnkDownload);
            dataGridView6.CellContentClick += new DataGridViewCellEventHandler(DataGridView62_CellClick);

        }
        private void DataGridView62_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload62"))
                {
                    DataGridViewRow row = dataGridView6.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT 
                                            [ID]
                                            ,[KINDS]
                                            ,[IANUMERS]
                                            ,[REGISTERNO]
                                            ,[MANUNAMES]
                                            ,[ADDRESS]
                                            ,[CHECKS]
                                            ,[NAMES]
                                            ,[ORIS]
                                            ,[MANUS]
                                            ,[PROALLGENS]
                                            ,[MANUALLGENS]
                                            ,[PRIMES]
                                            ,[COLORS]
                                            ,[TASTES]
                                            ,[CHARS]
                                            ,[PACKAGES]
                                            ,[WEIGHTS]
                                            ,[SPECS]
                                            ,[SAVEDAYS]
                                            ,[SAVECONDITIONS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            FROM [TKRESEARCH].[dbo].[TBDB6]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView71()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload71";
            lnkDownload.HeaderText = "產前會議Download";
            lnkDownload.Text = "Download";

            dataGridView7.Columns.Insert(dataGridView7.ColumnCount, lnkDownload);
            dataGridView7.CellContentClick += new DataGridViewCellEventHandler(DataGridView71_CellClick);

        }
        private void DataGridView71_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload71"))
                {
                    DataGridViewRow row = dataGridView7.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT 
                                         [ID]
                                        ,[KINDS]
                                        ,[NAMES]
                                        ,[COMMEMTS]
                                        ,[CREATEDATES]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        FROM [TKRESEARCH].[dbo].[TBDB7]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView72()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload72";
            lnkDownload.HeaderText = "技術移轉表Download";
            lnkDownload.Text = "Download";

            dataGridView7.Columns.Insert(dataGridView7.ColumnCount, lnkDownload);
            dataGridView7.CellContentClick += new DataGridViewCellEventHandler(DataGridView72_CellClick);

        }
        private void DataGridView72_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload72"))
                {
                    DataGridViewRow row = dataGridView7.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                        SELECT 
                                         [ID]
                                        ,[KINDS]
                                        ,[NAMES]
                                        ,[COMMEMTS]
                                        ,[CREATEDATES]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        FROM [TKRESEARCH].[dbo].[TBDB7]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView73()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload73";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView7.Columns.Insert(dataGridView7.ColumnCount, lnkDownload);
            dataGridView7.CellContentClick += new DataGridViewCellEventHandler(DataGridView73_CellClick);

        }
        private void DataGridView73_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload73"))
                {
                    DataGridViewRow row = dataGridView7.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT 
                                         [ID]
                                        ,[KINDS]
                                        ,[NAMES]
                                        ,[COMMEMTS]
                                        ,[CREATEDATES]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        FROM [TKRESEARCH].[dbo].[TBDB7]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS3"];
                                contentType = sdr["CONTENTTYPES3"].ToString();
                                fileName = sdr["DOCNAMES3"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView81()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload81";
            lnkDownload.HeaderText = "研發紀錄表Download";
            lnkDownload.Text = "Download";

            dataGridView8.Columns.Insert(dataGridView8.ColumnCount, lnkDownload);
            dataGridView8.CellContentClick += new DataGridViewCellEventHandler(DataGridView81_CellClick);

        }
        private void DataGridView81_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload81"))
                {
                    DataGridViewRow row = dataGridView8.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT 
                                            [ID]
                                            ,[KINDS]
                                            ,[NAMES]
                                            ,[RECORDS]
                                            ,[REPORTS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB8]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }
            
        }

        public void SETdataGridView82()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload72";
            lnkDownload.HeaderText = "產品圖片Download";
            lnkDownload.Text = "Download";

            dataGridView8.Columns.Insert(dataGridView8.ColumnCount, lnkDownload);
            dataGridView8.CellContentClick += new DataGridViewCellEventHandler(DataGridView82_CellClick);

        }
        private void DataGridView82_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload72"))
                {
                    DataGridViewRow row = dataGridView8.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                       SELECT 
                                            [ID]
                                            ,[KINDS]
                                            ,[NAMES]
                                            ,[RECORDS]
                                            ,[REPORTS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB8]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS2"];
                                contentType = sdr["CONTENTTYPES2"].ToString();
                                fileName = sdr["DOCNAMES2"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

        }

        public void SETdataGridView83()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload83";
            lnkDownload.HeaderText = "產品測試報告片Download";
            lnkDownload.Text = "Download";

            dataGridView8.Columns.Insert(dataGridView8.ColumnCount, lnkDownload);
            dataGridView8.CellContentClick += new DataGridViewCellEventHandler(DataGridView83_CellClick);

        }
        private void DataGridView83_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload83"))
                {
                    DataGridViewRow row = dataGridView8.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                        SELECT 
                                            [ID]
                                            ,[KINDS]
                                            ,[NAMES]
                                            ,[RECORDS]
                                            ,[REPORTS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            ,[DOCNAMES2]
                                            ,[CONTENTTYPES2]
                                            ,[DATAS2]
                                            ,[DOCNAMES3]
                                            ,[CONTENTTYPES3]
                                            ,[DATAS3]
                                            FROM [TKRESEARCH].[dbo].[TBDB8]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS3"];
                                contentType = sdr["CONTENTTYPES3"].ToString();
                                fileName = sdr["DOCNAMES3"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }

            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        public void SETdataGridView91()
        {
            DataGridViewLinkColumn lnkDownload = new DataGridViewLinkColumn();
            lnkDownload.UseColumnTextForLinkValue = true;
            lnkDownload.LinkBehavior = LinkBehavior.SystemDefault;
            lnkDownload.Name = "lnkDownload91";
            lnkDownload.HeaderText = "檔案Download";
            lnkDownload.Text = "Download";

            dataGridView9.Columns.Insert(dataGridView9.ColumnCount, lnkDownload);
            dataGridView9.CellContentClick += new DataGridViewCellEventHandler(DataGridView91_CellClick);

        }
        private void DataGridView91_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                StringBuilder SQL = new StringBuilder();

                DataGridView dgv = (DataGridView)sender;
                string columnName = dgv.Columns[e.ColumnIndex].Name;

                if (e.RowIndex >= 0 && columnName.Equals("lnkDownload91"))
                {
                    DataGridViewRow row = dataGridView9.Rows[e.RowIndex];

                    int ID = Convert.ToInt16((row.Cells["ID"].Value));
                    byte[] bytes;
                    string fileName, contentType;

                    // 20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    using (SqlConnection con = sqlConn)
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            SQL.AppendFormat(@"
                                         SELECT 
                                            [ID]
                                            ,[NAMES]
                                            ,[CONTENTS]
                                            ,[COMMEMTS]
                                            ,[CREATEDATES]
                                            ,[DOCNAMES1]
                                            ,[CONTENTTYPES1]
                                            ,[DATAS1]
                                            FROM [TKRESEARCH].[dbo].[TBDB9]
                                            WHERE ID=@ID
                                            ORDER BY [ID] 
                                       
                                            ");
                            cmd.CommandText = SQL.ToString();
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Connection = con;
                            con.Open();

                            using (SqlDataReader sdr = cmd.ExecuteReader())
                            {
                                sdr.Read();
                                bytes = (byte[])sdr["DATAS1"];
                                contentType = sdr["CONTENTTYPES1"].ToString();
                                fileName = sdr["DOCNAMES1"].ToString();

                                Stream stream;
                                SaveFileDialog saveFileDialog = new SaveFileDialog();
                                saveFileDialog.Filter = "All files (*.*)|*.*";
                                saveFileDialog.FilterIndex = 1;
                                saveFileDialog.RestoreDirectory = true;
                                saveFileDialog.FileName = fileName;
                                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                {
                                    stream = saveFileDialog.OpenFile();
                                    stream.Write(bytes, 0, bytes.Length);
                                    stream.Close();
                                }
                            }
                        }
                        con.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("下載失敗-附件檔異常");
            }

            
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1B.Text = null;
            textBox1C.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox131.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (dataGridView1.CurrentCell.RowIndex > 0 || dataGridView1.CurrentCell.ColumnIndex > 0)
                {
                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;                }


                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1B.Text = row.Cells["ID"].Value.ToString();
                    textBox1C.Text = row.Cells["ID"].Value.ToString();
                    textBox14.Text = row.Cells["文件編號"].Value.ToString();
                    textBox15.Text = row.Cells["備註"].Value.ToString();
                    textBox131.Text = row.Cells["文件編號"].Value.ToString();

                }
                else
                {
                    textBox1B.Text = null;
                    textBox1C.Text = null;
                    textBox14.Text = null;
                    textBox15.Text = null;
                    textBox131.Text = null;

                }
            }

        }
       

    

        public void OPEN1()
        {
            string FILETYPE = null;
            CONTENTTYPES1 = "";
            BYTES1 = null;
            DOCNAMES1 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;
                    DOCNAMES1 = Path.GetFileName(fileName);
                    textBox13.Text = fileName;

                    BYTES1 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES1 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES1 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES1 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES1 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES1 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES1 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES1 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES1 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES1 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if(LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1+" 檔案超過10M，無法上傳");

                       
                    }



                }
            }
        }

        public void OPEN2()
        {
            string FILETYPE = null;
            CONTENTTYPES2 = "";
            BYTES2 = null;
            DOCNAMES2 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES2 = Path.GetFileName(fileName);
                    textBox216.Text = fileName;

                    BYTES2 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES2 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES2 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES2 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES2 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES2 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES2 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES2 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES2 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES2 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN31()
        {
            string FILETYPE = null;
            CONTENTTYPES31 = "";
            BYTES31 = null;
            DOCNAMES31 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES31 = Path.GetFileName(fileName);
                    textBox320.Text = fileName;

                    BYTES31 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES31 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES31 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES31 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES31 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES31 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES31 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES31 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES31 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES31 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }


                }
            }
        }

        public void OPEN32()
        {
            string FILETYPE = null;
            CONTENTTYPES32 = "";
            BYTES32 = null;
            DOCNAMES32 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES32 = Path.GetFileName(fileName);
                    textBox321.Text = fileName;

                    BYTES32 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES32 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES32 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES32 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES32 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES32 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES32 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES32 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES32 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES32 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }


                }
            }
        }

        public void OPEN33()
        {
            string FILETYPE = null;
            CONTENTTYPES33 = "";
            BYTES33 = null;
            DOCNAMES33 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES33 = Path.GetFileName(fileName);
                    textBox322.Text = fileName;

                    BYTES33 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES33 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES33 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES33 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES33 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES33 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES33 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES33 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES33 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES33 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }


                }
            }
        }

        public void OPEN34()
        {
            string FILETYPE = null;
            CONTENTTYPES34 = "";
            BYTES34 = null;
            DOCNAMES34 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES34 = Path.GetFileName(fileName);
                    textBox323.Text = fileName;

                    BYTES34 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES34 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES34 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES34 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES34 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES34 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES34 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES34 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES34 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES34 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }


                }
            }
        }

        public void OPEN35()
        {
            string FILETYPE = null;
            CONTENTTYPES35 = "";
            BYTES35 = null;
            DOCNAMES35 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES35 = Path.GetFileName(fileName);
                    textBox324.Text = fileName;

                    BYTES35 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES35 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES35 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES35 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES35 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES35 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES35 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES35 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES35 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES35 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN36()
        {
            string FILETYPE = null;
            CONTENTTYPES36 = "";
            BYTES36 = null;
            DOCNAMES36 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES36 = Path.GetFileName(fileName);
                    textBox325.Text = fileName;

                    BYTES36 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES36 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES36 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES36 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES36 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES36 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES36 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES36 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES36 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES36 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN41()
        {
            string FILETYPE = null;
            CONTENTTYPES41 = "";
            BYTES41 = null;
            DOCNAMES41 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES41 = Path.GetFileName(fileName);
                    textBox410.Text = fileName;

                    BYTES41 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES41 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES41 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES41 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES41 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES41 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES41 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES41 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES41 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES41 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }


                }
            }
        }

        public void OPEN42()
        {
            string FILETYPE = null;
            CONTENTTYPES42 = "";
            BYTES42 = null;
            DOCNAMES42 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES42 = Path.GetFileName(fileName);
                    textBox411.Text = fileName;

                    BYTES42 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES42 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES42 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES42 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES42 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES42 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES42 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES42 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES42 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES42 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }
        public void OPEN51()
        {
            string FILETYPE = null;
            CONTENTTYPES51 = "";
            BYTES51 = null;
            DOCNAMES51 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES51 = Path.GetFileName(fileName);
                    textBox504.Text = fileName;

                    BYTES51 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES51 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES51 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES51 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES51 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES51 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES51 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES51 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES51 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES51 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }
        public void OPEN52()
        {
            string FILETYPE = null;
            CONTENTTYPES52 = "";
            BYTES52 = null;
            DOCNAMES52 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES52 = Path.GetFileName(fileName);
                    textBox505.Text = fileName;

                    BYTES52 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES52 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES52 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES52 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES52 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES52 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES52 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES52 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES52 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES52 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }
        public void OPEN53()
        {
            string FILETYPE = null;
            CONTENTTYPES53 = "";
            BYTES53 = null;
            DOCNAMES53 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES53 = Path.GetFileName(fileName);
                    textBox506.Text = fileName;

                    BYTES53 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES53 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES53 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES53 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES53 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES53 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES53 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES53 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES53 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES53 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN61()
        {
            string FILETYPE = null;
            CONTENTTYPES61 = "";
            BYTES61 = null;
            DOCNAMES61 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES61 = Path.GetFileName(fileName);
                    textBox620.Text = fileName;

                    BYTES61 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES61 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES61 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES61 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES61 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES61 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES61 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES61 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES61 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES61 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }
        public void OPEN62()
        {
            string FILETYPE = null;
            CONTENTTYPES62 = "";
            BYTES62 = null;
            DOCNAMES62 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES62 = Path.GetFileName(fileName);
                    textBox621.Text = fileName;

                    BYTES62 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES62 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES62 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES62 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES62 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES62 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES62 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES62 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES62 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES62 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN641()
        {
            string FILETYPE = null;
            CONTENTTYPES641 = "";
            BYTES641 = null;
            DOCNAMES641 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES641 = Path.GetFileName(fileName);
                    textBox6643.Text = fileName;

                    BYTES641 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES641 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES641 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES641 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES641 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES641 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES641 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES641 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES641 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES641 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN642()
        {
            string FILETYPE = null;
            CONTENTTYPES642 = "";
            BYTES642 = null;
            DOCNAMES642 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES642 = Path.GetFileName(fileName);
                    textBox6644.Text = fileName;

                    BYTES642 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES642 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES642 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES642 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES642 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES642 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES642 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES642 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES642 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES642 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }


        public void OPEN71()
        {
            string FILETYPE = null;
            CONTENTTYPES71 = "";
            BYTES71 = null;
            DOCNAMES71 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES71 = Path.GetFileName(fileName);
                    textBox703.Text = fileName;

                    BYTES71 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES71 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES71 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES71 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES71 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES71 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES71 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES71 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES71 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES71 = "image/bmp";
                            break;
                    }
                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }

                }
            }
        }

        public void OPEN72()
        {
            string FILETYPE = null;
            CONTENTTYPES72= "";
            BYTES72 = null;
            DOCNAMES72 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES72 = Path.GetFileName(fileName);
                    textBox704.Text = fileName;

                    BYTES72 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES72 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES72 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES72 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES72 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES72 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES72 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES72 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES72 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES72 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }

        public void OPEN73()
        {
            string FILETYPE = null;
            CONTENTTYPES73 = "";
            BYTES73 = null;
            DOCNAMES73 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES73 = Path.GetFileName(fileName);
                    textBox705.Text = fileName;

                    BYTES73 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES73 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES73 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES73 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES73 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES73 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES73 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES73 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES73 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES73 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }

        public void OPEN81()
        {
            string FILETYPE = null;
            CONTENTTYPES81 = "";
            BYTES81 = null;
            DOCNAMES81 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES81 = Path.GetFileName(fileName);
                    textBox805.Text = fileName;

                    BYTES81 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES81 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES81 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES81 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES81 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES81 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES81 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES81 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES81 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES81 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }


        public void OPEN82()
        {
            string FILETYPE = null;
            CONTENTTYPES82 = "";
            BYTES82 = null;
            DOCNAMES82 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES82 = Path.GetFileName(fileName);
                    textBox806.Text = fileName;

                    BYTES82 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES82 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES82 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES82 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES82 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES82 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES82 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES82 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES82 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES82 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }

        public void OPEN83()
        {
            string FILETYPE = null;
            CONTENTTYPES83 = "";
            BYTES83 = null;
            DOCNAMES83 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES83 = Path.GetFileName(fileName);
                    textBox807.Text = fileName;

                    BYTES83 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES83 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES83 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES83 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES83 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES83 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES83 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES83 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES83 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES83 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }

        public void OPEN91()
        {
            string FILETYPE = null;
            CONTENTTYPES91 = "";
            BYTES91 = null;
            DOCNAMES91 = null;

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog1.FileName;

                    DOCNAMES91 = Path.GetFileName(fileName);
                    textBox904.Text = fileName;

                    BYTES91 = File.ReadAllBytes(fileName);

                    //Set the contenttype based on File Extension

                    switch (Path.GetExtension(fileName))
                    {
                        case ".docx":
                            CONTENTTYPES91 = "application/msword";
                            break;
                        case ".doc":
                            CONTENTTYPES91 = "application/msword";
                            break;
                        case ".xls":
                            CONTENTTYPES91 = "application/vnd.ms-excel";
                            break;
                        case ".xlsx":
                            CONTENTTYPES91 = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            break;
                        case ".pdf":
                            CONTENTTYPES91 = "application/pdf";
                            break;
                        case ".jpg":
                            CONTENTTYPES91 = "image/jpeg";
                            break;
                        case ".png":
                            CONTENTTYPES91 = "image/png";
                            break;
                        case ".gif":
                            CONTENTTYPES91 = "image/gif";
                            break;
                        case ".bmp":
                            CONTENTTYPES91 = "image/bmp";
                            break;
                    }

                    long LONG_FILESIZE = openFileDialog1.OpenFile().Length;
                    if (LONG_FILESIZE > FILESIZE)
                    {
                        MessageBox.Show(DOCNAMES1 + " 檔案超過10M，無法上傳");


                    }
                }
            }
        }


        public void ADD_TO_TBDB1(string DOCID,string COMMENTS, string DOCNAMES, string CONTENTTYPES,byte[] BYTES)
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(DOCID))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB1]
                                        (
                                        [DOCID]
                                        ,[COMMENTS] 
                                        ,[DOCNAMES]
                                        ,[CONTENTTYPES]
                                        ,[DATAS]
                                        )
                                        VALUES
                                        (
                                        @DOCID
                                        ,@COMMENTS                                   
                                        ,@DOCNAMES
                                        ,@CONTENTTYPES
                                        ,@DATAS
                                        )
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@DOCID", DOCID);
                            cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                            cmd.Parameters.AddWithValue("@DOCNAMES", DOCNAMES);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES", CONTENTTYPES);
                            cmd.Parameters.AddWithValue("@DATAS", BYTES);
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }
            
           
                
            

          
        }
        public void  UPDATE_TO_TBDB1(string ID,string DOCID,string COMMENTS)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE[TKRESEARCH].[dbo].[TBDB1]
                                        SET DOCID=@DOCID,COMMENTS=@COMMENTS
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@DOCID", DOCID);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }

          
        }

        public void DELETE_TO_TBDB1(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB1]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                    
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }

          
        }

        public void ADD_TO_TBDB2(
                                string NAMES
                                , string CHARS
                                , string ORIS
                                , string SPECS
                                , string PRICES
                                , string SAVEDAYS
                                , string SAVEMETHODS
                                , string PRIMES
                                , string ALLERGENS
                                , string OWNERS
                                , string EATES
                                , string CHECKSUNITS
                                , string CHECKS
                                , string OTHERS
                                , string COMMENTS
                                , string DOCNAMES
                                , string CONTENTTYPES,
                                byte[] BYTES
                                )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"
                                       
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB2]
                                        (
                                        [NAMES]
                                        ,[CHARS]
                                        ,[ORIS]
                                        ,[SPECS]
                                        ,[PRICES]
                                        ,[SAVEDAYS]
                                        ,[SAVEMETHODS]
                                        ,[PRIMES]
                                        ,[ALLERGENS]
                                        ,[OWNERS]
                                        ,[EATES]
                                        ,[CHECKSUNITS]
                                        ,[CHECKS]
                                        ,[OTHERS]
                                        ,[COMMENTS]
                                        ,[DOCNAMES]
                                        ,[CONTENTTYPES]
                                        ,[DATAS]
                                        )
                                        VALUES
                                        (
                                        @NAMES
                                        ,@CHARS
                                        ,@ORIS
                                        ,@SPECS
                                        ,@PRICES
                                        ,@SAVEDAYS
                                        ,@SAVEMETHODS
                                        ,@PRIMES
                                        ,@ALLERGENS
                                        ,@OWNERS
                                        ,@EATES
                                        ,@CHECKSUNITS
                                        ,@CHECKS
                                        ,@OTHERS
                                        ,@COMMENTS
                                        ,@DOCNAMES
                                        ,@CONTENTTYPES
                                        ,@DATAS
                                        )
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@CHARS", CHARS);
                            cmd.Parameters.AddWithValue("@ORIS", ORIS);
                            cmd.Parameters.AddWithValue("@SPECS", SPECS);
                            cmd.Parameters.AddWithValue("@PRICES", PRICES);
                            cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                            cmd.Parameters.AddWithValue("@SAVEMETHODS", SAVEMETHODS);
                            cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                            cmd.Parameters.AddWithValue("@ALLERGENS", ALLERGENS);
                            cmd.Parameters.AddWithValue("@OWNERS", OWNERS);
                            cmd.Parameters.AddWithValue("@EATES", EATES);
                            cmd.Parameters.AddWithValue("@CHECKSUNITS", CHECKSUNITS);
                            cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                            cmd.Parameters.AddWithValue("@OTHERS", OTHERS);
                            cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                            cmd.Parameters.AddWithValue("@DOCNAMES", DOCNAMES);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES", CONTENTTYPES);
                            cmd.Parameters.AddWithValue("@DATAS", BYTES);
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }

            

           
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {

            textBox2B.Text = null;
            textBox2C.Text = null;
            textBox221.Text = null;
            textBox222.Text = null;
            textBox223.Text = null;
            textBox224.Text = null;
            textBox225.Text = null;
            textBox226.Text = null;
            textBox227.Text = null;
            textBox228.Text = null;
            textBox229.Text = null;
            textBox230.Text = null;
            textBox231.Text = null;
            textBox232.Text = null;
            textBox233.Text = null;
            textBox234.Text = null;
            textBox235.Text = null;
            textBox241.Text = null;


            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;               

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox2B.Text = row.Cells["ID"].Value.ToString();
                    textBox2C.Text = row.Cells["ID"].Value.ToString();
                    textBox221.Text = row.Cells["產品品名"].Value.ToString();
                    textBox222.Text = row.Cells["產品特色"].Value.ToString();
                    textBox223.Text = row.Cells["產品成分"].Value.ToString();
                    textBox224.Text = row.Cells["產品規格"].Value.ToString();
                    textBox225.Text = row.Cells["售價"].Value.ToString();
                    textBox226.Text = row.Cells["保存期限"].Value.ToString();
                    textBox227.Text = row.Cells["保存方式"].Value.ToString();
                    textBox228.Text = row.Cells["素別"].Value.ToString();
                    textBox229.Text = row.Cells["過敏原"].Value.ToString();
                    textBox230.Text = row.Cells["負責廠商"].Value.ToString();
                    textBox231.Text = row.Cells["建議食用方式"].Value.ToString();
                    textBox232.Text = row.Cells["驗證單位"].Value.ToString();
                    textBox233.Text = row.Cells["驗證證書字號"].Value.ToString();
                    textBox234.Text = row.Cells["其他口味"].Value.ToString();
                    textBox235.Text = row.Cells["備註"].Value.ToString();
                    textBox241.Text = row.Cells["產品品名"].Value.ToString();

                }
                else
                {
                    textBox2B.Text = null;
                    textBox2C.Text = null;
                    textBox221.Text = null;
                    textBox222.Text = null;
                    textBox223.Text = null;
                    textBox224.Text = null;
                    textBox225.Text = null;
                    textBox226.Text = null;
                    textBox227.Text = null;
                    textBox228.Text = null;
                    textBox229.Text = null;
                    textBox230.Text = null;
                    textBox231.Text = null;
                    textBox232.Text = null;
                    textBox233.Text = null;
                    textBox234.Text = null;
                    textBox235.Text = null;
                    textBox241.Text = null;

                }
            }
        }

        public void UPDATE_TO_TBDB2(string ID
                                , string NAMES
                                , string CHARS
                                , string ORIS
                                , string SPECS
                                , string PRICES
                                , string SAVEDAYS
                                , string SAVEMETHODS
                                , string PRIMES
                                , string ALLERGENS
                                , string OWNERS
                                , string EATES
                                , string CHECKSUNITS
                                , string CHECKS
                                , string OTHERS
                                , string COMMENTS
                                )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB2]
                                        SET [NAMES]=@NAMES
                                        ,[CHARS]=@CHARS
                                        ,[ORIS]=@ORIS
                                        ,[SPECS]=@SPECS
                                        ,[PRICES]=@PRICES
                                        ,[SAVEDAYS]=@SAVEDAYS
                                        ,[SAVEMETHODS]=@SAVEMETHODS
                                        ,[PRIMES]=@PRIMES
                                        ,[ALLERGENS]=@ALLERGENS
                                        ,[OWNERS]=@OWNERS
                                        ,[EATES]=@EATES
                                        ,[CHECKSUNITS]=@CHECKSUNITS
                                        ,[CHECKS]=@CHECKS
                                        ,[OTHERS]=@OTHERS
                                        ,[COMMENTS]=@COMMENTS
                                 
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@CHARS", CHARS);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PRICES", PRICES);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVEMETHODS", SAVEMETHODS);
                        cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                        cmd.Parameters.AddWithValue("@ALLERGENS", ALLERGENS);
                        cmd.Parameters.AddWithValue("@OWNERS", OWNERS);
                        cmd.Parameters.AddWithValue("@EATES", EATES);
                        cmd.Parameters.AddWithValue("@CHECKSUNITS", CHECKSUNITS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@OTHERS", OTHERS);
                        cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

      


        public void ADD_TO_TBDB3(
                                string KINDS
                                , string SUPPLYS
                                , string NAMES
                                , string ORIS
                                , string SPECS
                                , string PROALLGENS
                                , string MANUALLGENS
                                , string PLACES
                                , string OUTS
                                , string COLORS
                                , string TASTES
                                , string LOTS
                                , string CHECKS
                                , string SAVEDAYS
                                , string SAVECONDITIONS
                                , string BASEONS
                                , string COA
                                , string INCHECKRATES
                                , string RULES
                                , string COMMEMTS
                                , string DOCNAMES1
                                , string CONTENTTYPES1
                                , byte[] DATAS1
                                , string DOCNAMES2
                                , string CONTENTTYPES2
                                , byte[] DATAS2
                                , string DOCNAMES3
                                , string CONTENTTYPES3
                                , byte[] DATAS3
                                , string DOCNAMES4
                                , string CONTENTTYPES4
                                , byte[] DATAS4
                                , string DOCNAMES5
                                , string CONTENTTYPES5
                                , byte[] DATAS5
                                , string DOCNAMES6
                                , string CONTENTTYPES6
                                , byte[] DATAS6
                                )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"
                                       
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB3]
                                        (
                                         [KINDS]
                                        ,[SUPPLYS]
                                        ,[NAMES]
                                        ,[ORIS]
                                        ,[SPECS]
                                        ,[PROALLGENS]
                                        ,[MANUALLGENS]
                                        ,[PLACES]
                                        ,[OUTS]
                                        ,[COLORS]
                                        ,[TASTES]
                                        ,[LOTS]
                                        ,[CHECKS]
                                        ,[SAVEDAYS]
                                        ,[SAVECONDITIONS]
                                        ,[BASEONS]
                                        ,[COA]
                                        ,[INCHECKRATES]
                                        ,[RULES]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        ,[DOCNAMES4]
                                        ,[CONTENTTYPES4]
                                        ,[DATAS4]
                                        ,[DOCNAMES5]
                                        ,[CONTENTTYPES5]
                                        ,[DATAS5]
                                        ,[DOCNAMES6]
                                        ,[CONTENTTYPES6]
                                        ,[DATAS6]
                                       
                                        )
                                        VALUES
                                        (
                                          @KINDS
                                        ,@SUPPLYS
                                        ,@NAMES
                                        ,@ORIS
                                        ,@SPECS
                                        ,@PROALLGENS
                                        ,@MANUALLGENS
                                        ,@PLACES
                                        ,@OUTS
                                        ,@COLORS
                                        ,@TASTES
                                        ,@LOTS
                                        ,@CHECKS
                                        ,@SAVEDAYS
                                        ,@SAVECONDITIONS
                                        ,@BASEONS
                                        ,@COA
                                        ,@INCHECKRATES
                                        ,@RULES
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        ,@DOCNAMES3
                                        ,@CONTENTTYPES3
                                        ,@DATAS3
                                        ,@DOCNAMES4
                                        ,@CONTENTTYPES4
                                        ,@DATAS4
                                        ,@DOCNAMES5
                                        ,@CONTENTTYPES5
                                        ,@DATAS5
                                        ,@DOCNAMES6
                                        ,@CONTENTTYPES6
                                        ,@DATAS6
                                        )
                                        
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@KINDS", KINDS);
                            cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@ORIS", ORIS);
                            cmd.Parameters.AddWithValue("@SPECS", SPECS);
                            cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                            cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                            cmd.Parameters.AddWithValue("@PLACES", PLACES);
                            cmd.Parameters.AddWithValue("@OUTS", OUTS);
                            cmd.Parameters.AddWithValue("@COLORS", COLORS);
                            cmd.Parameters.AddWithValue("@TASTES", TASTES);
                            cmd.Parameters.AddWithValue("@LOTS", LOTS);
                            cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                            cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                            cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                            cmd.Parameters.AddWithValue("@BASEONS", BASEONS);
                            cmd.Parameters.AddWithValue("@COA", COA);
                            cmd.Parameters.AddWithValue("@INCHECKRATES", INCHECKRATES);
                            cmd.Parameters.AddWithValue("@RULES", RULES);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                            cmd.Parameters.AddWithValue("@DOCNAMES3", DOCNAMES3);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES3", CONTENTTYPES3);
                            cmd.Parameters.AddWithValue("@DATAS3", DATAS3);
                            cmd.Parameters.AddWithValue("@DOCNAMES4", DOCNAMES4);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES4", CONTENTTYPES4);
                            cmd.Parameters.AddWithValue("@DATAS4", DATAS4);
                            cmd.Parameters.AddWithValue("@DOCNAMES5", DOCNAMES5);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES5", CONTENTTYPES5);
                            cmd.Parameters.AddWithValue("@DATAS5", DATAS5);
                            cmd.Parameters.AddWithValue("@DOCNAMES6", DOCNAMES6);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES6", CONTENTTYPES6);
                            cmd.Parameters.AddWithValue("@DATAS6", DATAS6);

                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }
            


        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox331.Text = null;
            textBox332.Text = null;
            textBox333.Text = null;
            textBox334.Text = null;
            textBox335.Text = null;
            textBox336.Text = null;
            textBox337.Text = null;
            textBox338.Text = null;
            textBox339.Text = null;
            textBox340.Text = null;
            textBox341.Text = null;
            textBox342.Text = null;
            textBox343.Text = null;
            textBox344.Text = null;
            textBox345.Text = null;
            textBox346.Text = null;
            textBox347.Text = null;
            textBox348.Text = null;
            textBox349.Text = null;
            textBox350.Text = null;
            textBox3B.Text = null;
            textBox3C.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
              

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox3B.Text = row.Cells["ID"].Value.ToString();
                    textBox3C.Text = row.Cells["ID"].Value.ToString();

                    comboBox2.Text = row.Cells["分類"].Value.ToString();
                    textBox331.Text = row.Cells["供應商"].Value.ToString();
                    textBox332.Text = row.Cells["產品品名"].Value.ToString();
                    textBox333.Text = row.Cells["產品成分"].Value.ToString();
                    textBox334.Text = row.Cells["產品規格"].Value.ToString();
                    textBox335.Text = row.Cells["產品過敏原"].Value.ToString();
                    textBox336.Text = row.Cells["產線過敏原"].Value.ToString();
                    textBox337.Text = row.Cells["產地"].Value.ToString();
                    textBox338.Text = row.Cells["產品外觀"].Value.ToString();
                    textBox339.Text = row.Cells["產品色澤"].Value.ToString();
                    textBox340.Text = row.Cells["產品風味"].Value.ToString();
                    textBox341.Text = row.Cells["產品批號"].Value.ToString();
                    textBox342.Text = row.Cells["外包裝及驗收標準"].Value.ToString();
                    textBox343.Text = row.Cells["保存期限"].Value.ToString();
                    textBox344.Text = row.Cells["保存條件"].Value.ToString();
                    textBox345.Text = row.Cells["基改/非基改"].Value.ToString();
                    textBox346.Text = row.Cells["檢附COA"].Value.ToString();
                    textBox347.Text = row.Cells["抽驗頻率"].Value.ToString();
                    textBox348.Text = row.Cells["相關法規"].Value.ToString();
                    textBox349.Text = row.Cells["備註"].Value.ToString();

                    textBox350.Text = row.Cells["產品品名"].Value.ToString();
                    


                }
                else
                {
                    textBox331.Text = null;
                    textBox332.Text = null;
                    textBox333.Text = null;
                    textBox334.Text = null;
                    textBox335.Text = null;
                    textBox336.Text = null;
                    textBox337.Text = null;
                    textBox338.Text = null;
                    textBox339.Text = null;
                    textBox340.Text = null;
                    textBox341.Text = null;
                    textBox342.Text = null;
                    textBox343.Text = null;
                    textBox344.Text = null;
                    textBox345.Text = null;
                    textBox346.Text = null;
                    textBox347.Text = null;
                    textBox348.Text = null;
                    textBox349.Text = null;
                    textBox350.Text = null;
                    textBox3B.Text = null;
                    textBox3C.Text = null;
                }
            }

        }

        public void UPDATE_TO_TBDB3(
                                string ID
                               , string KINDS
                               , string SUPPLYS
                               , string NAMES
                               , string ORIS
                               , string SPECS
                               , string PROALLGENS
                               , string MANUALLGENS
                               , string PLACES
                               , string OUTS
                               , string COLORS
                               , string TASTES
                               , string LOTS
                               , string CHECKS
                               , string SAVEDAYS
                               , string SAVECONDITIONS
                               , string BASEONS
                               , string COA
                               , string INCHECKRATES
                               , string RULES
                               , string COMMEMTS
                              
                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB3]
                                        SET
                                        [KINDS]=@KINDS
                                        ,[SUPPLYS]=@SUPPLYS
                                        ,[NAMES]=@NAMES
                                        ,[ORIS]=@ORIS
                                        ,[SPECS]=@SPECS
                                        ,[PROALLGENS]=@PROALLGENS
                                        ,[MANUALLGENS]=@MANUALLGENS
                                        ,[PLACES]=@PLACES
                                        ,[OUTS]=@OUTS
                                        ,[COLORS]=@COLORS
                                        ,[TASTES]=@TASTES
                                        ,[LOTS]=@LOTS
                                        ,[CHECKS]=@CHECKS
                                        ,[SAVEDAYS]=@SAVEDAYS
                                        ,[SAVECONDITIONS]=@SAVECONDITIONS
                                        ,[BASEONS]=@BASEONS
                                        ,[COA]=@COA
                                        ,[INCHECKRATES]=@INCHECKRATES
                                        ,[RULES]=@RULES
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                        cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                        cmd.Parameters.AddWithValue("@PLACES", PLACES);
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);
                        cmd.Parameters.AddWithValue("@TASTES", TASTES);
                        cmd.Parameters.AddWithValue("@LOTS", LOTS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                        cmd.Parameters.AddWithValue("@BASEONS", BASEONS);
                        cmd.Parameters.AddWithValue("@COA", COA);
                        cmd.Parameters.AddWithValue("@INCHECKRATES", INCHECKRATES);
                        cmd.Parameters.AddWithValue("@RULES", RULES);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                       

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB2(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB2]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void DELETE_TO_TBDB3(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB3]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH4(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                      SELECT  
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[SPECS] AS '產品規格'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品顏色'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[COA] AS '檢附COA'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                        
                                        ,[DOCNAMES2] AS '產品圖片'
                                        
                                        FROM [TKRESEARCH].[dbo].[TBDB4]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT  
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[SPECS] AS '產品規格'
                                        ,[OUTS] AS '產品外觀'
                                        ,[COLORS] AS '產品顏色'
                                        ,[CHECKS] AS '外包裝及驗收標準'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[COA] AS '檢附COA'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                        
                                        ,[DOCNAMES2] AS '產品圖片'
                                        
                                        FROM [TKRESEARCH].[dbo].[TBDB4]
                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView4.DataSource = ds1.Tables["ds1"];

                    dataGridView4.AutoResizeColumns();



                }
                else
                {
                    dataGridView3.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB4(
                               string KINDS
                               , string SUPPLYS
                               , string NAMES                               
                               , string SPECS                             
                               , string OUTS
                               , string COLORS                            
                               , string CHECKS
                               , string SAVEDAYS                              
                               , string COA                              
                               , string COMMEMTS
                               , string DOCNAMES1
                               , string CONTENTTYPES1
                               , byte[] DATAS1
                               , string DOCNAMES2
                               , string CONTENTTYPES2
                               , byte[] DATAS2
                          
                               )
        {

            try
            {

                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"                                      
                                        
                                        INSERT INTO  [TKRESEARCH].[dbo].[TBDB4]
                                        (
                                        [KINDS]
                                        ,[SUPPLYS]
                                        ,[NAMES]
                                        ,[SPECS]
                                        ,[OUTS]
                                        ,[COLORS]
                                        ,[CHECKS]
                                        ,[SAVEDAYS]
                                        ,[COA]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        )
                                        VALUES
                                        (
                                        @KINDS
                                        ,@SUPPLYS
                                        ,@NAMES
                                        ,@SPECS
                                        ,@OUTS
                                        ,@COLORS
                                        ,@CHECKS
                                        ,@SAVEDAYS
                                        ,@COA
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        )
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@KINDS", KINDS);
                            cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@SPECS", SPECS);
                            cmd.Parameters.AddWithValue("@OUTS", OUTS);
                            cmd.Parameters.AddWithValue("@COLORS", COLORS);
                            cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                            cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                            cmd.Parameters.AddWithValue("@COA", COA);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);



                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }
            


        }

        public void UPDATE_TO_TBDB4(
                                string ID
                               , string KINDS
                               , string SUPPLYS
                               , string NAMES                            
                               , string SPECS                             
                               , string OUTS
                               , string COLORS                            
                               , string CHECKS
                               , string SAVEDAYS                              
                               , string COA                            
                               , string COMMEMTS

                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB4]
                                        SET
                                        [KINDS]=@KINDS
                                        ,[SUPPLYS]=@SUPPLYS
                                        ,[NAMES]=@NAMES                                       
                                        ,[SPECS]=@SPECS                                       
                                        ,[OUTS]=@OUTS
                                        ,[COLORS]=@COLORS                                        
                                        ,[CHECKS]=@CHECKS
                                        ,[SAVEDAYS]=@SAVEDAYS                                       
                                        ,[COA]=@COA                                      
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);                       
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);                       
                        cmd.Parameters.AddWithValue("@OUTS", OUTS);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);                        
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);                        
                        cmd.Parameters.AddWithValue("@COA", COA);                        
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);


                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBox4B.Text = null;
            textBox4C.Text = null;
            textBox420.Text = null;
            textBox421.Text = null;
            textBox422.Text = null;
            textBox423.Text = null;
            textBox424.Text = null;
            textBox425.Text = null;
            textBox426.Text = null;
            textBox427.Text = null;
            textBox428.Text = null;
            textBox430.Text = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox4B.Text = row.Cells["ID"].Value.ToString();
                    textBox4C.Text = row.Cells["ID"].Value.ToString();

                    comboBox4.Text= row.Cells["分類"].Value.ToString();
                    textBox420.Text = row.Cells["供應商"].Value.ToString();
                    textBox421.Text = row.Cells["產品品名"].Value.ToString();
                    textBox422.Text = row.Cells["產品規格"].Value.ToString();
                    textBox423.Text = row.Cells["產品外觀"].Value.ToString();
                    textBox424.Text = row.Cells["產品顏色"].Value.ToString();
                    textBox425.Text = row.Cells["外包裝及驗收標準"].Value.ToString();
                    textBox426.Text = row.Cells["保存期限"].Value.ToString();
                    textBox427.Text = row.Cells["檢附COA"].Value.ToString();
                    textBox428.Text = row.Cells["備註"].Value.ToString();
                    textBox430.Text = row.Cells["產品品名"].Value.ToString();


                }
                else
                {
                    textBox4B.Text = null;
                    textBox4C.Text = null;
                    textBox420.Text = null;
                    textBox421.Text = null;
                    textBox422.Text = null;
                    textBox423.Text = null;
                    textBox424.Text = null;
                    textBox425.Text = null;
                    textBox426.Text = null;
                    textBox427.Text = null;
                    textBox428.Text = null;
                    textBox430.Text = null;

                }
            }

        }

        public void DELETE_TO_TBDB4(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB4]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH5(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                      SELECT 
                                        [ID]
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                       
                                        ,[DOCNAMES2] AS '產品成分'
                                        
                                        ,[DOCNAMES3] AS '產品責任險'
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB5]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                         SELECT 
                                        [ID]
                                        ,[SUPPLYS] AS '供應商'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '三證'
                                       
                                        ,[DOCNAMES2] AS '產品成分'
                                        
                                        ,[DOCNAMES3] AS '產品責任險'
                                      
                                        FROM [TKRESEARCH].[dbo].[TBDB5]

                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView5.DataSource = ds1.Tables["ds1"];

                    dataGridView5.AutoResizeColumns();



                }
                else
                {
                    dataGridView5.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }
        public void ADD_TO_TBDB5(
                             string SUPPLYS
                             , string NAMES                            
                             , string COMMEMTS
                             , string DOCNAMES1
                             , string CONTENTTYPES1
                             , byte[] DATAS1
                             , string DOCNAMES2
                             , string CONTENTTYPES2
                             , byte[] DATAS2
                             , string DOCNAMES3
                             , string CONTENTTYPES3
                             , byte[] DATAS3

                             )
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"                                      
                                        
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB5]
                                        (
                                        [SUPPLYS]
                                        ,[NAMES]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        )
                                        VALUES
                                        (
                                        @SUPPLYS
                                        ,@NAMES
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        ,@DOCNAMES3
                                        ,@CONTENTTYPES3
                                        ,@DATAS3
                                        )
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);
                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                            cmd.Parameters.AddWithValue("@DOCNAMES3", DOCNAMES3);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES3", CONTENTTYPES3);
                            cmd.Parameters.AddWithValue("@DATAS3", DATAS3);



                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }

            }
            catch
            {
                MessageBox.Show("新增失敗");
            }


           


        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox5B.Text = null;
            textBox5C.Text = null;
            textBox521.Text = null;
            textBox522.Text = null;
            textBox523.Text = null;
            textBox531.Text = null;
      

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox5B.Text = row.Cells["ID"].Value.ToString();
                    textBox5C.Text = row.Cells["ID"].Value.ToString();
                   
                    textBox521.Text = row.Cells["供應商"].Value.ToString();
                    textBox522.Text = row.Cells["產品品名"].Value.ToString();
                    textBox523.Text = row.Cells["備註"].Value.ToString();
                    textBox531.Text = row.Cells["產品品名"].Value.ToString();


                }
                else
                {
                    textBox5B.Text = null;
                    textBox5C.Text = null;
                    textBox521.Text = null;
                    textBox522.Text = null;
                    textBox523.Text = null;
                    textBox531.Text = null;

                }
            }
        }

        public void UPDATE_TO_TBDB5(
                                string ID                              
                               , string SUPPLYS
                               , string NAMES                            
                               , string COMMEMTS

                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB5]
                                        SET
                                        [SUPPLYS]=@SUPPLYS
                                        ,[NAMES]=@NAMES          
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);                     
                        cmd.Parameters.AddWithValue("@SUPPLYS", SUPPLYS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);                       
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);


                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB5(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB5]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH6(string KEYS,string STATUS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            StringBuilder SQLUERY = new StringBuilder();

            DataSet ds1 = new DataSet();

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

                if(STATUS.Equals("N"))
                {
                    SQLUERY.AppendFormat(@"
                                            AND [NAMES] NOT LIKE '%暫停%'
                                            ");
                }
                else
                {
                    SQLUERY.AppendFormat(@"
                                            AND [NAMES] LIKE '%暫停%'
                                            ");
                }

                sbSql.Clear();

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                      SELECT 
                                        [ID] 
                                        ,[KINDS] AS '分類'
                                        ,[IANUMERS] AS '國際條碼'
                                        ,[REGISTERNO] AS '食品業者登錄字號'
                                        ,[MANUNAMES] AS '製造商名稱'
                                        ,[ADDRESS] AS '製造商地址'
                                        ,[CHECKS] AS '品質認證'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[MANUS] AS '製造流程'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線及生產設備過敏原'
                                        ,[PRIMES] AS '素別'
                                        ,[COLORS] AS '色澤'
                                        ,[TASTES] AS '風味'
                                        ,[CHARS] AS '性狀'
                                        ,[PACKAGES] AS '材質'
                                        ,[WEIGHTS] AS '淨重量'
                                        ,[SPECS] AS '規格'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '營養標示'
                                    
                                        ,[DOCNAMES2] AS '產品圖片'
                                       
                                        FROM [TKRESEARCH].[dbo].[TBDB6]

                                        WHERE NAMES LIKE '%{0}%'
                                        {1}
                                        ORDER BY  [ID]
                                    ", KEYS, SQLUERY.ToString());
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                         SELECT 
                                        [ID] 
                                        ,[KINDS] AS '分類'
                                        ,[IANUMERS] AS '國際條碼'
                                        ,[REGISTERNO] AS '食品業者登錄字號'
                                        ,[MANUNAMES] AS '製造商名稱'
                                        ,[ADDRESS] AS '製造商地址'
                                        ,[CHECKS] AS '品質認證'
                                        ,[NAMES] AS '產品品名'
                                        ,[ORIS] AS '產品成分'
                                        ,[MANUS] AS '製造流程'
                                        ,[PROALLGENS] AS '產品過敏原'
                                        ,[MANUALLGENS] AS '產線及生產設備過敏原'
                                        ,[PRIMES] AS '素別'
                                        ,[COLORS] AS '色澤'
                                        ,[TASTES] AS '風味'
                                        ,[CHARS] AS '性狀'
                                        ,[PACKAGES] AS '材質'
                                        ,[WEIGHTS] AS '淨重量'
                                        ,[SPECS] AS '規格'
                                        ,[SAVEDAYS] AS '保存期限'
                                        ,[SAVECONDITIONS] AS '保存條件'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '營養標示'
                                    
                                        ,[DOCNAMES2] AS '產品圖片'
                                       
                                        FROM [TKRESEARCH].[dbo].[TBDB6]
                                        WHERE 1=1
                                        {0}
                                        ORDER BY  [ID]
                                    ", SQLUERY.ToString());
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView6.DataSource = ds1.Tables["ds1"];

                    dataGridView6.AutoResizeColumns();

                }
                else
                {
                    dataGridView6.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB6(
                            string KINDS
                            , string IANUMERS
                            , string REGISTERNO
                            , string MANUNAMES
                            , string ADDRESS
                            , string CHECKS
                            , string NAMES
                            , string ORIS
                            , string MANUS
                            , string PROALLGENS
                            , string MANUALLGENS
                            , string PRIMES
                            , string COLORS
                            , string TASTES
                            , string CHARS
                            , string PACKAGES
                            , string WEIGHTS
                            , string SPECS
                            , string SAVEDAYS
                            , string SAVECONDITIONS
                            , string COMMEMTS
                            , string DOCNAMES1
                            , string CONTENTTYPES1
                            , byte[] DATAS1
                            , string DOCNAMES2
                            , string CONTENTTYPES2
                            , byte[] DATAS2                       

                             )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"   
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB6]
                                        (
                                        [KINDS]
                                        ,[IANUMERS]
                                        ,[REGISTERNO]
                                        ,[MANUNAMES]
                                        ,[ADDRESS]
                                        ,[CHECKS]
                                        ,[NAMES]
                                        ,[ORIS]
                                        ,[MANUS]
                                        ,[PROALLGENS]
                                        ,[MANUALLGENS]
                                        ,[PRIMES]
                                        ,[COLORS]
                                        ,[TASTES]
                                        ,[CHARS]
                                        ,[PACKAGES]
                                        ,[WEIGHTS]
                                        ,[SPECS]
                                        ,[SAVEDAYS]
                                        ,[SAVECONDITIONS]
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        )
                                        VALUES
                                        (
                                        @KINDS
                                        ,@IANUMERS
                                        ,@REGISTERNO
                                        ,@MANUNAMES
                                        ,@ADDRESS
                                        ,@CHECKS
                                        ,@NAMES
                                        ,@ORIS
                                        ,@MANUS
                                        ,@PROALLGENS
                                        ,@MANUALLGENS
                                        ,@PRIMES
                                        ,@COLORS
                                        ,@TASTES
                                        ,@CHARS
                                        ,@PACKAGES
                                        ,@WEIGHTS
                                        ,@SPECS
                                        ,@SAVEDAYS
                                        ,@SAVECONDITIONS
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        )
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@KINDS", KINDS);
                            cmd.Parameters.AddWithValue("@IANUMERS", IANUMERS);
                            cmd.Parameters.AddWithValue("@REGISTERNO", REGISTERNO);
                            cmd.Parameters.AddWithValue("@MANUNAMES", MANUNAMES);
                            cmd.Parameters.AddWithValue("@ADDRESS", ADDRESS);
                            cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@ORIS", ORIS);
                            cmd.Parameters.AddWithValue("@MANUS", MANUS);
                            cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                            cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                            cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                            cmd.Parameters.AddWithValue("@COLORS", COLORS);
                            cmd.Parameters.AddWithValue("@TASTES", TASTES);
                            cmd.Parameters.AddWithValue("@CHARS", CHARS);
                            cmd.Parameters.AddWithValue("@PACKAGES", PACKAGES);
                            cmd.Parameters.AddWithValue("@WEIGHTS", WEIGHTS);
                            cmd.Parameters.AddWithValue("@SPECS", SPECS);
                            cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                            cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);



                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }


            


        }

        public void UPDATE_TO_TBDB6(
                                string ID
                               , string KINDS
                                , string IANUMERS
                                , string REGISTERNO
                                , string MANUNAMES
                                , string ADDRESS
                                , string CHECKS
                                , string NAMES
                                , string ORIS
                                , string MANUS
                                , string PROALLGENS
                                , string MANUALLGENS
                                , string PRIMES
                                , string COLORS
                                , string TASTES
                                , string CHARS
                                , string PACKAGES
                                , string WEIGHTS
                                , string SPECS
                                , string SAVEDAYS
                                , string SAVECONDITIONS
                                , string COMMEMTS

                               )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB6]
                                        SET
                                        [KINDS]=@KINDS
                                        ,[IANUMERS]=@IANUMERS
                                        ,[REGISTERNO]=@REGISTERNO
                                        ,[MANUNAMES]=@MANUNAMES
                                        ,[ADDRESS]=@ADDRESS
                                        ,[CHECKS]=@CHECKS
                                        ,[NAMES]=@NAMES
                                        ,[ORIS]=@ORIS
                                        ,[MANUS]=@MANUS
                                        ,[PROALLGENS]=@PROALLGENS
                                        ,[MANUALLGENS]=@MANUALLGENS
                                        ,[PRIMES]=@PRIMES
                                        ,[COLORS]=@COLORS
                                        ,[TASTES]=@TASTES
                                        ,[CHARS]=@CHARS
                                        ,[PACKAGES]=@PACKAGES
                                        ,[WEIGHTS]=@WEIGHTS
                                        ,[SPECS]=@SPECS
                                        ,[SAVEDAYS]=@SAVEDAYS
                                        ,[SAVECONDITIONS]=@SAVECONDITIONS
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@IANUMERS", IANUMERS);
                        cmd.Parameters.AddWithValue("@REGISTERNO", REGISTERNO);
                        cmd.Parameters.AddWithValue("@MANUNAMES", MANUNAMES);
                        cmd.Parameters.AddWithValue("@ADDRESS", ADDRESS);
                        cmd.Parameters.AddWithValue("@CHECKS", CHECKS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@ORIS", ORIS);
                        cmd.Parameters.AddWithValue("@MANUS", MANUS);
                        cmd.Parameters.AddWithValue("@PROALLGENS", PROALLGENS);
                        cmd.Parameters.AddWithValue("@MANUALLGENS", MANUALLGENS);
                        cmd.Parameters.AddWithValue("@PRIMES", PRIMES);
                        cmd.Parameters.AddWithValue("@COLORS", COLORS);
                        cmd.Parameters.AddWithValue("@TASTES", TASTES);
                        cmd.Parameters.AddWithValue("@CHARS", CHARS);
                        cmd.Parameters.AddWithValue("@PACKAGES", PACKAGES);
                        cmd.Parameters.AddWithValue("@WEIGHTS", WEIGHTS);
                        cmd.Parameters.AddWithValue("@SPECS", SPECS);
                        cmd.Parameters.AddWithValue("@SAVEDAYS", SAVEDAYS);
                        cmd.Parameters.AddWithValue("@SAVECONDITIONS", SAVECONDITIONS);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void UPDATE_TO_TBDB6_ATTACHS(
                            string ID                           
                            , string DOCNAMES1
                            , string CONTENTTYPES1
                            , byte[] DATAS1
                            , string DOCNAMES2
                            , string CONTENTTYPES2
                            , byte[] DATAS2

                             )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(ID))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"   
                                        UPDATE  [TKRESEARCH].[dbo].[TBDB6]
                                        SET [DOCNAMES1]=@DOCNAMES1
                                        ,[CONTENTTYPES1]=@CONTENTTYPES1
                                        ,[DATAS1]=@DATAS1
                                        ,[DOCNAMES2]=@DOCNAMES2
                                        ,[CONTENTTYPES2]=@CONTENTTYPES2
                                        ,[DATAS2]=@DATAS2
                                        WHERE [ID]=@ID
                                       
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@ID", ID);                        

                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);



                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }





        }
        public void DELETE_TO_TBDB6(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB6]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            textBox6B.Text = null;
            textBox6C.Text = null;
            textBox631.Text = null;
            textBox632.Text = null;
            textBox633.Text = null;
            textBox634.Text = null;
            textBox635.Text = null;
            textBox636.Text = null;
            textBox637.Text = null;
            textBox638.Text = null;
            textBox639.Text = null;
            textBox640.Text = null;
            textBox641.Text = null;
            textBox642.Text = null;
            textBox643.Text = null;
            textBox644.Text = null;
            textBox645.Text = null;
            textBox646.Text = null;
            textBox647.Text = null;
            textBox648.Text = null;
            textBox649.Text = null;
            textBox650.Text = null;
            textBox661.Text = null;
            textBox6641.Text = null;
            textBox6642.Text = null;
            textBox6643.Text = null;
            textBox6644.Text = null;

            textBox6E.Text = null;
            textBox6511.Text = null;

            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    textBox6B.Text = row.Cells["ID"].Value.ToString();
                    textBox6C.Text = row.Cells["ID"].Value.ToString();
                    
                    comboBox6.Text = row.Cells["分類"].Value.ToString();
                    textBox631.Text = row.Cells["國際條碼"].Value.ToString();
                    textBox632.Text = row.Cells["食品業者登錄字號"].Value.ToString();
                    textBox633.Text = row.Cells["製造商名稱"].Value.ToString();
                    textBox634.Text = row.Cells["製造商地址"].Value.ToString();
                    textBox635.Text = row.Cells["品質認證"].Value.ToString();
                    textBox636.Text = row.Cells["產品品名"].Value.ToString();
                    textBox637.Text = row.Cells["產品成分"].Value.ToString();
                    textBox638.Text = row.Cells["製造流程"].Value.ToString();
                    textBox639.Text = row.Cells["產品過敏原"].Value.ToString();
                    textBox640.Text = row.Cells["產線及生產設備過敏原"].Value.ToString();
                    textBox641.Text = row.Cells["素別"].Value.ToString();
                    textBox642.Text = row.Cells["色澤"].Value.ToString();
                    textBox643.Text = row.Cells["風味"].Value.ToString();
                    textBox644.Text = row.Cells["性狀"].Value.ToString();
                    textBox645.Text = row.Cells["材質"].Value.ToString();
                    textBox646.Text = row.Cells["淨重量"].Value.ToString();
                    textBox647.Text = row.Cells["規格"].Value.ToString();
                    textBox648.Text = row.Cells["保存期限"].Value.ToString();
                    textBox649.Text = row.Cells["保存條件"].Value.ToString();
                    textBox650.Text = row.Cells["備註"].Value.ToString();
                    textBox661.Text = row.Cells["產品品名"].Value.ToString();

                    textBox6641.Text = row.Cells["ID"].Value.ToString();
                    textBox6642.Text = row.Cells["產品品名"].Value.ToString();

                    textBox6E.Text = row.Cells["ID"].Value.ToString();
                    textBox6511.Text = row.Cells["產品品名"].Value.ToString();
                }
                else
                {

                    textBox6B.Text = null;
                    textBox6C.Text = null;
                    textBox631.Text = null;
                    textBox632.Text = null;
                    textBox633.Text = null;
                    textBox634.Text = null;
                    textBox635.Text = null;
                    textBox636.Text = null;
                    textBox637.Text = null;
                    textBox638.Text = null;
                    textBox639.Text = null;
                    textBox640.Text = null;
                    textBox641.Text = null;
                    textBox642.Text = null;
                    textBox643.Text = null;
                    textBox644.Text = null;
                    textBox645.Text = null;
                    textBox646.Text = null;
                    textBox647.Text = null;
                    textBox648.Text = null;
                    textBox649.Text = null;
                    textBox650.Text = null;
                    textBox661.Text = null;

                    textBox6641.Text = null;
                    textBox6642.Text = null;
                    textBox6643.Text = null;
                    textBox6644.Text = null;

                    textBox6E.Text = null;
                    textBox6511.Text = null;
                }

            }
        }

        public void  SEARCH7(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                       SELECT 
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                        ,[DOCNAMES1]  AS '產前會議'
                                        ,[DOCNAMES2]  AS '技術移轉表'
                                        ,[DOCNAMES3] AS '產品圖片'

                                        FROM [TKRESEARCH].[dbo].[TBDB7]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                        [ID]
                                        ,[KINDS] AS '分類'
                                        ,[NAMES] AS '產品品名'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                        ,[DOCNAMES1]  AS '產前會議'
                                        ,[DOCNAMES2]  AS '技術移轉表'
                                        ,[DOCNAMES3] AS '產品圖片'

                                        FROM [TKRESEARCH].[dbo].[TBDB7]

                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView7.DataSource = ds1.Tables["ds1"];

                    dataGridView7.AutoResizeColumns();

                }
                else
                {
                    dataGridView7.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB7(
                          string KINDS                         
                          , string NAMES
                          , string COMMEMTS
                          , string DOCNAMES1
                          , string CONTENTTYPES1
                          , byte[] DATAS1
                          , string DOCNAMES2
                          , string CONTENTTYPES2
                          , byte[] DATAS2
                          , string DOCNAMES3
                          , string CONTENTTYPES3
                          , byte[] DATAS3

                           )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"   
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB7]
                                        (
                                        [KINDS]                                       
                                        ,[NAMES]                                       
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        )
                                        VALUES
                                        (
                                        @KINDS                                       
                                        ,@NAMES                                       
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        ,@DOCNAMES3
                                        ,@CONTENTTYPES3
                                        ,@DATAS3
                                        )
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@KINDS", KINDS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                            cmd.Parameters.AddWithValue("@DOCNAMES3", DOCNAMES3);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES3", CONTENTTYPES3);
                            cmd.Parameters.AddWithValue("@DATAS3", DATAS3);


                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }

            


        }

        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            textBox7B.Text = null;
            textBox7C.Text = null;
            textBox721.Text = null;
            textBox722.Text = null;
            textBox731.Text = null;

            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    textBox7B.Text = row.Cells["ID"].Value.ToString();
                    textBox7C.Text = row.Cells["ID"].Value.ToString();

                    comboBox8.Text = row.Cells["分類"].Value.ToString();
                    textBox721.Text = row.Cells["產品品名"].Value.ToString();
                    textBox722.Text = row.Cells["備註"].Value.ToString();
                    textBox731.Text = row.Cells["產品品名"].Value.ToString();

                }
                else
                {
                    textBox7B.Text = null;
                    textBox7C.Text = null;
                    textBox721.Text = null;
                    textBox722.Text = null;
                    textBox731.Text = null;

                }
            }
        }

        public void UPDATE_TO_TBDB7(
                              string ID
                             , string KINDS                           
                              , string NAMES                            
                              , string COMMEMTS

                             )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB7]
                                        SET
                                        [KINDS]=@KINDS                                        
                                        ,[NAMES]=@NAMES                                      
                                        ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);                        
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);                      
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB7(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB7]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH8(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                       SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[NAMES] AS '產品品名'
                                        ,[RECORDS] AS '產品品評回饋'
                                        ,[REPORTS] AS '產品測試報告'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                        ,[DOCNAMES1] AS '研發紀錄表'
                                        ,[DOCNAMES2] AS '產品圖片'
                                        ,[DOCNAMES3] AS '產品測試報告'
                                        FROM [TKRESEARCH].[dbo].[TBDB8]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                         [ID]
                                        ,[KINDS] AS '分類'
                                        ,[NAMES] AS '產品品名'
                                        ,[RECORDS] AS '產品品評回饋'
                                        ,[REPORTS] AS '產品測試報告'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112)  AS '填表日期'
                                        ,[DOCNAMES1] AS '研發紀錄表'
                                        ,[DOCNAMES2] AS '產品圖片'
                                        ,[DOCNAMES3] AS '產品測試報告'
                                        FROM [TKRESEARCH].[dbo].[TBDB8]

                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView8.DataSource = ds1.Tables["ds1"];

                    dataGridView8.AutoResizeColumns();

                }
                else
                {
                    dataGridView8.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB8(
                          string KINDS
                          , string NAMES
                          , string RECORDS
                          , string REPORTS
                          , string COMMEMTS
                          , string DOCNAMES1
                          , string CONTENTTYPES1
                          , byte[] DATAS1
                          , string DOCNAMES2
                          , string CONTENTTYPES2
                          , byte[] DATAS2
                          , string DOCNAMES3
                          , string CONTENTTYPES3
                          , byte[] DATAS3

                           )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"   
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB8]
                                        (
                                        [KINDS]                                       
                                        ,[NAMES]   
                                        ,[RECORDS]   
                                        ,[REPORTS]                                       
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                        ,[DOCNAMES2]
                                        ,[CONTENTTYPES2]
                                        ,[DATAS2]
                                        ,[DOCNAMES3]
                                        ,[CONTENTTYPES3]
                                        ,[DATAS3]
                                        )
                                        VALUES
                                        (
                                        @KINDS                                       
                                        ,@NAMES    
                                        ,@RECORDS               
                                        ,@REPORTS                                                  
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                        ,@DOCNAMES2
                                        ,@CONTENTTYPES2
                                        ,@DATAS2
                                        ,@DOCNAMES3
                                        ,@CONTENTTYPES3
                                        ,@DATAS3
                                        )
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@KINDS", KINDS);
                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@RECORDS", RECORDS);
                            cmd.Parameters.AddWithValue("@REPORTS", REPORTS);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);
                            cmd.Parameters.AddWithValue("@DOCNAMES2", DOCNAMES2);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES2", CONTENTTYPES2);
                            cmd.Parameters.AddWithValue("@DATAS2", DATAS2);
                            cmd.Parameters.AddWithValue("@DOCNAMES3", DOCNAMES3);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES3", CONTENTTYPES3);
                            cmd.Parameters.AddWithValue("@DATAS3", DATAS3);


                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }
          


        }

        public void UPDATE_TO_TBDB8(
                              string ID
                             , string KINDS
                              , string NAMES
                             , string RECORDS
                             , string REPORTS
                              , string COMMEMTS

                             )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB8]
                                        SET
                                        [KINDS]=@KINDS                                        
                                        ,[NAMES]=@NAMES                                      
                                        ,[RECORDS]=@RECORDS
                                       ,[REPORTS]=@REPORTS
                                       ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@KINDS", KINDS);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@RECORDS", RECORDS);
                        cmd.Parameters.AddWithValue("@REPORTS", REPORTS);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB8(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB8]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }
        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            textBox8B.Text = null;
            textBox8C.Text = null;
            textBox821.Text = null;
            textBox822.Text = null;
            textBox823.Text = null;
            textBox824.Text = null;
            textBox831.Text = null;

            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    

                    textBox8B.Text = row.Cells["ID"].Value.ToString();
                    textBox8C.Text = row.Cells["ID"].Value.ToString();

                    comboBox10.Text= row.Cells["分類"].Value.ToString();
                    textBox821.Text = row.Cells["產品品名"].Value.ToString();
                    textBox822.Text = row.Cells["產品品評回饋"].Value.ToString();
                    textBox823.Text = row.Cells["產品測試報告"].Value.ToString();
                    textBox824.Text = row.Cells["備註"].Value.ToString();
                    textBox831.Text = row.Cells["產品品名"].Value.ToString();

                }
                else
                {
                    textBox8B.Text = null;
                    textBox8C.Text = null;
                    textBox821.Text = null;
                    textBox822.Text = null;
                    textBox823.Text = null;
                    textBox824.Text = null;
                    textBox831.Text = null;
                }
            }
        }

        public void SEARCH9(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                        [ID]
                                        ,[NAMES] AS '名稱'
                                        ,[CONTENTS] AS '文件內容敘述'
                                        ,[COMMEMTS] AS '備註'
                                        ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '檔案'

                                        FROM [TKRESEARCH].[dbo].[TBDB9]

                                        WHERE NAMES LIKE '%{0}%'
                                        ORDER BY  [ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                         SELECT
                                         [ID]
                                        ,[NAMES] AS '名稱'
                                        ,[CONTENTS] AS '文件內容敘述'
                                        ,[COMMEMTS] AS '備註'
                                         ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                        ,[DOCNAMES1] AS '檔案'

                                        FROM [TKRESEARCH].[dbo].[TBDB9]

                                        ORDER BY  [ID]
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView9.DataSource = ds1.Tables["ds1"];

                    dataGridView9.AutoResizeColumns();

                }
                else
                {
                    dataGridView9.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBDB9(
                         string NAMES
                         , string CONTENTS                       
                         , string COMMEMTS
                         , string DOCNAMES1
                         , string CONTENTTYPES1
                         , byte[] DATAS1                         

                          )
        {

            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(NAMES))
                    {
                        StringBuilder ADDSQL = new StringBuilder();
                        ADDSQL.AppendFormat(@"   
                                        INSERT INTO [TKRESEARCH].[dbo].[TBDB9]
                                        (
                                       [NAMES]   
                                        ,[CONTENTS]                                                                          
                                        ,[COMMEMTS]
                                        ,[DOCNAMES1]
                                        ,[CONTENTTYPES1]
                                        ,[DATAS1]
                                     
                                        )
                                        VALUES
                                        (
                                       @NAMES    
                                        ,@CONTENTS                                  
                                        ,@COMMEMTS
                                        ,@DOCNAMES1
                                        ,@CONTENTTYPES1
                                        ,@DATAS1
                                       
                                        )
                                       
                                        ");

                        string sql = ADDSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {


                            cmd.Parameters.AddWithValue("@NAMES", NAMES);
                            cmd.Parameters.AddWithValue("@CONTENTS", CONTENTS);
                            cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                            cmd.Parameters.AddWithValue("@DOCNAMES1", DOCNAMES1);
                            cmd.Parameters.AddWithValue("@CONTENTTYPES1", CONTENTTYPES1);
                            cmd.Parameters.AddWithValue("@DATAS1", DATAS1);



                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("新增失敗");
            }

           


        }

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            textBox9B.Text = null;
            textBox9C.Text = null;
            textBox921.Text = null;
            textBox922.Text = null;
            textBox923.Text = null;       
            textBox931.Text = null;

            if (dataGridView9.CurrentRow != null)
            {
                int rowindex = dataGridView9.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView9.Rows[rowindex];
                    textBox9B.Text = row.Cells["ID"].Value.ToString();
                    textBox9C.Text = row.Cells["ID"].Value.ToString();
                    textBox921.Text = row.Cells["名稱"].Value.ToString();
                    textBox922.Text = row.Cells["文件內容敘述"].Value.ToString();
                    textBox923.Text = row.Cells["備註"].Value.ToString();
                    textBox931.Text = row.Cells["名稱"].Value.ToString();

                }
                else
                {
                    textBox9B.Text = null;
                    textBox9C.Text = null;
                    textBox921.Text = null;
                    textBox922.Text = null;
                    textBox923.Text = null;
                    textBox931.Text = null;
                }
            }
        }

        public void UPDATE_TO_TBDB9(
                             string ID
                            , string NAMES
                             , string CONTENTS
                             , string COMMEMTS

                            )
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(NAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        UPDATE [TKRESEARCH].[dbo].[TBDB9]
                                        SET
                                        [NAMES]=@NAMES                                      
                                        ,[CONTENTS]=@CONTENTS                                     
                                       ,[COMMEMTS]=@COMMEMTS
                                        WHERE [ID]=@ID
                                       
                                       
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);                 
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@CONTENTS", CONTENTS);
                        cmd.Parameters.AddWithValue("@COMMEMTS", COMMEMTS);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }


        }

        public void DELETE_TO_TBDB9(string ID)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(ID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        DELETE[TKRESEARCH].[dbo].[TBDB9]
                                        WHERE ID=@ID
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SEARCH10(string KEYS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

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

                if (!string.IsNullOrEmpty(KEYS))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT 
                                        [ID]
                                        ,[KIND] AS '分類'
                                        ,[PARAID] AS '名稱'
                                        ,[PARANAME] AS '內容'
                                        FROM [TKRESEARCH].[dbo].[TBPARA]
                                        WHERE [KIND] IN ('原料','成品','物料')                                  

                                        AND  PARAID LIKE '%{0}%'

                                        ORDER BY [KIND],[ID]
                                    ", KEYS);
                }
                else
                {
                    sbSql.AppendFormat(@"                                          
                                        SELECT 
                                        [ID]
                                        ,[KIND] AS '分類'
                                        ,[PARAID] AS '名稱'
                                        ,[PARANAME] AS '內容'
                                        FROM [TKRESEARCH].[dbo].[TBPARA]
                                        WHERE [KIND] IN ('原料','成品','物料')

                                        ORDER BY [KIND],[ID]

                                       
                                    ");
                }


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    dataGridView10.DataSource = ds1.Tables["ds1"];

                    dataGridView10.AutoResizeColumns();

                }
                else
                {
                    dataGridView10.DataSource = null;

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        public void ADD_TO_TBPARA(string KIND, string PARAID, string PARANAME)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(PARAID))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        INSERT INTO  [TKRESEARCH].[dbo].[TBPARA]
                                        (
                                        [KIND] 
                                        ,[PARAID] 
                                        ,[PARANAME] 
                                        )
                                        VALUES
                                        (
                                        @KIND
                                        ,@PARAID
                                        ,@PARANAME
                                        )
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@KIND", KIND);
                        cmd.Parameters.AddWithValue("@PARAID", PARAID);
                        cmd.Parameters.AddWithValue("@PARANAME", PARANAME);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void ADD_TO_TBDBLOG(string USERNAMES,string DBNAMES, string DOCID,string NAMES, string ACTION, string ATTACHNAMES)
        {
            // 20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            using (SqlConnection conn = sqlConn)
            {
                if (!string.IsNullOrEmpty(USERNAMES))
                {
                    StringBuilder ADDSQL = new StringBuilder();
                    ADDSQL.AppendFormat(@"
                                        INSERT INTO  [TKRESEARCH].[dbo].[TBDBLOG]
                                        (
                                        [USERNAMES]
                                        ,[DBNAMES]
                                        ,[DOCID]
                                        ,[NAMES]
                                        ,[ACTION]
                                        ,[ATTACHNAMES]
                                        )
                                        VALUES
                                        (
                                        @USERNAMES
                                        ,@DBNAMES
                                        ,@DOCID
                                        ,@NAMES
                                        ,@ACTION
                                        ,@ATTACHNAMES
                                        )
                                        
                                        ");

                    string sql = ADDSQL.ToString();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@USERNAMES", USERNAMES);
                        cmd.Parameters.AddWithValue("@DBNAMES", DBNAMES);
                        cmd.Parameters.AddWithValue("@DOCID", DOCID);
                        cmd.Parameters.AddWithValue("@NAMES", NAMES);
                        cmd.Parameters.AddWithValue("@ACTION", ACTION);
                        cmd.Parameters.AddWithValue("@ATTACHNAMES", ATTACHNAMES);

                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }

            }
        }

        public void SETFASTREPORT(string ID)
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

            SQL1 = SETSQL(ID);
            Report report1 = new Report();
            report1.Load(@"REPORT\產品規格表v2.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string ID)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                SELECT 
                                [ID] 
                                ,[KINDS] AS '分類'
                                ,[IANUMERS] AS '國際條碼'
                                ,[REGISTERNO] AS '食品業者登錄字號'
                                ,[MANUNAMES] AS '製造商名稱'
                                ,[ADDRESS] AS '製造商地址'
                                ,[CHECKS] AS '品質認證'
                                ,[NAMES] AS '產品品名'
                                ,[ORIS] AS '產品成分'
                                ,[MANUS] AS '製造流程'
                                ,[PROALLGENS] AS '產品過敏原'
                                ,[MANUALLGENS] AS '產線及生產設備過敏原'
                                ,[PRIMES] AS '素別'
                                ,[COLORS] AS '色澤'
                                ,[TASTES] AS '風味'
                                ,[CHARS] AS '性狀'
                                ,[PACKAGES] AS '材質'
                                ,[WEIGHTS] AS '淨重量'
                                ,[SPECS] AS '規格'
                                ,[SAVEDAYS] AS '保存期限'
                                ,[SAVECONDITIONS] AS '保存條件'
                                ,[COMMEMTS] AS '備註'
                                ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '填表日期'
                                ,[DOCNAMES1] AS '營養標示'
                                    
                                ,[DOCNAMES2] AS '產品圖片'
                                ,[DATAS1]
                                ,[DATAS2]
                                       
                                FROM [TKRESEARCH].[dbo].[TBDB6]
                                WHERE [ID]='{0}'
                                ORDER BY  [ID]
                           
                           
                            ", ID);

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox1A.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OPEN1();

        }
        private void button4_Click(object sender, EventArgs e)
        {
            string DOCID = textBox11.Text;
            string COMMENTS = textBox12.Text;
            string DOCNAMES = "";
            string CONTENTTYPES = "";
            byte[] DATAS = new byte[] { 1 };
        

            if (!string.IsNullOrEmpty(DOCNAMES1))
            {
                DOCNAMES = DOCNAMES1;
                CONTENTTYPES = CONTENTTYPES1;
                DATAS = BYTES1;
            }
         

            ADD_TO_TBDB1(DOCID, COMMENTS, DOCNAMES, CONTENTTYPES, DATAS);
            SEARCH(textBox1A.Text.Trim());

            ADD_TO_TBDBLOG(shareArea.UserName, "DB1", "", DOCID, "ADD", DOCNAMES);
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATE_TO_TBDB1(textBox1B.Text,textBox14.Text,textBox15.Text);
            SEARCH(textBox1A.Text.Trim());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB1", textBox1C.Text,textBox131.Text, "DELETE", "");

                DELETE_TO_TBDB1(textBox1C.Text);
                SEARCH(textBox1A.Text.Trim());

            
               
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH2(textBox2A.Text.Trim());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OPEN2();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string NAMES = textBox201.Text;
            string CHARS = textBox202.Text;
            string ORIS = textBox203.Text;
            string SPECS = textBox204.Text;
            string PRICES = textBox205.Text;
            string SAVEDAYS = textBox206.Text;
            string SAVEMETHODS = textBox207.Text;
            string PRIMES = textBox208.Text; 
            string ALLERGENS = textBox209.Text;
            string OWNERS = textBox210.Text;
            string EATES = textBox211.Text;
            string CHECKSUNITS = textBox212.Text;
            string CHECKS = textBox213.Text;
            string OTHERS = textBox214.Text;
            string COMMENTS = textBox215.Text;

            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };

            if (!string.IsNullOrEmpty(DOCNAMES2))
            {
                DOCNAMES1 = DOCNAMES2;
                CONTENTTYPES1 = CONTENTTYPES2;
                DATAS1 = BYTES2;
            }


            ADD_TO_TBDB2( NAMES
                                ,  CHARS
                                ,  ORIS
                                ,  SPECS
                                ,  PRICES
                                ,  SAVEDAYS
                                ,  SAVEMETHODS
                                ,  PRIMES
                                ,  ALLERGENS
                                ,  OWNERS
                                ,  EATES
                                ,  CHECKSUNITS
                                ,  CHECKS
                                ,  OTHERS
                                ,  COMMENTS
                                ,  DOCNAMES1
                                ,  CONTENTTYPES1
                                ,  DATAS1
                                );
            SEARCH2(textBox2A.Text.Trim());

            ADD_TO_TBDBLOG(shareArea.UserName, "DB2", "", NAMES, "ADD", DOCNAMES1);
        }
        private void button11_Click(object sender, EventArgs e)
        {
            string ID = textBox2B.Text;
            string NAMES = textBox221.Text;
            string CHARS = textBox222.Text;
            string ORIS = textBox223.Text;
            string SPECS = textBox224.Text;
            string PRICES = textBox225.Text;
            string SAVEDAYS = textBox226.Text;
            string SAVEMETHODS = textBox227.Text;
            string PRIMES = textBox228.Text;
            string ALLERGENS = textBox229.Text;
            string OWNERS = textBox230.Text;
            string EATES = textBox231.Text;
            string CHECKSUNITS = textBox232.Text;
            string CHECKS = textBox233.Text;
            string OTHERS = textBox234.Text;
            string COMMENTS = textBox235.Text;

            UPDATE_TO_TBDB2(ID, NAMES, CHARS, ORIS, SPECS, PRICES, SAVEDAYS, SAVEMETHODS, PRIMES, ALLERGENS, OWNERS, EATES, CHECKSUNITS, CHECKS, OTHERS, COMMENTS);
            SEARCH2(textBox2A.Text.Trim());
        }
        private void button10_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB2", textBox2C.Text,textBox241.Text, "DELETE", "");

                DELETE_TO_TBDB2(textBox2C.Text);
                SEARCH2(textBox2A.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SEARCH3(textBox3A.Text);
        }
        private void button13_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox1.Text.ToString();
            string SUPPLYS = textBox301.Text.ToString();
            string NAMES = textBox302.Text.ToString();
            string ORIS = textBox303.Text.ToString();
            string SPECS = textBox304.Text.ToString();
            string PROALLGENS = textBox305.Text.ToString();
            string MANUALLGENS = textBox306.Text.ToString();
            string PLACES = textBox307.Text.ToString();
            string OUTS = textBox308.Text.ToString();
            string COLORS = textBox309.Text.ToString();
            string TASTES = textBox310.Text.ToString();
            string LOTS = textBox311.Text.ToString();
            string CHECKS = textBox312.Text.ToString();
            string SAVEDAYS = textBox313.Text.ToString();
            string SAVECONDITIONS = textBox314.Text.ToString();
            string BASEONS = textBox315.Text.ToString();
            string COA = textBox315.Text.ToString();
            string INCHECKRATES = textBox317.Text.ToString();
            string RULES = textBox318.Text.ToString();
            string COMMEMTS = textBox319.Text.ToString();
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1= new byte[] {1};
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
            string DOCNAMES3 = "";
            string CONTENTTYPES3 = "";
            byte[] DATAS3 = new byte[] { 1 };
            string DOCNAMES4 = "";
            string CONTENTTYPES4 = "";
            byte[] DATAS4 = new byte[] { 1 };
            string DOCNAMES5 = "";
            string CONTENTTYPES5 = "";
            byte[] DATAS5 = new byte[] { 1 };
            string DOCNAMES6 = "";
            string CONTENTTYPES6 = "";
            byte[] DATAS6 = new byte[] { 1 };

            if (!string.IsNullOrEmpty(DOCNAMES31))
            {
                DOCNAMES1 = DOCNAMES31;
                CONTENTTYPES1 = CONTENTTYPES31;
                DATAS1 = BYTES31;
            }
            if (!string.IsNullOrEmpty(DOCNAMES32))
            {
                DOCNAMES2 = DOCNAMES32;
                CONTENTTYPES2 = CONTENTTYPES32;
                DATAS2 = BYTES32;
            }
            if (!string.IsNullOrEmpty(DOCNAMES33))
            {
                DOCNAMES3 = DOCNAMES33;
                CONTENTTYPES3 = CONTENTTYPES33;
                DATAS3 = BYTES33;
            }
            if (!string.IsNullOrEmpty(DOCNAMES34))
            {
                DOCNAMES4 = DOCNAMES34;
                CONTENTTYPES4 = CONTENTTYPES34;
                DATAS4 = BYTES34;
            }
            if (!string.IsNullOrEmpty(DOCNAMES35))
            {
                DOCNAMES5 = DOCNAMES35;
                CONTENTTYPES5 = CONTENTTYPES35;
                DATAS5 = BYTES35;
            }
            if (!string.IsNullOrEmpty(DOCNAMES36))
            {
                DOCNAMES6 = DOCNAMES36;
                CONTENTTYPES6 = CONTENTTYPES36;
                DATAS6 = BYTES36;
            }


      

            ADD_TO_TBDB3(
                        KINDS
                        , SUPPLYS
                        , NAMES
                        , ORIS
                        , SPECS
                        , PROALLGENS
                        , MANUALLGENS
                        , PLACES
                        , OUTS
                        , COLORS
                        , TASTES
                        , LOTS
                        , CHECKS
                        , SAVEDAYS
                        , SAVECONDITIONS
                        , BASEONS
                        , COA
                        , INCHECKRATES
                        , RULES
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        , DOCNAMES3
                        , CONTENTTYPES3
                        , DATAS3
                        , DOCNAMES4
                        , CONTENTTYPES4
                        , DATAS4
                        , DOCNAMES5
                        , CONTENTTYPES5
                        , DATAS5
                        , DOCNAMES6
                        , CONTENTTYPES6
                        , DATAS6
                        );

            SEARCH3(textBox3A.Text);


            ADD_TO_TBDBLOG(shareArea.UserName, "DB3", "", NAMES, "ADD", DOCNAMES1+","+ DOCNAMES2+ "," + DOCNAMES3+ "," + DOCNAMES4+ "," + DOCNAMES5+ "," + DOCNAMES6);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OPEN31();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            OPEN32();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            OPEN33();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            OPEN34();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            OPEN35();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            OPEN36();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string ID = textBox3B.Text;
            string KINDS = comboBox2.Text;

            string SUPPLYS = textBox331.Text;
            string NAMES = textBox332.Text;
            string ORIS = textBox333.Text;
            string SPECS = textBox334.Text;
            string PROALLGENS = textBox335.Text;
            string MANUALLGENS = textBox336.Text;
            string PLACES = textBox337.Text;
            string OUTS = textBox338.Text;
            string COLORS = textBox339.Text;
            string TASTES = textBox340.Text;
            string LOTS = textBox341.Text;
            string CHECKS = textBox342.Text;
            string SAVEDAYS = textBox343.Text;
            string SAVECONDITIONS = textBox344.Text;
            string BASEONS = textBox345.Text;
            string COA = textBox346.Text;
            string INCHECKRATES = textBox347.Text;
            string RULES = textBox348.Text;
            string COMMEMTS = textBox349.Text;

            UPDATE_TO_TBDB3(
                                ID
                              , KINDS
                              , SUPPLYS
                              , NAMES
                              , ORIS
                              , SPECS
                              , PROALLGENS
                              , MANUALLGENS
                              , PLACES
                              , OUTS
                              , COLORS
                              , TASTES
                              , LOTS
                              , CHECKS
                              , SAVEDAYS
                              , SAVECONDITIONS
                              , BASEONS
                              , COA
                              , INCHECKRATES
                              , RULES
                              , COMMEMTS

                              );

            SEARCH3(textBox3A.Text.Trim());
        }
        private void button15_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB3", textBox3C.Text,textBox350.Text, "DELETE", "");

                DELETE_TO_TBDB3(textBox3C.Text);
                SEARCH3(textBox3A.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button21_Click(object sender, EventArgs e)
        {
            SEARCH4(textBox4A.Text.Trim());
        }

        private void button22_Click(object sender, EventArgs e)
        {
            OPEN41();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            OPEN42();
        }
        private void button28_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox3.Text.ToString();
            string SUPPLYS = textBox401.Text.ToString();
            string NAMES = textBox402.Text.ToString();
            string SPECS = textBox403.Text.ToString();           
            string OUTS = textBox404.Text.ToString();
            string COLORS = textBox405.Text.ToString();           
            string CHECKS = textBox406.Text.ToString();
            string SAVEDAYS = textBox407.Text.ToString();          
            string COMMEMTS = textBox408.Text.ToString();
            string COA = textBox409.Text.ToString();
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
           

            if (!string.IsNullOrEmpty(DOCNAMES41))
            {
                DOCNAMES1 = DOCNAMES41;
                CONTENTTYPES1 = CONTENTTYPES41;
                DATAS1 = BYTES41;
            }
            if (!string.IsNullOrEmpty(DOCNAMES42))
            {
                DOCNAMES2 = DOCNAMES42;
                CONTENTTYPES2 = CONTENTTYPES42;
                DATAS2 = BYTES42;
            }
           




            ADD_TO_TBDB4(
                        KINDS
                        , SUPPLYS
                        , NAMES                     
                        , SPECS                     
                        , OUTS
                        , COLORS                      
                        , CHECKS
                        , SAVEDAYS                       
                        , COA                      
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                       
                        );

            SEARCH4(textBox4A.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB4", "", NAMES, "ADD", DOCNAMES1 + "," + DOCNAMES2);
        }
        private void button25_Click(object sender, EventArgs e)
        {
            string ID = textBox4B.Text;
            string KINDS = comboBox4.Text;
            string SUPPLYS = textBox420.Text;
            string NAMES = textBox421.Text;
            string SPECS = textBox422.Text;
            string OUTS = textBox423.Text;
            string COLORS = textBox424.Text;
            string CHECKS = textBox425.Text;
            string SAVEDAYS = textBox426.Text;
            string COA = textBox427.Text;
            string COMMEMTS = textBox428.Text;

            UPDATE_TO_TBDB4(
                                ID
                               , KINDS
                               , SUPPLYS
                               , NAMES
                               , SPECS
                               , OUTS
                               , COLORS
                               , CHECKS
                               , SAVEDAYS
                               , COA
                               , COMMEMTS
                               );

         SEARCH4(textBox4A.Text);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB4", textBox4C.Text,textBox430.Text, "DELETE", "");

                DELETE_TO_TBDB4(textBox4C.Text);
                SEARCH4(textBox4A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            SEARCH5(textBox5A.Text);
        }
        private void button30_Click(object sender, EventArgs e)
        {
         
            string SUPPLYS = textBox501.Text.ToString();
            string NAMES = textBox502.Text.ToString();          
            string COMMEMTS = textBox503.Text.ToString();
    
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
            string DOCNAMES3 = "";
            string CONTENTTYPES3 = "";
            byte[] DATAS3 = new byte[] { 1 };


            if (!string.IsNullOrEmpty(DOCNAMES51))
            {
                DOCNAMES1 = DOCNAMES51;
                CONTENTTYPES1 = CONTENTTYPES51;
                DATAS1 = BYTES51;
            }
            if (!string.IsNullOrEmpty(DOCNAMES52))
            {
                DOCNAMES2 = DOCNAMES52;
                CONTENTTYPES2 = CONTENTTYPES52;
                DATAS2 = BYTES52;
            }
            if (!string.IsNullOrEmpty(DOCNAMES53))
            {
                DOCNAMES3 = DOCNAMES53;
                CONTENTTYPES3 = CONTENTTYPES53;
                DATAS3 = BYTES53;
            }




            ADD_TO_TBDB5(
                        SUPPLYS
                        , NAMES
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        , DOCNAMES3
                        , CONTENTTYPES3
                        , DATAS3

                        );

            SEARCH5(textBox5A.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB5", "", NAMES, "ADD", DOCNAMES1 + "," + DOCNAMES2 + "," + DOCNAMES3);
        }
        private void button26_Click(object sender, EventArgs e)
        {
            OPEN51();
        }

        private void button29_Click(object sender, EventArgs e)
        {
            OPEN52();
        }

        private void button31_Click(object sender, EventArgs e)
        {
            OPEN53();
        }
        private void button33_Click(object sender, EventArgs e)
        {
            string ID = textBox5B.Text;
          
            string SUPPLYS = textBox521.Text;
            string NAMES = textBox522.Text;          
            string COMMEMTS = textBox523.Text;

            UPDATE_TO_TBDB5(
                                ID                             
                               , SUPPLYS
                               , NAMES                             
                               , COMMEMTS
                               );

            SEARCH5(textBox5A.Text);
        }
        private void button32_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB5", textBox5C.Text,textBox531.Text, "DELETE", "");

                DELETE_TO_TBDB5(textBox5C.Text);
                SEARCH5(textBox5A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button34_Click(object sender, EventArgs e)
        {
            SEARCH6(textBox6A.Text, comboBox12.Text);
        }


        private void button41_Click(object sender, EventArgs e)
        {

            string KINDS = comboBox5.Text.ToString();
            string IANUMERS = textBox601.Text.ToString();
            string REGISTERNO = textBox602.Text.ToString();
            string MANUNAMES = textBox603.Text.ToString();
            string ADDRESS = textBox604.Text.ToString();
            string CHECKS = textBox605.Text.ToString();
            string NAMES = textBox606.Text.ToString();
            string ORIS = textBox607.Text.ToString();
            string MANUS = textBox608.Text.ToString();
            string PROALLGENS = textBox609.Text.ToString();
            string MANUALLGENS = textBox610.Text.ToString();
            string PRIMES = textBox611.Text.ToString();
            string COLORS = textBox612.Text.ToString();
            string TASTES = textBox613.Text.ToString();
            string CHARS = textBox614.Text.ToString();
            string PACKAGES = textBox615.Text.ToString();
            string WEIGHTS = textBox616.Text.ToString();
            string SPECS = textBox617.Text.ToString();
            string SAVEDAYS = textBox618.Text.ToString();
            string SAVECONDITIONS = textBox622.Text.ToString();
            string COMMEMTS = textBox619.Text.ToString();

            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
           


            if (!string.IsNullOrEmpty(DOCNAMES61))
            {
                DOCNAMES1 = DOCNAMES61;
                CONTENTTYPES1 = CONTENTTYPES61;
                DATAS1 = BYTES61;
            }
            if (!string.IsNullOrEmpty(DOCNAMES62))
            {
                DOCNAMES2 = DOCNAMES62;
                CONTENTTYPES2 = CONTENTTYPES62;
                DATAS2 = BYTES62;
            }
          




            ADD_TO_TBDB6(
                        KINDS
                        , IANUMERS
                        , REGISTERNO
                        , MANUNAMES
                        , ADDRESS
                        , CHECKS
                        , NAMES
                        , ORIS
                        , MANUS
                        , PROALLGENS
                        , MANUALLGENS
                        , PRIMES
                        , COLORS
                        , TASTES
                        , CHARS
                        , PACKAGES
                        , WEIGHTS
                        , SPECS
                        , SAVEDAYS
                        , SAVECONDITIONS
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        );

            SEARCH6(textBox6A.Text, comboBox12.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB6", "", NAMES, "ADD", DOCNAMES1 + "," + DOCNAMES2 );
        }

        private void button35_Click(object sender, EventArgs e)
        {
            OPEN61();
        }

        private void button40_Click(object sender, EventArgs e)
        {
            OPEN62();
        }
        private void button38_Click(object sender, EventArgs e)
        {
            string ID = textBox6B.Text;

            string KINDS = comboBox6.Text;
            string IANUMERS = textBox631.Text;
            string REGISTERNO = textBox632.Text;
            string MANUNAMES = textBox633.Text;
            string ADDRESS = textBox634.Text;
            string CHECKS = textBox635.Text;
            string NAMES = textBox636.Text;
            string ORIS = textBox637.Text;
            string MANUS = textBox638.Text;
            string PROALLGENS = textBox639.Text;
            string MANUALLGENS = textBox640.Text;
            string PRIMES = textBox641.Text;
            string COLORS = textBox642.Text;
            string TASTES = textBox643.Text;
            string CHARS = textBox644.Text;
            string PACKAGES = textBox645.Text;
            string WEIGHTS = textBox646.Text;
            string SPECS = textBox647.Text;
            string SAVEDAYS = textBox648.Text;
            string SAVECONDITIONS = textBox649.Text;
            string COMMEMTS = textBox650.Text;
          

            UPDATE_TO_TBDB6(
                            ID
                            , KINDS
                            , IANUMERS
                            , REGISTERNO
                            , MANUNAMES
                            , ADDRESS
                            , CHECKS
                            , NAMES
                            , ORIS
                            , MANUS
                            , PROALLGENS
                            , MANUALLGENS
                            , PRIMES
                            , COLORS
                            , TASTES
                            , CHARS
                            , PACKAGES
                            , WEIGHTS
                            , SPECS
                            , SAVEDAYS
                            , SAVECONDITIONS
                            , COMMEMTS
                               );

            SEARCH6(textBox6A.Text, comboBox12.Text);
        }
        private void button43_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB6", textBox6C.Text,textBox661.Text, "DELETE", "");

                DELETE_TO_TBDB6(textBox6C.Text);
                SEARCH6(textBox6A.Text, comboBox12.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
          
        }

        private void button36_Click(object sender, EventArgs e)
        {
            SEARCH7(textBox7A.Text);
        }
        private void button42_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox7.Text.ToString();    
            string NAMES = textBox701.Text.ToString();      
            string COMMEMTS = textBox702.Text.ToString();

            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
            string DOCNAMES3 = "";
            string CONTENTTYPES3 = "";
            byte[] DATAS3 = new byte[] { 1 };



            if (!string.IsNullOrEmpty(DOCNAMES71))
            {
                DOCNAMES1 = DOCNAMES71;
                CONTENTTYPES1 = CONTENTTYPES71;
                DATAS1 = BYTES71;
            }
            if (!string.IsNullOrEmpty(DOCNAMES72))
            {
                DOCNAMES2 = DOCNAMES72;
                CONTENTTYPES2 = CONTENTTYPES72;
                DATAS2 = BYTES72;
            }
            if (!string.IsNullOrEmpty(DOCNAMES73))
            {
                DOCNAMES3 = DOCNAMES73;
                CONTENTTYPES3 = CONTENTTYPES73;
                DATAS3 = BYTES73;
            }
            

            ADD_TO_TBDB7(
                        KINDS                       
                        , NAMES
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        , DOCNAMES3
                        , CONTENTTYPES3
                        , DATAS3
                        );

            SEARCH7(textBox7A.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB7", "", NAMES, "ADD", DOCNAMES1 + "," + DOCNAMES2 + "," + DOCNAMES3);
        }

        private void button37_Click(object sender, EventArgs e)
        {
            OPEN71();
        }

        private void button39_Click(object sender, EventArgs e)
        {
            OPEN72();
        }

        private void button44_Click(object sender, EventArgs e)
        {
            OPEN73();
        }
        private void button48_Click(object sender, EventArgs e)
        {
            string ID = textBox7B.Text;
            string KINDS = comboBox8.Text.ToString();
            string NAMES = textBox721.Text.ToString();
            string COMMEMTS = textBox722.Text.ToString();
    


            UPDATE_TO_TBDB7(
                        ID
                        ,KINDS
                        , NAMES
                        , COMMEMTS                     
                        );

            SEARCH7(textBox7A.Text);
        }

        private void button45_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB7", textBox7C.Text,textBox731.Text, "DELETE", "");

                DELETE_TO_TBDB7(textBox7C.Text);
                SEARCH7(textBox7A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button46_Click(object sender, EventArgs e)
        {
            SEARCH8(textBox8A.Text);
        }

        private void button50_Click(object sender, EventArgs e)
        {
            string KINDS = comboBox8.Text.ToString();
            string NAMES = textBox801.Text.ToString();
            string RECORDS = textBox802.Text.ToString();
            string REPORTS = textBox803.Text.ToString();
            string COMMEMTS = textBox804.Text.ToString();

            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };
            string DOCNAMES3 = "";
            string CONTENTTYPES3 = "";
            byte[] DATAS3 = new byte[] { 1 };



            if (!string.IsNullOrEmpty(DOCNAMES81))
            {
                DOCNAMES1 = DOCNAMES81;
                CONTENTTYPES1 = CONTENTTYPES81;
                DATAS1 = BYTES81;
            }
            if (!string.IsNullOrEmpty(DOCNAMES82))
            {
                DOCNAMES2 = DOCNAMES82;
                CONTENTTYPES2 = CONTENTTYPES82;
                DATAS2 = BYTES82;
            }
            if (!string.IsNullOrEmpty(DOCNAMES83))
            {
                DOCNAMES3 = DOCNAMES83;
                CONTENTTYPES3 = CONTENTTYPES83;
                DATAS3 = BYTES83;
            }


            ADD_TO_TBDB8(
                        KINDS
                        , NAMES
                        , RECORDS
                        , REPORTS
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                        , DOCNAMES2
                        , CONTENTTYPES2
                        , DATAS2
                        , DOCNAMES3
                        , CONTENTTYPES3
                        , DATAS3
                        );


            SEARCH8(textBox8A.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB8", "", NAMES, "ADD", DOCNAMES1 + "," + DOCNAMES2 + "," + DOCNAMES3);
        }

        private void button47_Click(object sender, EventArgs e)
        {
            OPEN81();
        }

        private void button49_Click(object sender, EventArgs e)
        {
            OPEN82();
        }

        private void button51_Click(object sender, EventArgs e)
        {
            OPEN83();
        }
        private void button52_Click(object sender, EventArgs e)
        {
            string ID = textBox8B.Text;
            string KINDS = comboBox10.Text.ToString();
            string NAMES = textBox821.Text.ToString();
            string RECORDS = textBox822.Text.ToString();
            string REPORTS = textBox823.Text.ToString();
            string COMMEMTS = textBox824.Text.ToString();



            UPDATE_TO_TBDB8(
                        ID
                        , KINDS
                        , NAMES
                        , RECORDS
                        , REPORTS
                        , COMMEMTS
                        );

            SEARCH8(textBox8A.Text);
        }

        private void button53_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB8", textBox8C.Text,textBox831.Text, "DELETE", "");

                DELETE_TO_TBDB8(textBox8C.Text);
                SEARCH8(textBox8A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button54_Click(object sender, EventArgs e)
        {
            SEARCH9(textBox9A.Text);
        }
        private void button55_Click(object sender, EventArgs e)
        {            
            string NAMES = textBox901.Text.ToString();
            string CONTENTS = textBox902.Text.ToString();        
            string COMMEMTS = textBox903.Text.ToString();

            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
       



            if (!string.IsNullOrEmpty(DOCNAMES91))
            {
                DOCNAMES1 = DOCNAMES91;
                CONTENTTYPES1 = CONTENTTYPES91;
                DATAS1 = BYTES91;
            }
          


            ADD_TO_TBDB9(
                        NAMES
                        , CONTENTS                     
                        , COMMEMTS
                        , DOCNAMES1
                        , CONTENTTYPES1
                        , DATAS1
                       
                        );

            SEARCH9(textBox9A.Text);

            ADD_TO_TBDBLOG(shareArea.UserName, "DB9", "", NAMES, "ADD", DOCNAMES1);
        }

        private void button57_Click(object sender, EventArgs e)
        {
            OPEN91();
        }


        private void button56_Click(object sender, EventArgs e)
        {
            string ID = textBox9B.Text;
            string NAMES = textBox921.Text.ToString();
            string CONTENTS = textBox922.Text.ToString();           
            string COMMEMTS = textBox923.Text.ToString();



            UPDATE_TO_TBDB9(
                        ID
                        , NAMES
                        , CONTENTS
                        , COMMEMTS
                        );

            SEARCH9(textBox9A.Text);
        }

        private void button60_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ADD_TO_TBDBLOG(shareArea.UserName, "DB9", textBox9C.Text,textBox931.Text, "DELETE", "");

                DELETE_TO_TBDB9(textBox9C.Text);
                SEARCH9(textBox9A.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button58_Click(object sender, EventArgs e)
        {
            SEARCH10(textBox10A.Text);
        }
        private void button59_Click(object sender, EventArgs e)
        {
            string KIND = comboBox11.Text;
            string PARAID = textBoxA1.Text;
            string PARANAME = textBoxA2.Text;

            ADD_TO_TBPARA(KIND, PARAID, PARANAME);
            SEARCH10(textBox10A.Text);
        }
        private void button62_Click(object sender, EventArgs e)
        {
            OPEN641();
        }

        private void button61_Click(object sender, EventArgs e)
        {
            OPEN642();
        }

        private void button63_Click(object sender, EventArgs e)
        {
            string ID = textBox6641.Text;
            string DOCNAMES1 = "";
            string CONTENTTYPES1 = "";
            byte[] DATAS1 = new byte[] { 1 };
            string DOCNAMES2 = "";
            string CONTENTTYPES2 = "";
            byte[] DATAS2 = new byte[] { 1 };



            if (!string.IsNullOrEmpty(DOCNAMES641))
            {
                DOCNAMES1 = DOCNAMES641;
                CONTENTTYPES1 = CONTENTTYPES641;
                DATAS1 = BYTES641;
            }
            if (!string.IsNullOrEmpty(DOCNAMES642))
            {
                DOCNAMES2 = DOCNAMES642;
                CONTENTTYPES2 = CONTENTTYPES642;
                DATAS2 = BYTES642;
            }



            UPDATE_TO_TBDB6_ATTACHS(ID, DOCNAMES1, CONTENTTYPES1, DATAS1, DOCNAMES2, CONTENTTYPES2, DATAS2);

            SEARCH6(textBox6A.Text, comboBox12.Text);
        }
        private void button66_Click(object sender, EventArgs e)
        {
            SETFASTREPORT( textBox6E.Text);
        }


        #endregion


    }
}

