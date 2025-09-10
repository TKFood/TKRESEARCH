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
        string currentID = null;

        public FrmORIENTS_CHECKLISTS()
        {
            InitializeComponent();
        }
        private void FrmORIENTS_CHECKLISTS_Load(object sender, EventArgs e)
        {

        }

        #region FUNCTION
        public void SEARCH(string PRODUCTNAME)
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

                //if (!string.IsNullOrEmpty(ISCLOSED))
                //{
                //    QUERYS.AppendFormat(@"AND ISCLOSED='{0}'", ISCLOSED);
                //}

                //if (!string.IsNullOrEmpty(MB001))
                //{
                //    QUERYS2.AppendFormat(@"AND (MB001 LIKE '%{0}%' OR MB002 LIKE '%{0}%')", MB001);
                //}
                //else
                //{
                //    QUERYS2.AppendFormat(@"");
                //}

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
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
                                        ICON AS '圖示'
                                    FROM [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS]
                                    WHERE 1=1
                                    ORDER BY CATEGORY,SUPPLIER,PRODUCTNAME
                                    ", QUERYS.ToString(), QUERYS2.ToString());


                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds = new DataSet(); // 這樣就不需要再 Clear()
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH("");
        }
        #endregion


    }
}
