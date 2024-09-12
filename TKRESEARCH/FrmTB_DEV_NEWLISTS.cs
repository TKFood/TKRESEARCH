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
using System.Net.Mail;
using TKITDLL;
using System.Globalization;
using System.Collections;
using System.Xml;
using System.Xml.Linq;

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

            comboBox1load();
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
            Sequel.AppendFormat(@" 
                                SELECT 
                                [ID]
                                ,[NAMES]
                                FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS_SALES]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();


        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox5.Text = null;
            if (comboBox1.SelectedValue!=null && !string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
            {
                textBox5.Text = comboBox1.SelectedValue.ToString();
            }
        }
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
                                    ,[SALES] AS '業務'
                                    ,[COMMENTS] AS '註記'
                                    ,[INGREDIENTS] AS '差異特色'
                                    ,CONVERT(NVARCHAR,[GETDATES],112)  AS '打樣日期'
                                    ,[REPLY] AS '業務回覆'
                                    ,[SALESID] AS '業務ID'
                                    ,[COSTS] AS '成本'
                                    ,[MOQS] AS 'MOQ'
                                    ,[MANUPRODS] AS '一天產能量'
                                    ,CONVERT(NVARCHAR,[CARESTEDATES],112) AS '建立日期'
                                    ,[ID]
                                   ,(SELECT TOP 1 [DOC_NBR] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'DOC_NBR'
                                    ,(SELECT TOP 1 [F01FieldValue] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'F01FieldValue'
                                    ,(SELECT TOP 1 [F09FieldValue] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'F09FieldValue'
                                    ,(SELECT TOP 1 [COMMENT] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'COMMENT'
    
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SETTEXT();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                DataGridViewRow row = dataGridView1.Rows[rowindex];

                textBoxid.Text = row.Cells["ID"].Value.ToString();
                textBox2.Text = row.Cells["編號"].Value.ToString();
                textBox3.Text = row.Cells["商品"].Value.ToString();
                textBox4.Text = row.Cells["規格"].Value.ToString();
                //textBox5.Text = row.Cells["需求人"].Value.ToString();
                comboBox1.SelectedValue = row.Cells["業務ID"].Value.ToString();
                textBox5.Text = row.Cells["業務ID"].Value.ToString();
                textBox6.Text = row.Cells["差異特色"].Value.ToString();
                textBox7.Text = row.Cells["成本"].Value.ToString();
                textBox8.Text = row.Cells["MOQ"].Value.ToString();
                textBox9.Text = row.Cells["一天產能量"].Value.ToString();
                textBox10.Text = row.Cells["業務回覆"].Value.ToString();
                textBox11.Text = row.Cells["註記"].Value.ToString();
                textBox12.Text = row.Cells["DOC_NBR"].Value.ToString();
                textBox13.Text = row.Cells["F09FieldValue"].Value.ToString();

                //dateTimePicker2.Value= row.Cells["開發日期"].Value.ToString();
                DateTime dateTime2;
                if (DateTime.TryParseExact(row.Cells["建立日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime2))
                {
                    dateTimePicker2.Value = dateTime2;
                }
                DateTime dateTime3;
                if (DateTime.TryParseExact(row.Cells["打樣日期"].Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime3))
                {
                    dateTimePicker3.Value = dateTime3;
                }


            }
        }
                

        public void SETTEXT()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;

            textBoxid.Text = null;

            //comboBox1.SelectedValue = null;
        }
        public string GETMAXNO(string NO)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
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
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                        SELECT
                                        ISNULL(MAX(NO),'0')  AS 'NO'
                                        FROM  [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                                        WHERE [NO] LIKE '{0}%'
                                        ORDER BY [NO] DESC
                                        ", NO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        string NEWNO = ds1.Tables["ds1"].Rows[0]["NO"].ToString();

                        if (NEWNO.Equals("0"))
                        {
                            return NO + "-" + "001";
                        }

                        else
                        {
                            int serno = Convert.ToInt16(NEWNO.Substring(6, 3));
                            serno = serno + 1;
                            string temp = serno.ToString();
                            temp = temp.PadLeft(3, '0');
                            return NO + "-" + temp.ToString();
                        }


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

        public void ADD_TB_DEVE_NEWLISTS(
            string NO
            , string NAMES
            , string SPECS
            , string COMMENTS
            , string INGREDIENTS
            , string COSTS
            , string MOQS
            , string MANUPRODS
            , string GETDATES
            , string REPLY
            , string CARESTEDATES
            , string SALES
            , string SALESID

            )
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

                // 關閉再開啟資料庫連線，並開始交易
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                // 清空 StringBuilder 並建立插入語句
                sbSql.Clear();
                sbSql.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                                    (
                                        [NO],
                                        [NAMES],
                                        [SPECS],
                                        [COMMENTS],
                                        [INGREDIENTS],
                                        [COSTS],
                                        [MOQS],
                                        [MANUPRODS],
                                        [GETDATES],
                                        [REPLY],
                                        [CARESTEDATES],
                                        [SALES],
                                        [SALESID]
                                    )
                                    VALUES
                                    (
                                        @NO,
                                        @NAMES,
                                        @SPECS,
                                        @COMMENTS,
                                        @INGREDIENTS,
                                        @COSTS,
                                        @MOQS,
                                        @MANUPRODS,
                                        @GETDATES,
                                        @REPLY,
                                        @CARESTEDATES,
                                        @SALES,
                                        @SALESID
                                    )
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                // 使用參數化查詢，並對每個參數進行賦值
                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@NAMES", NAMES);
                cmd.Parameters.AddWithValue("@SPECS", SPECS);
                cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                cmd.Parameters.AddWithValue("@INGREDIENTS", INGREDIENTS);
                cmd.Parameters.AddWithValue("@COSTS", COSTS);
                cmd.Parameters.AddWithValue("@MOQS", MOQS);
                cmd.Parameters.AddWithValue("@MANUPRODS", MANUPRODS);
                cmd.Parameters.AddWithValue("@GETDATES", GETDATES);
                cmd.Parameters.AddWithValue("@REPLY", REPLY);
                cmd.Parameters.AddWithValue("@CARESTEDATES", CARESTEDATES);
                cmd.Parameters.AddWithValue("@SALES", SALES);
                cmd.Parameters.AddWithValue("@SALESID", SALESID);

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATE_TB_DEVE_NEWLISTS(
            string ID
            ,string NO
            , string NAMES
            , string SPECS
            , string COMMENTS
            , string INGREDIENTS
            , string COSTS
            , string MOQS
            , string MANUPRODS
            , string GETDATES
            , string REPLY
            , string CARESTEDATES
            , string SALES
            , string SALESID
            )
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
                                    UPDATE [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                                    SET 
                                        [NO]=@NO,
                                        [NAMES]=@NAMES,
                                        [SPECS]=@SPECS,
                                        [COMMENTS]=@COMMENTS,
                                        [INGREDIENTS]=@INGREDIENTS,
                                        [COSTS]=@COSTS,
                                        [MOQS]=@MOQS,
                                        [MANUPRODS]=@MANUPRODS,
                                        [GETDATES]=@GETDATES,
                                        [REPLY]=@REPLY,
                                        [CARESTEDATES]=@CARESTEDATES,
                                        [SALES]=@SALES,
                                        [SALESID]=@SALESID
                                    WHERE [ID]=@ID
                                ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;

                // 使用參數化查詢，並對每個參數進行賦值
                cmd.Parameters.AddWithValue("@ID", ID);
                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@NAMES", NAMES);
                cmd.Parameters.AddWithValue("@SPECS", SPECS);
                cmd.Parameters.AddWithValue("@COMMENTS", COMMENTS);
                cmd.Parameters.AddWithValue("@INGREDIENTS", INGREDIENTS);
                cmd.Parameters.AddWithValue("@COSTS", COSTS);
                cmd.Parameters.AddWithValue("@MOQS", MOQS);
                cmd.Parameters.AddWithValue("@MANUPRODS", MANUPRODS);
                cmd.Parameters.AddWithValue("@GETDATES", GETDATES);
                cmd.Parameters.AddWithValue("@REPLY", REPLY);
                cmd.Parameters.AddWithValue("@CARESTEDATES", CARESTEDATES);
                cmd.Parameters.AddWithValue("@SALES", SALES);
                cmd.Parameters.AddWithValue("@SALESID", SALESID);

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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DEL_TB_DEVE_NEWLISTS(string ID)
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
                                    DELETE [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                                    WHERE [ID]='{0}'
                                    ", ID

                                    );

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

        public void ADD_TB_DEVE_NEWLISTS_TB_WKF_EXTERNAL_TASK(string NO)
        {
            string DBNAME = "UOF";
            DataTable DT = SEARCH_TB_DEVE_NEWLISTS(NO);

            if(DT!=null && DT.Rows.Count>=1)
            {
                DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["ACCOUNT"].ToString());

                string account = DT.Rows[0]["ACCOUNT"].ToString();
                string groupId = DT.Rows[0]["GROUP_ID"].ToString();
                string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
                string fillerName = DT.Rows[0]["MV002"].ToString();
                string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

                string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
                string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

                string EXTERNAL_FORM_NBR = DT.Rows[0]["NO"].ToString().Trim();

                int rowscounts = 0;

                XmlDocument xmlDoc = new XmlDocument();
                //建立根節點
                XmlElement Form = xmlDoc.CreateElement("Form");

                //正式的id
                string FORMID = SEARCHFORM_UOF_VERSION_ID("1006.樣品試吃回覆單");

                if (!string.IsNullOrEmpty(FORMID))
                {
                    Form.SetAttribute("formVersionId", FORMID);
                }


                Form.SetAttribute("urgentLevel", "2");
                //加入節點底下
                xmlDoc.AppendChild(Form);

                ////建立節點Applicant
                XmlElement Applicant = xmlDoc.CreateElement("Applicant");
                Applicant.SetAttribute("account", account);
                Applicant.SetAttribute("groupId", groupId);
                Applicant.SetAttribute("jobTitleId", jobTitleId);
                //加入節點底下
                Form.AppendChild(Applicant);

                //建立節點 Comment
                XmlElement Comment = xmlDoc.CreateElement("Comment");
                Comment.InnerText = "申請者意見";
                //加入至節點底下
                Applicant.AppendChild(Comment);

                //建立節點 FormFieldValue
                XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
                //加入至節點底下
                Form.AppendChild(FormFieldValue);

                //建立節點FieldItem
                //ID 表單編號	
                XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "ID");
                FieldItem.SetAttribute("fieldValue", "");
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);


                //建立節點FieldItem
                //F04	
                string USERID=XML_SETID(DT.Rows[0]["USER_GUID"].ToString());

                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F04");
                FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NAME"].ToString() +"("+ DT.Rows[0]["ACCOUNT"].ToString() + ")");
                FieldItem.SetAttribute("realValue", USERID);
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);
                //建立節點FieldItem
                //F01	
                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F01");
                FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NO"].ToString());
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);
                //建立節點FieldItem
                //F02	
                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F02");
                FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NAMES"].ToString());
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);
                //建立節點FieldItem
                //F03	
                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F03");
                FieldItem.SetAttribute("fieldValue", DT.Rows[0]["INGREDIENTS"].ToString());
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);
                //建立節點FieldItem
                //F05	
                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F05");
                FieldItem.SetAttribute("fieldValue", DT.Rows[0]["COMMENTS"].ToString());
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);
                //建立節點FieldItem
                //F09	
                FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "F09");
                FieldItem.SetAttribute("fieldValue","");
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", fillerName);
                FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
                FieldItem.SetAttribute("fillerAccount", account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                FormFieldValue.AppendChild(FieldItem);



                ////用ADDTACK，直接啟動起單
                //ADDTACK(Form);

                //ADD TO DB
                ////string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                connectionString = sqlConn.ConnectionString.ToString();

                StringBuilder queryString = new StringBuilder();




                queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {

                        SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                        command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                        command.Connection.Open();

                        int count = command.ExecuteNonQuery();

                        connection.Close();
                        connection.Dispose();

                        MessageBox.Show("送簽成功");
                    }
                }
                catch
                {
                    MessageBox.Show("送簽失敗");
                }
                finally
                {

                }
            }
            else
            {
                MessageBox.Show("送簽失敗，需求人的工號錯誤");
            }
           
        }

        public string XML_SETID(string REAL_userId)
        {
            // 创建一个XmlDocument对象
            XmlDocument xmlDoc = new XmlDocument();

            // 创建根元素<UserSet>
            XmlElement userSetElement = xmlDoc.CreateElement("UserSet");
            xmlDoc.AppendChild(userSetElement);

            // 创建子元素<Element>，并设置属性type='user'
            XmlElement element = xmlDoc.CreateElement("Element");
            element.SetAttribute("type", "user");
            userSetElement.AppendChild(element);

            // 创建子元素<userId>，并设置其值
            XmlElement userId = xmlDoc.CreateElement("userId");
            userId.InnerText = REAL_userId;
            element.AppendChild(userId);

            // 将XmlDocument对象保存为XML文件，或输出为字符串
            string xmlOutput = xmlDoc.OuterXml;

            return xmlOutput;
        }

        public DataTable SEARCH_TB_DEVE_NEWLISTS(string NO)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                    SELECT *
                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    FROM 
                                    (
                                    SELECT 
                                    [ID]
                                    ,[NO]
                                    ,[NAMES]
                                    ,[SPECS]
                                    ,[SALES]
                                    ,[SALESID]
                                    ,[COMMENTS]
                                    ,[INGREDIENTS]
                                    ,[COSTS]
                                    ,[MOQS]
                                    ,[MANUPRODS]
                                    ,[GETDATES]
                                    ,[REPLY]
                                    ,[CARESTEDATES]

                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=[SALESID]) AS 'MV002'
                                    ,ACCOUNT
                                    FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS],[192.168.1.223].[UOF].[dbo].[TB_EB_USER]
                                    WHERE 1=1
                                    AND [NO]='{0}'
                                    AND [TB_EB_USER].ACCOUNT=[SALESID]COLLATE Chinese_Taiwan_Stroke_BIN
                                    ) AS TEMP
                                    
                              
                                    ", NO);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
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

        public DataTable SEARCHUOFDEP(string ACCOUNT)
        {
            string DBNAME = "UOF";
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [GROUP_NAME] AS 'DEPNAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]+','+[GROUP_NAME]+',False' AS 'DEPNO'
                                    ,[TB_EB_USER].[USER_GUID]
                                    ,[ACCOUNT]
                                    ,[NAME]
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]
                                    ,[TITLE_ID]     
                                    ,[GROUP_NAME]
                                    ,[GROUP_CODE]
                                    ,[TB_EB_EMPL_DEP].ORDERS
                                    FROM [192.168.1.223].[{0}].[dbo].[TB_EB_USER],[192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[{0}].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{1}'
                                    ORDER BY [TB_EB_EMPL_DEP].ORDERS

                                    ", DBNAME, ACCOUNT);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
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

        public string SEARCHFORM_UOF_VERSION_ID(string FORM_NAME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                   SELECT TOP 1 RTRIM(LTRIM(TB_WKF_FORM_VERSION.FORM_VERSION_ID)) FORM_VERSION_ID,TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION,TB_WKF_FORM_VERSION.ISSUE_CTL
                                    ,TB_WKF_FORM.FORM_NAME
                                    FROM [UOF].dbo.TB_WKF_FORM_VERSION,[UOF].dbo.TB_WKF_FORM
                                    WHERE 1=1
                                    AND TB_WKF_FORM_VERSION.FORM_ID=TB_WKF_FORM.FORM_ID
                                    AND TB_WKF_FORM_VERSION.ISSUE_CTL=1
                                    AND FORM_NAME='{0}'
                                    ORDER BY TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION DESC

                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["FORM_VERSION_ID"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMM"),textBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string DATES = DateTime.Now.ToString("yyyy-MM");
            DATES = DATES.Substring(2, 5);
            string NO = GETMAXNO(DATES);
            textBox2.Text = NO;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string NO = textBox2.Text.Trim();
            string NAMES = textBox3.Text.Trim();
            string SPECS = textBox4.Text.Trim();
            string COMMENTS = textBox11.Text.Trim();
            string INGREDIENTS = textBox6.Text.Trim();
            string COSTS = textBox7.Text.Trim();
            string MOQS = textBox8.Text.Trim();
            string MANUPRODS = textBox9.Text.Trim();
            string GETDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd");
            string REPLY = textBox10.Text.Trim();
            string CARESTEDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd");
            string SALES = comboBox1.Text.ToString();
            string SALESID = textBox5.Text.Trim();

            ADD_TB_DEVE_NEWLISTS(
                NO
                , NAMES
                , SPECS
                , COMMENTS
                , INGREDIENTS
                , COSTS
                , MOQS
                , MANUPRODS
                , GETDATES
                , REPLY
                , CARESTEDATES
                , SALES
                , SALESID
                );

            SEARCH(dateTimePicker1.Value.ToString("yyyyMM"), textBox1.Text.Trim());

        }
        private void button6_Click(object sender, EventArgs e)
        {
            string ID = textBoxid.Text.Trim();
            string NO = textBox2.Text.Trim();
            string NAMES = textBox3.Text.Trim();
            string SPECS = textBox4.Text.Trim();
            string COMMENTS = textBox11.Text.Trim();
            string INGREDIENTS = textBox6.Text.Trim();
            string COSTS = textBox7.Text.Trim();
            string MOQS = textBox8.Text.Trim();
            string MANUPRODS = textBox9.Text.Trim();
            string GETDATES = dateTimePicker3.Value.ToString("yyyy/MM/dd");
            string REPLY = textBox10.Text.Trim();
            string CARESTEDATES = dateTimePicker2.Value.ToString("yyyy/MM/dd");
            string SALES = comboBox1.Text.ToString();
            string SALESID = textBox5.Text.Trim();

            UPDATE_TB_DEVE_NEWLISTS(
                ID
                ,NO
                , NAMES
                , SPECS
                , COMMENTS
                , INGREDIENTS
                , COSTS
                , MOQS
                , MANUPRODS
                , GETDATES
                , REPLY
                , CARESTEDATES
                , SALES
                , SALESID
                );


            SEARCH(dateTimePicker1.Value.ToString("yyyyMM"), textBox1.Text.Trim());
        }
        private void button7_Click(object sender, EventArgs e)
        {
            string ID = textBoxid.Text.Trim();

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DEL_TB_DEVE_NEWLISTS(ID);
                SEARCH(dateTimePicker1.Value.ToString("yyyyMM"), textBox1.Text.Trim()); ;

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        public void SETFASTREPORT(string YYYYMM)
        {
            string YY = YYYYMM.Substring(2, 2);
            string MM = YYYYMM.Substring(4, 2);
            string YYYY = YYYYMM.Substring(0, 4);
            string P1 = YYYY + "年/" + MM+ "月份";

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(YY+"-"+ MM);
            Report report1 = new Report();
            report1.Load(@"REPORT\研發新品清單.frx");

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", P1);
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string YYMM)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            [NO] AS '編號' 
                            ,[NAMES] AS '商品'
                            ,[SPECS] AS '規格'
                            ,[SALES] AS '業務'
                            ,[COMMENTS] AS '註記'
                            ,[INGREDIENTS] AS '差異特色'
                            ,CONVERT(NVARCHAR,[GETDATES],112)  AS '打樣日期'
                            ,[REPLY] AS '業務回覆'
                            ,[SALESID] AS '業務ID'
                            ,[COSTS] AS '成本'
                            ,[MOQS] AS 'MOQ'
                            ,[MANUPRODS] AS '一天產能量'
                            ,CONVERT(NVARCHAR,[CARESTEDATES],112) AS '建立日期'
                            ,[ID]

                                


                            FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]
                            WHERE 1=1
                            AND [NO] LIKE '%{0}%' 
                            ORDER BY [NO]
                            ", YYMM);

            return SB;

        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker4.Value.ToString("yyyyMM"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            ADD_TB_DEVE_NEWLISTS_TB_WKF_EXTERNAL_TASK(textBox2.Text);
        }


        #endregion


    }
}
