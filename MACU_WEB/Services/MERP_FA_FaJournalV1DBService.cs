using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Web.Configuration;
using MACU_WEB.BIServices;
using MACU_WEB.Models._Base;
using System.ComponentModel;
using System.Globalization;


namespace MACU_WEB.Services
{
    public class MERP_FA_FaJournalV1DBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_JournalV1";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_JournalV1
        public List<FA_JournalV1> FA_JournalV1_GetDataList()
        {

            return db.FA_JournalV1.ToList();
        }

        public FA_JournalV1 FA_JournalV1_GetDataById(int p_iId)
        {

            FA_JournalV1 l_oFindFile = db.FA_JournalV1.Find(p_iId);
            return l_oFindFile;
        }


        //20201217 CCL+
        public void FA_JournalV1_SqlDBCreate(IXLTable p_oNewTable)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
                                   SubpDate, SubpNo, AccountNo, SubjectName, 
                                   DetailAccountNo, DetailSubjectName, 
                                   DepartNo, DepartName, CreditAmount, DebitAmount,  
                                   FiscalYear, AccountPeriod, 
                                   CreateTime, UpdateTime)
                                   VALUES (
                                     N'{0}', N'{1}', N'{2}', N'{3}', 
                                     N'{4}', N'{5}', N'{6}', N'{7}', 
                                     N'{8}', N'{9}', N'{10}', N'{11}', 
                                     N'{12}', N'{13}'
                                   )";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //DateTimeConverter l_oDT = new DateTimeConverter();
            String l_sSubpDate = "";
            int l_iYear = 0, l_iMonth = 0;

            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //共41行
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                l_sSubpDate = l_row.Cell(1).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                l_sSubpDate = l_sSubpDate.Substring(0, l_sSubpDate.IndexOf(' '));

                if(l_iIndex == 2)
                {
                    //根據第一列的傳票日期,取出年,月
                    string l_sTmpStr = DateStringProcess.Del_MonthDayZero(l_sSubpDate, "/", ""); //最後參數為"",代表不替換
                    l_iYear = DateStringProcess.m_Year;
                    l_iMonth = DateStringProcess.m_Month;
                }


                l_sExeSqlCmd = string.Format(l_sInsSqlCmd,
                                                l_sSubpDate,                                                
                                                l_row.Cell(2).Value.ToString(),
                                                l_row.Cell(3).Value.ToString(),
                                                l_row.Cell(4).Value.ToString(),
                                                l_row.Cell(5).Value.ToString(),
                                                l_row.Cell(6).Value.ToString(),
                                                l_row.Cell(7).Value.ToString(),
                                                l_row.Cell(8).Value.ToString(),
                                                l_row.Cell(9).Value.ToString(),
                                                l_row.Cell(10).Value.ToString(),
                                                l_iYear.ToString(),
                                                l_iMonth.ToString(),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                try
                {

                    try
                    {
                        l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn, l_oSqlTrans);
                        l_iRtnInsCount = l_oSqlCmdObj.ExecuteNonQuery();


                    }
                    catch (Exception ex)
                    {
                        l_iRtnInsCount = -1;
                    }

                    if (l_iRtnInsCount == -1)
                    {
                        throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                    }


                }
                catch (Exception ex)
                {
                    l_oSqlTrans.Rollback();
                    l_oSqlConn.Close();
                    Trace.WriteLine("Err: " + l_iRtnInsCount);
                }


            }

            //成功寫入
            l_oSqlTrans.Commit();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);
        }

        //20201229 CCl+ 要求改成匯入的Excel File借方金額要在前,貸方金額在後 //////////////////////////////////
        public void FA_JournalV1_SqlDBCreateV1(IXLTable p_oNewTable)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            
            //string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
            //                       SubpDate, SubpNo, AccountNo, SubjectName, 
            //                       DetailAccountNo, DetailSubjectName, 
            //                       DepartNo, DepartName, CreditAmount, DebitAmount,  
            //                       FiscalYear, AccountPeriod, 
            //                       CreateTime, UpdateTime)
            //                       VALUES (
            //                         N'{0}', N'{1}', N'{2}', N'{3}', 
            //                         N'{4}', N'{5}', N'{6}', N'{7}', 
            //                         N'{8}', N'{9}', N'{10}', N'{11}', 
            //                         N'{12}', N'{13}'
            //                       )";

            //20201229 CCL Modify
            string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
                                   SubpDate, SubpNo, AccountNo, SubjectName, 
                                   DetailAccountNo, DetailSubjectName, 
                                   DepartNo, DepartName, DebitAmount, CreditAmount,   
                                   FiscalYear, AccountPeriod, 
                                   CreateTime, UpdateTime)
                                   VALUES (
                                     N'{0}', N'{1}', N'{2}', N'{3}', 
                                     N'{4}', N'{5}', N'{6}', N'{7}', 
                                     N'{8}', N'{9}', N'{10}', N'{11}', 
                                     N'{12}', N'{13}'
                                   )";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //DateTimeConverter l_oDT = new DateTimeConverter();
            String l_sSubpDate = "";
            int l_iYear = 0, l_iMonth = 0;

            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //共41行
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                l_sSubpDate = l_row.Cell(1).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                l_sSubpDate = l_sSubpDate.Substring(0, l_sSubpDate.IndexOf(' '));

                if (l_iIndex == 2)
                {
                    //根據第一列的傳票日期,取出年,月
                    string l_sTmpStr = DateStringProcess.Del_MonthDayZero(l_sSubpDate, "/", ""); //最後參數為"",代表不替換
                    l_iYear = DateStringProcess.m_Year;
                    l_iMonth = DateStringProcess.m_Month;
                }


                l_sExeSqlCmd = string.Format(l_sInsSqlCmd,
                                                l_sSubpDate,
                                                l_row.Cell(2).Value.ToString(),
                                                l_row.Cell(3).Value.ToString(),
                                                l_row.Cell(4).Value.ToString(),
                                                l_row.Cell(5).Value.ToString(),
                                                l_row.Cell(6).Value.ToString(),
                                                l_row.Cell(7).Value.ToString(),
                                                l_row.Cell(8).Value.ToString(),
                                                l_row.Cell(9).Value.ToString(),
                                                l_row.Cell(10).Value.ToString(),
                                                l_iYear.ToString(),
                                                l_iMonth.ToString(),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                try
                {

                    try
                    {
                        l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn, l_oSqlTrans);
                        l_iRtnInsCount = l_oSqlCmdObj.ExecuteNonQuery();


                    }
                    catch (Exception ex)
                    {
                        l_iRtnInsCount = -1;
                    }

                    if (l_iRtnInsCount == -1)
                    {
                        throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                    }


                }
                catch (Exception ex)
                {
                    l_oSqlTrans.Rollback();
                    l_oSqlConn.Close();
                    Trace.WriteLine("Err: " + l_iRtnInsCount);
                }


            }

            //成功寫入
            l_oSqlTrans.Commit();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);
        }


        // ///////////////////////////////////////////////////////////////////////////////////////////////////

        //20210204 CCL+ 
        public void FA_JournalV1_SqlDBDeleteByYearPeriod(string p_sYear, string p_sPeriod)
        {
            int l_iRtnDelCount = 0;
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            //SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sDelSqlCmd = @"DELETE FROM " + TB_NAME + @" WHERE 
                                   AccountPeriod = N'{0}' 
                                   AND FiscalYear = N'{1}'
                                   AND IsValid = 1";

            l_sExeSqlCmd = string.Format(l_sDelSqlCmd, p_sPeriod, p_sYear);

            try
            {


                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_iRtnDelCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnDelCount = -1;
                }

                if (l_iRtnDelCount == -1)
                {
                    throw new ArgumentException("FA_FaJournalV1資料刪除發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();

                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnDelCount);
                Trace.WriteLine("ErrMsg: " + errmsg);
            }

            //成功關閉DB Conn          
            l_oSqlConn.Close();

        }


        //20201217 CCL+
        /* 20210204 CCL-
        public void FA_JournalV1_SqlDBDeleteByPeriod(string p_sPeriod)
        {
            int l_iRtnDelCount = 0;
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            //SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sDelSqlCmd = @"DELETE FROM " + TB_NAME + @" WHERE 
                                   AccountPeriod = N'{0}'  AND IsValid = 1";

            l_sExeSqlCmd = string.Format(l_sDelSqlCmd, p_sPeriod);

            try
            {


                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_iRtnDelCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnDelCount = -1;
                }

                if (l_iRtnDelCount == -1)
                {
                    throw new ArgumentException("FA_FaJournal資料刪除發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();

                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnDelCount);
                Trace.WriteLine("ErrMsg: " + errmsg);
            }

            //成功關閉DB Conn          
            l_oSqlConn.Close();

        }
        */

        //用EntityFramework優化後仍需6分鐘,改用ADO.NET
        /*
        public void FA_JournalV1_DBCreate(IXLTable p_oNewTable)
        {
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();
            for (int i = 0; i < p_oNewTable.RowCount() / BATCH_COUNT; i++)
            {

                //分批次加入,分次存
                for (int j = 1; j <= BATCH_COUNT; j++)
                {
                    if (j == 1 && i == 0) continue; //跳過第一行中文欄位

                    l_oNewRow = new FA_FaJournal();

                    l_iIndex = (i * BATCH_COUNT) + j;

                    IXLRangeRow l_row = p_oNewTable.Row(l_iIndex);
                    //共39行
                    l_oNewRow.SubpType = l_row.Cell(1).Value.ToString();    //傳票類別
                    l_oNewRow.DebitAmount = l_row.Cell(2).Value.ToString();     //借方金額
                    l_oNewRow.CreditAmount = l_row.Cell(3).Value.ToString();    //貸方金額
                    l_oNewRow.CurrencyName = l_row.Cell(4).Value.ToString();      //幣別
                    l_oNewRow.FiscalYear = l_row.Cell(5).Value.ToString();  //會計年度
                    l_oNewRow.AccountPeriod = l_row.Cell(6).Value.ToString(); //會計期別
                    //l_oNewRow.GUID = Guid.Parse(l_row.Cell(7).Value.ToString());      //GUID
                    l_oNewRow.GUID = l_row.Cell(7).Value.ToString();      //GUID
                    l_oNewRow.SubpDate = l_row.Cell(8).Value.ToString().Trim();      //傳票日期
                    l_oNewRow.SubpNo = l_row.Cell(9).Value.ToString();        //傳票號碼
                    l_oNewRow.SubpSummary = l_row.Cell(10).Value.ToString();   //傳票摘要
                    l_oNewRow.SubjectName = l_row.Cell(11).Value.ToString();   //科目名稱
                    l_oNewRow.AccountNo = l_row.Cell(12).Value.ToString();   //科目編號
                    l_oNewRow.DetailAccountNo = l_row.Cell(13).Value.ToString();   //明細科目編號
                    l_oNewRow.DetailSubjectName = l_row.Cell(14).Value.ToString();   //明細科目名稱
                    l_oNewRow.DepartNo = l_row.Cell(15).Value.ToString();   //部門代號
                    l_oNewRow.DepartName = l_row.Cell(16).Value.ToString();   //部門簡稱
                    l_oNewRow.ProjectNo = l_row.Cell(17).Value.ToString();   //專案代號
                    l_oNewRow.ProjectAbbr = l_row.Cell(18).Value.ToString();   //專案簡稱
                    l_oNewRow.ObjectCateg = l_row.Cell(19).Value.ToString();   //對象類別
                    l_oNewRow.ObjectNo = l_row.Cell(20).Value.ToString();   //對象編號
                    l_oNewRow.CurrencyNo = l_row.Cell(21).Value.ToString();   //幣別代號
                    l_oNewRow.ExchangeRate = l_row.Cell(22).Value.ToString().ToString();   //匯率
                    l_oNewRow.OriginCurrency = l_row.Cell(23).Value.ToString();   //原幣金額
                    l_oNewRow.LocalCurrencyAmount = l_row.Cell(24).Value.ToString();   //本幣金額
                    l_oNewRow.Spare1No = l_row.Cell(25).Value.ToString();   //備用1編號
                    l_oNewRow.Spare1Abbr = l_row.Cell(26).Value.ToString();   //備用1簡稱
                    l_oNewRow.Spare2No = l_row.Cell(27).Value.ToString();   //備用2編號
                    l_oNewRow.Spare2Abbr = l_row.Cell(28).Value.ToString();   //備用2簡稱
                    l_oNewRow.Spare3No = l_row.Cell(29).Value.ToString();   //備用3編號
                    l_oNewRow.Spare3Abbr = l_row.Cell(30).Value.ToString();   //備用3簡稱
                    l_oNewRow.Spare4No = l_row.Cell(31).Value.ToString();   //備用4簡稱
                    l_oNewRow.Spare4Abbr = l_row.Cell(32).Value.ToString();   //備用4簡稱
                    l_oNewRow.Spare5No = l_row.Cell(33).Value.ToString();   //備用5簡稱
                    l_oNewRow.Spare5Abbr = l_row.Cell(34).Value.ToString();   //備用5簡稱
                    l_oNewRow.Summary1 = l_row.Cell(35).ToString();   //摘要1
                    l_oNewRow.AccountSubjects = l_row.Cell(36).Value.ToString();   //會計科目
                    l_oNewRow.Summary = l_row.Cell(37).Value.ToString();   //摘要
                    l_oNewRow.Category = l_row.Cell(38).Value.ToString();   //類別
                    l_oNewRow.SubjectAlias = l_row.Cell(39).Value.ToString();   //科目別名
                    l_oNewRow.CreateTime = DateTime.Now;
                    l_oNewRow.UpdateTime = DateTime.Now;

                    db.FA_FaJournal.Add(l_oNewRow);
                }

                try
                {
                    l_iDBStatus = db.SaveChanges();
                    Trace.WriteLine((i + 1) + "次 : " + l_iDBStatus.ToString() + " 筆");
                }
                catch (Exception ex)
                {
                    int l_iError = l_iDBStatus;
                    string errmsg = ex.Message.ToString();
                }


            }

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

        }
        */

        /* 20210204 CCL-
        public List<FA_JournalV1> FA_JournalV1_GetDataByMonthVal(string p_sVal)
        {
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oFaJournals;

        }

        public List<FA_JournalV1> FA_JournalV1_GetDataByMonthValPage(string p_sVal, int p_iPageing)
        {
            //分頁傳回Paging Data
            const int PAGE_COUNT = 50;
            int l_iShowRange = PAGE_COUNT * p_iPageing;

            //去除之前的範圍
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => m.AccountPeriod == p_sVal).OrderBy(m => m.Id).Skip(l_iShowRange).ToList();
            //再從剩下的傳回300筆
            List<FA_JournalV1> l_oLastFaJournals = l_oFaJournals.Take(PAGE_COUNT).ToList();

            //List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oLastFaJournals;

        }
        */

        //20210204 CCL+ 修正以抓年月 ///////////////////////////////////////////////////////////////////////
        public List<FA_JournalV1> FA_JournalV1_GetDataByYearMonthVal(string p_sYear, string p_sMonth)
        {
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => (m.AccountPeriod == p_sMonth)
                                                                    && (m.FiscalYear == p_sYear)).ToList();

            return l_oFaJournals;

        }

        public List<FA_JournalV1> FA_JournalV1_GetDataByYearMonthValPage(string p_sYear, string p_sMonth, int p_iPageing)
        {
            //分頁傳回Paging Data
            const int PAGE_COUNT = 50;
            int l_iShowRange = PAGE_COUNT * p_iPageing;

            //去除之前的範圍
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => (m.AccountPeriod == p_sMonth) && 
                                                                    (m.FiscalYear == p_sYear) ).OrderBy(m => m.Id).Skip(l_iShowRange).ToList();
            //再從剩下的傳回300筆
            List<FA_JournalV1> l_oLastFaJournals = l_oFaJournals.Take(PAGE_COUNT).ToList();

            //List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oLastFaJournals;

        }

        public Boolean FA_JournalV1_FindDataByYearMonthVal(string p_sYear, string p_sMonth)
        {
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => (m.AccountPeriod == p_sMonth) && 
                                                                        (m.FiscalYear == p_sYear)).ToList();

            if (l_oFaJournals.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        /// ////////////////////////////////////////////////////////////////////////////////////////////////

        /* 20210204 CCL-
        public Boolean FA_JournalV1_FindDataByMonthVal(string p_sVal)
        {
            List<FA_JournalV1> l_oFaJournals = db.FA_JournalV1.Where(m => m.AccountPeriod == p_sVal).ToList();

            if (l_oFaJournals.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        */

        //20201218 CCL+ For Processing 區間日期Excel 資料庫處理 /////////////////////////////////////
        /* 20210204 CCL-
        public DataSet FA_JournalV1_SqlGetDataListByOptions(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                     WHERE AccountPeriod = N'{0}' AND 
                                           SubpDate BETWEEN N'{1}' AND N'{2}' AND IsValid = 1 
                                     ORDER BY Id
                                   ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                            p_oOption.m_sAccountPeriod,
                                            p_oOption.m_sStartDate,
                                            p_oOption.m_sEndDate);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }
        */

        public DataSet FA_JournalV1_SqlGetDataListByOptions(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                     WHERE AccountPeriod = N'{0}' AND 
                                           FiscalYear = N'{1}' AND 
                                           SubpDate BETWEEN N'{2}' AND N'{3}' AND IsValid = 1 
                                     ORDER BY Id
                                   ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                            p_oOption.m_sAccountPeriod,
                                            p_oOption.m_sFiscalYear,
                                            p_oOption.m_sStartDate,
                                            p_oOption.m_sEndDate);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournalV1資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }

        public List<FA_JournalV1> FA_JournalV1_DataTableTo_FaJournalsList(DataSet p_oDataSet)
        {
            string l_sGuidStr = "";
            List<FA_JournalV1> l_oRtnList = new List<FA_JournalV1>();
            FA_JournalV1 l_oTmpFARow = null;


            foreach (DataRow l_row in p_oDataSet.Tables[0].Rows)
            {
                //Guid l_oGuid = new Guid(l_row.Field<Guid>(7).ToString());
                //GuidConverter l_oCov = new GuidConverter();
                //Guid l_oGuid = (Guid)l_oCov.ConvertFromString(l_row.Field<string>(7));
                //l_sGuidStr = GUIDStringProcess.GuidProcess();

                l_oTmpFARow = new FA_JournalV1();
                l_oTmpFARow.Id = l_row.Field<int>(0);
                l_oTmpFARow.SubpDate = l_row.Field<string>(1);
                l_oTmpFARow.SubpNo = l_row.Field<string>(2);
                l_oTmpFARow.AccountNo = l_row.Field<string>(3);
                l_oTmpFARow.SubjectName = l_row.Field<string>(4);
                l_oTmpFARow.DetailAccountNo = l_row.Field<string>(5);
                l_oTmpFARow.DetailSubjectName = l_row.Field<string>(6);
                l_oTmpFARow.DepartNo = l_row.Field<string>(7);
                l_oTmpFARow.DepartName = l_row.Field<string>(8);
                l_oTmpFARow.CreditAmount = l_row.Field<string>(9);
                l_oTmpFARow.DebitAmount = l_row.Field<string>(10);
                l_oTmpFARow.FiscalYear = l_row.Field<string>(11);
                l_oTmpFARow.AccountPeriod = l_row.Field<string>(12);
                l_oTmpFARow.IsValid = l_row.Field<int>(13);              
                l_oTmpFARow.CreateTime = l_row.Field<DateTime>(14);
                l_oTmpFARow.UpdateTime = l_row.Field<DateTime>(15);

                l_oRtnList.Add(l_oTmpFARow);
            }

            return l_oRtnList;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////

        //20201221 CCL+ Sql方式////////////////////////////////////////////////////////////
        /* 20210204 CCL-
        public DataSet FA_JournalV1_SqlGetDataListByOptions2(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            // 20201225 CCL-
            //string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
            //                         WHERE AccountPeriod = N'{0}' AND 
            //                               SubpDate BETWEEN N'{1}' AND N'{2}' AND 
            //                               DepartNo = N'{3}' 
            //                               AND IsValid = 1 
            //                         ORDER BY Id
            //                       ";
            

            string l_sSelSqlCmd = @"SELECT* FROM " + TB_NAME + @"
                                    WHERE AccountPeriod = N'{0}' AND
                                           CONVERT(DATE, SubpDate) BETWEEN CONVERT(DATE, N'{1}') AND CONVERT(DATE, N'{2}') AND
                                           DepartNo = N'{3}'
                                           AND IsValid = 1
                                     ORDER BY Id
                                  ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                            p_oOption.m_sAccountPeriod,
                                            p_oOption.m_sStartDate,
                                            p_oOption.m_sEndDate,
                                            p_oOption.m_sShop);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }
        */

        public DataSet FA_JournalV1_SqlGetDataListByOptions2(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            /* 20201225 CCL-
            string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                     WHERE AccountPeriod = N'{0}' AND 
                                           SubpDate BETWEEN N'{1}' AND N'{2}' AND 
                                           DepartNo = N'{3}' 
                                           AND IsValid = 1 
                                     ORDER BY Id
                                   ";
            */

            string l_sSelSqlCmd = @"SELECT* FROM " + TB_NAME + @"
                                    WHERE AccountPeriod = N'{0}' AND 
                                          FiscalYear = N'{1}' AND 
                                           CONVERT(DATE, SubpDate) BETWEEN CONVERT(DATE, N'{2}') AND CONVERT(DATE, N'{3}') AND
                                           DepartNo = N'{4}'
                                           AND IsValid = 1
                                     ORDER BY Id
                                  ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                            p_oOption.m_sAccountPeriod,
                                            p_oOption.m_sFiscalYear,
                                            p_oOption.m_sStartDate,
                                            p_oOption.m_sEndDate,
                                            p_oOption.m_sShop);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournalV1資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }

        ///////////////////////////////////////////////////////////////////////////////////


        //20201227 CCL+ Sql方式////////////////////////////////////////////////////////////
        /* 20210204 CCL-
        public DataSet FA_JournalV1_SqlGetDataListByOptions3(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            // 20201225 CCL-
            //string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
            //                         WHERE AccountPeriod = N'{0}' AND 
            //                               SubpDate BETWEEN N'{1}' AND N'{2}' AND 
            //                               DepartNo = N'{3}' 
            //                               AND IsValid = 1 
            //                         ORDER BY Id
            //                       ";
            

            string l_sSelSqlCmd = @"SELECT* FROM " + TB_NAME + @"
                                    WHERE AccountPeriod = N'{0}' AND
                                           CONVERT(DATE, SubpDate) BETWEEN CONVERT(DATE, N'{1}') AND CONVERT(DATE, N'{2}') AND
                                           DepartNo = N'{3}'
                                           AND IsValid = 1
                                     ORDER BY Id
                                  ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
            //                                p_oOption.m_sAccountPeriod,
            //                                p_oOption.m_sStartDate,
            //                                p_oOption.m_sEndDate,
            //                                p_oOption.m_sShop);

            //20201227 CCL Mod
            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                           p_oOption.m_sAccountPeriod,
                                           p_oOption.m_sStartDate,
                                           p_oOption.m_sEndDate,
                                           p_oOption.m_sTmpShopNo);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }
        */

        public DataSet FA_JournalV1_SqlGetDataListByOptions3(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            /* 20201225 CCL-
            string l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                     WHERE AccountPeriod = N'{0}' AND 
                                           SubpDate BETWEEN N'{1}' AND N'{2}' AND 
                                           DepartNo = N'{3}' 
                                           AND IsValid = 1 
                                     ORDER BY Id
                                   ";
            */

            string l_sSelSqlCmd = @"SELECT* FROM " + TB_NAME + @"
                                    WHERE AccountPeriod = N'{0}' AND 
                                          FiscalYear = N'{1}' AND 
                                           CONVERT(DATE, SubpDate) BETWEEN CONVERT(DATE, N'{2}') AND CONVERT(DATE, N'{3}') AND
                                           DepartNo = N'{4}'
                                           AND IsValid = 1
                                     ORDER BY Id
                                  ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
            //                                p_oOption.m_sAccountPeriod,
            //                                p_oOption.m_sStartDate,
            //                                p_oOption.m_sEndDate,
            //                                p_oOption.m_sShop);

            //20201227 CCL Mod
            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                           p_oOption.m_sAccountPeriod,
                                           p_oOption.m_sFiscalYear,
                                           p_oOption.m_sStartDate,
                                           p_oOption.m_sEndDate,
                                           p_oOption.m_sTmpShopNo);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournalV1資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //轉成FA_Journals


            return l_oRtnDataSet;
        }

        ///////////////////////////////////////////////////////////////////////////////////


        /*
        //20201227 CCL+ Sql方式////////////////////////////////////////////////////////////
        public List<DataSet> FA_JournalV1_SqlGetDataListByOptions3(MERP_ProcessExcelOptions p_oOption)
        {
            string l_sShopList = "";
            int l_iShopCount = 0;

            DataSet l_oRtnDataSet = null;
            List < DataSet > l_oRtnDTList = new List<DataSet>();

            SqlCommand l_oSqlCmdObj = null;

            if(p_oOption.m_iShopCount > 0)
            {
                foreach(string shopNo in p_oOption.m_sShopList)
                {
                    ++l_iShopCount;
                    l_sShopList += "N'" + shopNo + "'";
                    if(l_iShopCount < p_oOption.m_iShopCount)
                    {
                        l_sShopList += ",";
                    }
                }
            } else
            {
                //Only One Shop
                l_sShopList += "N'" + p_oOption.m_sShop + "'";
            }

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            

            string l_sSelSqlCmd = @"SELECT* FROM " + TB_NAME + @"
                                    WHERE AccountPeriod = N'{0}' AND
                                           CONVERT(DATE, SubpDate) BETWEEN CONVERT(DATE, N'{1}') AND CONVERT(DATE, N'{2}') AND
                                           DepartNo IN ({3}) 
                                           AND IsValid = 1
                                     ORDER BY Id
                                  ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                            p_oOption.m_sAccountPeriod,
                                            p_oOption.m_sStartDate,
                                            p_oOption.m_sEndDate,
                                            l_sShopList);

            try
            {

                try
                {
                    l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                    l_oRtnDataSet = new DataSet(); //DataSet不能為null必須有物件
                    SqlDataAdapter l_oSqlDataAd = new SqlDataAdapter(l_oSqlCmdObj);

                    l_oSqlDataAd.Fill(l_oRtnDataSet);
                    //SqlDataReader l_oDataReader = l_oSqlCmdObj.ExecuteReader();                    
                    //l_iRtnCount = l_oSqlCmdObj.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    l_iRtnCount = -1;
                }

                if (l_iRtnCount == -1)
                {
                    throw new ArgumentException("FA_FaJournal資料新增發生錯誤!!");
                }

            }
            catch (Exception ex)
            {
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnCount);
            }


            //Release
            l_oSqlCmdObj.Dispose();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            //分群
            foreach(string shopNo in p_oOption.m_sShopList)
            {
                l_oRtnDataSet.
            }



            return l_oRtnDataSet;
        }
        */
        ///////////////////////////////////////////////////////////////////////////////////


        #endregion

    }
}