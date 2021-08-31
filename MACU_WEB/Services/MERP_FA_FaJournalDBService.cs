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
    public class MERP_FA_FaJournalDBService
    {
        private const int BATCH_COUNT = 100;
        private const string DB_NAME = "FA_FaJournal";
        //private const int ADD_COUNT = 100;

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_FaJournal
        public List<FA_FaJournal> FA_FaJournal_GetDataList()
        {            

            return db.FA_FaJournal.ToList();
        }

        public FA_FaJournal FA_FaJournal_GetDataById(int p_iId)
        {
            
            FA_FaJournal l_oFindFile = db.FA_FaJournal.Find(p_iId);
            return l_oFindFile;
        }


        //20201217 CCL+
        public void FA_FaJournal_SqlDBCreate(IXLTable p_oNewTable)
        {
            SqlCommand l_oSqlCmdObj = null;
     
            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sInsSqlCmd = @"INSERT INTO " + DB_NAME + @"(
                                   SubpType, DebitAmount, CreditAmount, CurrencyName, 
                                   FiscalYear, AccountPeriod, GUID, SubpDate, 
                                   SubpNo, SubpSummary, SubjectName, AccountNo, 
                                   DetailAccountNo, DetailSubjectName, DepartNo, DepartName, 
                                   ProjectNo, ProjectAbbr, ObjectCateg, ObjectNo, 
                                   CurrencyNo, ExchangeRate, OriginCurrency, LocalCurrencyAmount, 
                                   Spare1No, Spare1Abbr, Spare2No, Spare2Abbr, 
                                   Spare3No, Spare3Abbr, Spare4No, Spare4Abbr,
                                   Spare5No, Spare5Abbr, Summary1, AccountSubjects, 
                                   Summary, Category, SubjectAlias, CreateTime,
                                   UpdateTime)
                                   VALUES (
                                     N'{0}', N'{1}', N'{2}', N'{3}', 
                                     N'{4}', N'{5}', N'{6}', N'{7}', 
                                     N'{8}', N'{9}', N'{10}', N'{11}', 
                                     N'{12}', N'{13}', N'{14}', N'{15}', 
                                     N'{16}', N'{17}', N'{18}', N'{19}', 
                                     N'{20}', N'{21}', N'{22}', N'{23}', 
                                     N'{24}', N'{25}', N'{26}', N'{27}', 
                                     N'{28}', N'{29}', N'{30}', N'{31}', 
                                     N'{32}', N'{33}', N'{34}', N'{35}', 
                                     N'{36}', N'{37}', N'{38}', N'{39}', 
                                     N'{40}'
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

            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //共41行
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                l_sSubpDate =  l_row.Cell(8).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                l_sSubpDate = l_sSubpDate.Substring(0, l_sSubpDate.IndexOf(' '));

                l_sExeSqlCmd = string.Format(l_sInsSqlCmd,
                                                l_row.Cell(1).Value.ToString(),
                                                l_row.Cell(2).Value.ToString(),
                                                l_row.Cell(3).Value.ToString(),
                                                l_row.Cell(4).Value.ToString(),
                                                l_row.Cell(5).Value.ToString(),
                                                l_row.Cell(6).Value.ToString(),
                                                l_row.Cell(7).Value.ToString(),
                                                l_sSubpDate,
                                                l_row.Cell(9).Value.ToString(),
                                                l_row.Cell(10).Value.ToString(),
                                                l_row.Cell(11).Value.ToString(),
                                                l_row.Cell(12).Value.ToString(),
                                                l_row.Cell(13).Value.ToString(),
                                                l_row.Cell(14).Value.ToString(),
                                                l_row.Cell(15).Value.ToString(),
                                                l_row.Cell(16).Value.ToString(),
                                                l_row.Cell(17).Value.ToString(),
                                                l_row.Cell(18).Value.ToString(),
                                                l_row.Cell(19).Value.ToString(),
                                                l_row.Cell(20).Value.ToString(),
                                                l_row.Cell(21).Value.ToString(),
                                                l_row.Cell(22).Value.ToString(),
                                                l_row.Cell(23).Value.ToString(),
                                                l_row.Cell(24).Value.ToString(),
                                                l_row.Cell(25).Value.ToString(),
                                                l_row.Cell(26).Value.ToString(),
                                                l_row.Cell(27).Value.ToString(),
                                                l_row.Cell(28).Value.ToString(),
                                                l_row.Cell(29).Value.ToString(),
                                                l_row.Cell(30).Value.ToString(),
                                                l_row.Cell(31).Value.ToString(),
                                                l_row.Cell(32).Value.ToString(),
                                                l_row.Cell(33).Value.ToString(),
                                                l_row.Cell(34).Value.ToString(),
                                                l_row.Cell(35).Value.ToString(),
                                                l_row.Cell(36).Value.ToString(),
                                                l_row.Cell(37).Value.ToString(),
                                                l_row.Cell(38).Value.ToString(),
                                                l_row.Cell(39).Value.ToString(),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                try
                {

                    try
                    {
                        l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn, l_oSqlTrans);
                        l_iRtnInsCount = l_oSqlCmdObj.ExecuteNonQuery();
     

                    } catch(Exception ex)
                    {
                        l_iRtnInsCount = -1;
                    }
                    
                    if(l_iRtnInsCount == -1)
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

        //20201217 CCL+
        public void FA_FaJournal_SqlDBDeleteByPeriod(string p_sPeriod)
        {
            int l_iRtnDelCount = 0;
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            //SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sDelSqlCmd = @"DELETE FROM " + DB_NAME + @" WHERE 
                                   AccountPeriod = N'{0}' ";

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


        //用EntityFramework優化後仍需6分鐘,改用ADO.NET
        public void FA_FaJournal_DBCreate(IXLTable p_oNewTable)
        {
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();
            for (int i=0; i< p_oNewTable.RowCount() / BATCH_COUNT; i++)
            {
                
                //分批次加入,分次存
                for (int j=1; j<= BATCH_COUNT; j++)
                {
                    if (j == 1 && i==0) continue; //跳過第一行中文欄位

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
                    Trace.WriteLine((i+1) + "次 : " + l_iDBStatus.ToString() + " 筆");
                }               
                catch (Exception ex)
                {
                    int l_iError = l_iDBStatus;
                    string errmsg = ex.Message.ToString();
                }

                
            }

            //取模數
            /*
            int l_iModVal = p_oNewTable.RowCount() % BATCH_COUNT;
            for (int i= l_iIndex; i< l_iIndex+l_iModVal; i++)
            {
                IXLRangeRow l_row = p_oNewTable.Row(i);
                //共39行
                l_oNewRow.SubpType = l_row.Cell(1).Value.ToString();    //傳票類別
                l_oNewRow.DebitAmount = l_row.Cell(2).Value.ToString();     //借方金額
                l_oNewRow.CreditAmount = l_row.Cell(3).Value.ToString();    //貸方金額
                l_oNewRow.CurrencyName = l_row.Cell(4).Value.ToString();      //幣別
                l_oNewRow.FiscalYear = l_row.Cell(5).Value.ToString();  //會計年度
                l_oNewRow.AccountPeriod = l_row.Cell(6).Value.ToString(); //會計期別
                l_oNewRow.GUID = Guid.Parse(l_row.Cell(7).Value.ToString());      //GUID
                l_oNewRow.SubpDate = l_row.Cell(8).Value.ToString();      //傳票日期
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
            */

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            

            /*
            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位
                
                l_oNewRow = new FA_FaJournal();
                //共39行
                l_oNewRow.SubpType = l_row.Cell(1).Value.ToString();    //傳票類別
                l_oNewRow.DebitAmount = l_row.Cell(2).Value.ToString();     //借方金額
                l_oNewRow.CreditAmount = l_row.Cell(3).Value.ToString();    //貸方金額
                l_oNewRow.CurrencyName = l_row.Cell(4).Value.ToString();      //幣別
                l_oNewRow.FiscalYear = l_row.Cell(5).Value.ToString();  //會計年度
                l_oNewRow.AccountPeriod = l_row.Cell(6).Value.ToString(); //會計期別
                l_oNewRow.GUID = Guid.Parse(l_row.Cell(7).Value.ToString());      //GUID
                l_oNewRow.SubpDate = l_row.Cell(8).Value.ToString();      //傳票日期
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


            //Log l_oLog = new Log();
            //l_oLog.LogCount = l_oLog.LogCount + 1;
            //db.Log.Add(l_oLog);
            try
            {
                l_iDBStatus = db.SaveChanges();
            }
            catch(DbEntityValidationException ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }
            */
        }


        public List<FA_FaJournal> FA_FaJournal_GetDataByMonthVal(string p_sVal)
        {
            List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oFaJournals;

        }

        public List<FA_FaJournal> FA_FaJournal_GetDataByMonthValPage(string p_sVal, int p_iPageing)
        {
            //分頁傳回Paging Data
            const int PAGE_COUNT = 50;
            int l_iShowRange = PAGE_COUNT * p_iPageing;

            //去除之前的範圍
            List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).OrderBy(m => m.Id).Skip(l_iShowRange).ToList();
            //再從剩下的傳回300筆
            List<FA_FaJournal> l_oLastFaJournals = l_oFaJournals.Take(PAGE_COUNT).ToList();

            //List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oLastFaJournals;

        }


        public Boolean FA_FaJournal_FindDataByMonthVal(string p_sVal)
        {
            List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            if(l_oFaJournals.Count() > 0)
            {
                return true;
            } else
            {
                return false;
            }
            
        }

        /*
        public void FA_FaJournal_DBCreate(DataTable p_oNewTable)
        {
            FA_FaJournal l_oNewRow = null;
            int l_iIndex = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            foreach(DataRow l_row in p_oNewTable.Rows)
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue;

                l_oNewRow = new FA_FaJournal();
                l_oNewRow.SubpType = l_row.Field<int>(0);
                l_oNewRow.DebitAmount = l_row.Field<int>(1);
                l_oNewRow.CreditAmount = l_row.Field<int>(2);
                l_oNewRow.CurrencyName = l_row.Field<string>(3);
                l_oNewRow.FiscalYear = l_row.Field<int>(4);
                l_oNewRow.AccountPeriod = l_row.Field<int>(5);


                db.FA_FaJournal.Add(l_oNewRow);
            }
            
            //Log l_oLog = new Log();
            //l_oLog.LogCount = l_oLog.LogCount + 1;
            //db.Log.Add(l_oLog);
            try
            {
                l_iDBStatus = db.SaveChanges();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }

        }
        */

        public void FA_FaJournal_DBDeleteByPeriod(string p_sPeriod)
        {
            
            List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sPeriod).ToList();
            foreach(FA_FaJournal l_oDelRow in l_oFaJournals)
            {
                db.FA_FaJournal.Remove(l_oDelRow);
            }
            //FA_FaJournal l_oDelRow = db.FA_FaJournal.Find(p_sPeriod);
            try
            {
                db.SaveChanges();

            } catch(Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }
            
        }

        public void FA_FaJournal_DBDeleteByID(int p_iRowID)
        {
            FA_FaJournal l_oDelRow = db.FA_FaJournal.Find(p_iRowID);
            db.FA_FaJournal.Remove(l_oDelRow);
            db.SaveChanges();
        }


        //20201218 CCL+ For Processing 區間日期Excel 資料庫處理 /////////////////////////////////////
        public DataSet FA_FaJournal_SqlGetDataListByOptions(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null; 

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = @"SELECT * FROM " + DB_NAME + @"
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

        public List<FA_FaJournal> FA_FaJournal_DataTableTo_FaJournalsList(DataSet p_oDataSet)
        {
            string l_sGuidStr = "";
            List<FA_FaJournal> l_oRtnList = new List<FA_FaJournal>();
            FA_FaJournal l_oTmpFARow = null;
    

            foreach(DataRow l_row in p_oDataSet.Tables[0].Rows)
            {
                //Guid l_oGuid = new Guid(l_row.Field<Guid>(7).ToString());
                //GuidConverter l_oCov = new GuidConverter();
                //Guid l_oGuid = (Guid)l_oCov.ConvertFromString(l_row.Field<string>(7));
                //l_sGuidStr = GUIDStringProcess.GuidProcess();

                l_oTmpFARow = new FA_FaJournal();
                l_oTmpFARow.Id = l_row.Field<int>(0);
                l_oTmpFARow.SubpType = l_row.Field<string>(1);
                l_oTmpFARow.DebitAmount = l_row.Field<string>(2);
                l_oTmpFARow.CreditAmount = l_row.Field<string>(3);
                l_oTmpFARow.CurrencyName = l_row.Field<string>(4);
                l_oTmpFARow.FiscalYear = l_row.Field<string>(5);
                l_oTmpFARow.AccountPeriod = l_row.Field<string>(6);
                l_oTmpFARow.GUID = l_row.Field<string>(7);
                l_oTmpFARow.SubpDate = l_row.Field<string>(8);
                l_oTmpFARow.SubpNo = l_row.Field<string>(9);
                l_oTmpFARow.SubpSummary = l_row.Field<string>(10);
                l_oTmpFARow.SubjectName = l_row.Field<string>(11);
                l_oTmpFARow.AccountNo = l_row.Field<string>(12);
                l_oTmpFARow.DetailAccountNo = l_row.Field<string>(13);
                l_oTmpFARow.DetailSubjectName = l_row.Field<string>(14);
                l_oTmpFARow.DepartNo = l_row.Field<string>(15);
                l_oTmpFARow.DepartName = l_row.Field<string>(16);
                l_oTmpFARow.ProjectNo = l_row.Field<string>(17);
                l_oTmpFARow.ProjectAbbr = l_row.Field<string>(18);
                l_oTmpFARow.ObjectCateg = l_row.Field<string>(19);
                l_oTmpFARow.ObjectNo = l_row.Field<string>(20);
                l_oTmpFARow.CurrencyNo = l_row.Field<string>(21);
                l_oTmpFARow.ExchangeRate = l_row.Field<string>(22);
                l_oTmpFARow.OriginCurrency = l_row.Field<string>(23);
                l_oTmpFARow.LocalCurrencyAmount = l_row.Field<string>(24);
                l_oTmpFARow.Spare1No = l_row.Field<string>(25);
                l_oTmpFARow.Spare1Abbr = l_row.Field<string>(26);
                l_oTmpFARow.Spare2No = l_row.Field<string>(27);
                l_oTmpFARow.Spare2Abbr = l_row.Field<string>(28);
                l_oTmpFARow.Spare3No = l_row.Field<string>(29);
                l_oTmpFARow.Spare3Abbr = l_row.Field<string>(30);
                l_oTmpFARow.Spare4No = l_row.Field<string>(31);
                l_oTmpFARow.Spare4Abbr = l_row.Field<string>(32);
                l_oTmpFARow.Spare5No = l_row.Field<string>(33);
                l_oTmpFARow.Spare5Abbr = l_row.Field<string>(34);
                l_oTmpFARow.Summary1 = l_row.Field<string>(35);
                l_oTmpFARow.AccountSubjects = l_row.Field<string>(36);
                l_oTmpFARow.Summary = l_row.Field<string>(37);
                l_oTmpFARow.Category = l_row.Field<string>(38);
                l_oTmpFARow.SubjectAlias = l_row.Field<string>(39);
                l_oTmpFARow.IsValid = l_row.Field<int>(40);
                l_oTmpFARow.CreateTime = l_row.Field<DateTime>(41);
                l_oTmpFARow.UpdateTime = l_row.Field<DateTime>(42);

                l_oRtnList.Add(l_oTmpFARow);
            }

            return l_oRtnList;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////

        //20201221 CCL+ Sql方式////////////////////////////////////////////////////////////
        public DataSet FA_FaJournal_SqlGetDataListByOptions2(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = @"SELECT * FROM " + DB_NAME + @"
                                     WHERE AccountPeriod = N'{0}' AND 
                                           SubpDate BETWEEN N'{1}' AND N'{2}' AND 
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

        ///////////////////////////////////////////////////////////////////////////////////


        //20201220 CCL+ Sql方式////////////////////////////////////////////////////////////        

        public List<FA_FaJournal> FA_FaJournal_SqlGetDataList(ForPaging p_oPaging, ForSearch p_oSearch)
        {
            List<FA_FaJournal> l_RtnDataList = new List<FA_FaJournal>();

            if(!string.IsNullOrEmpty(p_oSearch.m_sSearch))
            {
                SetMaxPaging(p_oPaging, p_oSearch);
                l_RtnDataList = FA_FaJournal_SqlGetAllDataList(p_oPaging, p_oSearch);
            } else
            {
                SetMaxPaging(p_oPaging);
                l_RtnDataList = FA_FaJournal_SqlGetAllDataList(p_oPaging);
            }

            return l_RtnDataList;
        }

        public void SetMaxPaging(ForPaging p_oPaging)
        {
            int l_iRow = 0;
            int l_RowCount = 0;

            SqlCommand l_oSqlCmdObj = null;
            SqlDataReader l_oSqlDRObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            

            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = $@" SELECT COUNT(*) AS ROWCOUNT FROM " + DB_NAME +
                                  @" WHERE IsValid = 1 ;";

            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();


            l_sExeSqlCmd = l_sSelSqlCmd;

            try
            {
                l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
                l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                l_oSqlDRObj = l_oSqlCmdObj.ExecuteReader();

                while(l_oSqlDRObj.Read())
                {
                    //l_iRow = ;
                    if (!l_oSqlDRObj["ROWCOUNT"].Equals(DBNull.Value))
                    {
                        l_RowCount = (int)l_oSqlDRObj["ROWCOUNT"];
                    }
                }

            } catch(Exception ex)
            {
                string errmsg = ex.Message.ToString();
                
                Trace.WriteLine("Err: " + l_RowCount);
            }
            finally
            {
                l_oSqlConn.Close();
            }
            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            p_oPaging.m_iMaxPage = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(l_RowCount)
                                  / p_oPaging.m_iItemNum));

            p_oPaging.SetRightPage();
        }

        public void SetMaxPaging(ForPaging p_oPaging, ForSearch p_oSearch)
        {
            int l_iRow = 0;
            int l_RowCount = 0;

            SqlCommand l_oSqlCmdObj = null;
            SqlDataReader l_oSqlDRObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = $@" SELECT COUNT(*) AS ROWCOUNT FROM " + DB_NAME +
                                  @" WHERE IsValid = 1 AND SubpDate BETWEEN N'{0}' AND N'{1}' ;";

            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            if(!string.IsNullOrEmpty(p_oSearch.m_sSearch))
            {
                if(p_oSearch.m_iSearchTokenCount > 0)
                {
                l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                             p_oSearch.m_oSearchList[0],
                                             p_oSearch.m_oSearchList[1]);
                } else
                {
                    l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                             p_oSearch.m_sSearch,
                                             p_oSearch.m_sSearch);
                }

            }
            
            

            try
            {
                l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
                l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                l_oSqlDRObj = l_oSqlCmdObj.ExecuteReader();

                while (l_oSqlDRObj.Read())
                {
                    //l_iRow = ;
                    if (!l_oSqlDRObj["ROWCOUNT"].Equals(DBNull.Value))
                    {
                        l_RowCount = (int)l_oSqlDRObj["ROWCOUNT"];
                    }
                }

            }
            catch (Exception ex)
            {
                string errmsg = ex.Message.ToString();

                Trace.WriteLine("Err: " + l_RowCount);
            }
            finally
            {
                l_oSqlConn.Close();
            }
            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            p_oPaging.m_iMaxPage = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(l_RowCount)
                                  / p_oPaging.m_iItemNum));

            p_oPaging.SetRightPage();
        }

        public List<FA_FaJournal> FA_FaJournal_SqlGetAllDataList(ForPaging p_oPaging, ForSearch p_oSearch)
        {
            List<FA_FaJournal> l_RtnDataList = new List<FA_FaJournal>();

            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = $@" SELECT *  FROM " +
                      @"(SELECT row_number() OVER(ORDER BY Id) AS sort,* FROM " + DB_NAME +
                      @" WHERE SubpDate BETWEEN N'{0}' AND N'{1}' " +
                      @" ) m " +
                      @" WHERE m.sort BETWEEN N'{2}' AND N'{3}'  IsValid = 1  ;";



            return l_RtnDataList;
        }

        public List<FA_FaJournal> FA_FaJournal_SqlGetAllDataList(ForPaging p_oPaging)
        {
            List<FA_FaJournal> l_RtnDataList = new List<FA_FaJournal>();
            FA_FaJournal l_oTmpFARow = null;

            int l_iRow = 0;
            int l_RowCount = 0;

            SqlCommand l_oSqlCmdObj = null;
            SqlDataReader l_oSqlDRObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = $@" SELECT *  FROM " + 
                      @"(SELECT row_number() OVER(ORDER BY Id) AS sort,* FROM " + DB_NAME + @" ) m " +
                      @" WHERE m.sort BETWEEN N'{0}' AND N'{1}'  IsValid = 1  ;";

            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //if (!string.IsNullOrEmpty(p_oSearch.m_sSearch))
            //{
            //    if (p_oSearch.m_iSearchTokenCount > 0)
            //    {
                    l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                                                 (p_oPaging.m_iNowPage - 1) * p_oPaging.m_iItemNum + 1,
                                                 p_oPaging.m_iNowPage * p_oPaging.m_iItemNum);
            //    }

            //}

            try
            {
                l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
                l_oSqlCmdObj = new SqlCommand(l_sExeSqlCmd, l_oSqlConn);
                l_oSqlDRObj = l_oSqlCmdObj.ExecuteReader();

                while (l_oSqlDRObj.Read())
                {
                    //l_iRow = ;
                    l_oTmpFARow = new FA_FaJournal();
                    if (!l_oSqlDRObj[0].Equals(DBNull.Value))
                    {
                        l_oTmpFARow.Id = (int)l_oSqlDRObj[0];
                    }

                    l_RtnDataList.Add(l_oTmpFARow);
                }

            }
            catch (Exception ex)
            {
                string errmsg = ex.Message.ToString();

                Trace.WriteLine("Err: " + l_RowCount);
            }
            finally
            {
                l_oSqlConn.Close();
            }
            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

            return l_RtnDataList;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////

        #endregion
    }
}