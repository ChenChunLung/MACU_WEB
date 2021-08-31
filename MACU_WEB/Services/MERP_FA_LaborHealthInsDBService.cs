using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data.SqlClient;
using System.Data;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Web.Configuration;
using MACU_WEB.BIServices;
using MACU_WEB.Models._Base;
using System.ComponentModel;
using System.Globalization;
using ClosedXML.Excel;
using System.Web.Mvc;


namespace MACU_WEB.Services
{
    public class MERP_FA_LaborHealthInsDBService
    {

        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_LaborHealthIns";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_LaborHealthIns

        //public FA_LaborHealthIns FA_LHIns_GetDBService()
        //{
        //    return db.FA_LaborHealthIns;
        //}


        //1.
        public IEnumerable<FA_LaborHealthIns> FA_LaborHealthIns_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_LaborHealthIns
                             where DataVals.IsValid == 1
                             select DataVals;

            return l_oRtnData;
            //return db.FA_LaborHealthIns.ToList();
        }

        //2.
        public FA_LaborHealthIns FA_LaborHealthIns_GetDataById(int p_iId)
        {


            FA_LaborHealthIns l_oFindItem = db.FA_LaborHealthIns.
                    Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }

        //3.
        public void FA_LaborHealthIns_DBCreate(
                                 string p_sDepartName, string p_sPlusInsCompany, string p_sCoding, int p_iLaborIns,
                                 int p_iHealthIns, string p_sGroupIns, string p_sJobTitle, string p_sOnJobDate, 
                                 string p_sResignDate, string p_sSeniority, string p_sKeepSecret, string p_sSalary, 
                                 string p_sMemberName )
        {
            FA_LaborHealthIns l_oNewItem = new FA_LaborHealthIns();

            l_oNewItem.DepartName = p_sDepartName;
            l_oNewItem.PlusInsCompany = p_sPlusInsCompany;
            l_oNewItem.Coding = p_sCoding;
            l_oNewItem.LaborIns = p_iLaborIns.ToString();
            l_oNewItem.HealthIns = p_iHealthIns.ToString();
            l_oNewItem.GroupIns = p_sGroupIns;
            l_oNewItem.JobTitle = p_sJobTitle;
            l_oNewItem.OnJobDate = p_sOnJobDate;
            l_oNewItem.ResignDate = p_sResignDate;
            l_oNewItem.Seniority = p_sSeniority;
            l_oNewItem.KeepSecret = p_sKeepSecret;
            l_oNewItem.Salary = p_sSalary;
            l_oNewItem.MemberName = p_sMemberName;
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_LaborHealthIns.Add(l_oNewItem);

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


        //4.
        public void FA_LaborHealthIns_DBDeleteByID(int p_iItemID)
        {
            FA_LaborHealthIns l_oDelItem = db.FA_LaborHealthIns.Find(p_iItemID);
            db.FA_LaborHealthIns.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //5.
        public void FA_LaborHealthIns_DBUpdate(int p_iItemID, FA_LaborHealthIns p_oNewUpdItem)
        {
            FA_LaborHealthIns l_oUpdItem = db.FA_LaborHealthIns.Find(p_iItemID);

            l_oUpdItem.DepartName = p_oNewUpdItem.DepartName;
            l_oUpdItem.PlusInsCompany = p_oNewUpdItem.PlusInsCompany;
            l_oUpdItem.Coding = p_oNewUpdItem.Coding;
            l_oUpdItem.LaborIns = p_oNewUpdItem.LaborIns;
            l_oUpdItem.HealthIns = p_oNewUpdItem.HealthIns;
            l_oUpdItem.GroupIns = p_oNewUpdItem.GroupIns;
            l_oUpdItem.JobTitle = p_oNewUpdItem.JobTitle;
            l_oUpdItem.OnJobDate = p_oNewUpdItem.OnJobDate;
            l_oUpdItem.ResignDate = p_oNewUpdItem.ResignDate;
            l_oUpdItem.Seniority = p_oNewUpdItem.Seniority;
            l_oUpdItem.KeepSecret = p_oNewUpdItem.KeepSecret;
            l_oUpdItem.Salary = p_oNewUpdItem.Salary;
            l_oUpdItem.MemberName = p_oNewUpdItem.MemberName;
            l_oUpdItem.IsValid = 1;           
            l_oUpdItem.UpdateTime = DateTime.Now;



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

        //6.
        public IEnumerable<FA_LaborHealthIns>  FA_LaborHealthIns_GetDataByYearMon(int p_iYear, int p_iMonth)
        {
            var l_oRtnData = from DataVals in db.FA_LaborHealthIns
                             where DataVals.IsValid == 1 && DataVals.DataYear == p_iYear.ToString() 
                              && DataVals.DataMonth == p_iMonth.ToString()
                             select DataVals;

            return l_oRtnData;

        }

        public Boolean FA_LaborHealthIns_ChkDataByYearMon(int p_iYear, int p_iMonth)
        {
            
            bool l_bIsExist = db.FA_LaborHealthIns.ToList().Exists(m => (m.IsValid == 1) &&
                                                      (m.DataYear == p_iYear.ToString()) &&
                                                      (m.DataMonth == p_iMonth.ToString()));

            return l_bIsExist;
           
        }

        public void FA_LaborHealthIns_DBDeleteByYearMon(int p_iYear, int p_iMonth)
        {
            List<FA_LaborHealthIns> l_oDelItems = db.FA_LaborHealthIns.ToList()
                                          .Where(m => (m.IsValid == 1) &&
                                          (m.DataYear == p_iYear.ToString()) &&
                                          (m.DataMonth == p_iMonth.ToString()) ).ToList();
            foreach(FA_LaborHealthIns Item in l_oDelItems)
            {
                db.FA_LaborHealthIns.Remove(Item);
            }
           
            db.SaveChanges();
        }

        public void FA_LaborHealthIns_SqlDBDeleteByYearMon(int p_iYear, int p_iMonth)
        {
            int l_iRtnDelCount = 0;
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            //SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sDelSqlCmd = @"DELETE FROM " + TB_NAME + @" WHERE 
                                   DataYear = N'{0}' AND DataMonth = N'{1}'  AND IsValid = 1";

            l_sExeSqlCmd = string.Format(l_sDelSqlCmd,
                                         p_iYear.ToString(),
                                         p_iMonth.ToString() );

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
                    throw new ArgumentException("FA_LaborHealthIns資料刪除發生錯誤!!");
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


        public void FA_LaborHealthIns_SqlDBCreate(IXLTable p_oNewTable, int p_iYear, int p_iMonth)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
                                   DepartName, PlusInsCompany, Coding, LaborIns, 
                                   HealthIns, GroupIns, JobTitle, OnJobDate, 
                                   ResignDate, Seniority, KeepSecret, Salary,  
                                   MemberName, DataYear, DataMonth, 
                                   CreateTime, UpdateTime)
                                   VALUES (
                                     N'{0}', N'{1}', N'{2}', N'{3}', 
                                     N'{4}', N'{5}', N'{6}', N'{7}', 
                                     N'{8}', N'{9}', N'{10}', N'{11}', 
                                     N'{12}', N'{13}', N'{14}', N'{15}', 
                                     N'{16}'
                                   )";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_LaborHealthIns l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //DateTimeConverter l_oDT = new DateTimeConverter();           

            int l_iYear = p_iYear;
            int l_iMonth = p_iMonth;

            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //共41行
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                ///l_sSubpDate = l_row.Cell(1).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                ///l_sSubpDate = l_sSubpDate.Substring(0, l_sSubpDate.IndexOf(' '));

                ///if (l_iIndex == 2)
                ///{
                    //根據第一列的傳票日期,取出年,月
                ///    string l_sTmpStr = DateStringProcess.Del_MonthDayZero(l_sSubpDate, "/", ""); //最後參數為"",代表不替換
                ///    l_iYear = DateStringProcess.m_Year;
                ///    l_iMonth = DateStringProcess.m_Month;
                ///}


                l_sExeSqlCmd = string.Format(l_sInsSqlCmd,
                                                l_row.Cell(1).Value.ToString(),
                                                l_row.Cell(2).Value.ToString(),
                                                l_row.Cell(3).Value.ToString(),
                                                l_row.Cell(4).Value.ToString(),
                                                l_row.Cell(5).Value.ToString(),
                                                l_row.Cell(6).Value.ToString(),
                                                l_row.Cell(7).Value.ToString(),
                                                l_row.Cell(8).Value.ToString(),
                                                l_row.Cell(9).Value.ToString(),
                                                l_row.Cell(10).Value.ToString(),
                                                l_row.Cell(11).Value.ToString(),
                                                l_row.Cell(12).Value.ToString(),
                                                l_row.Cell(13).Value.ToString(),
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
                        throw new ArgumentException("FA_LaborHealthIns資料新增發生錯誤!!");
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


        public List<FA_LaborHealthIns> FA_LaborHealthIns_GetDataByYearMonthPage(string p_sYear,
                                                          string p_sMonth, int p_iPageing)
        {
            //分頁傳回Paging Data
            const int PAGE_COUNT = 50;
            int l_iShowRange = PAGE_COUNT * p_iPageing;

            //去除之前的範圍
            List<FA_LaborHealthIns> l_oFALaborHealthIns = db.FA_LaborHealthIns.Where(m => (m.DataYear == p_sYear) && 
                                                                                     (m.DataMonth == p_sMonth) && 
                                                                                     (m.IsValid == 1) )
                                                                        .OrderBy(m => m.Id).Skip(l_iShowRange).ToList();
            //再從剩下的傳回300筆
            List<FA_LaborHealthIns> l_oLastFALaborHealthIns = l_oFALaborHealthIns.Take(PAGE_COUNT).ToList();

            //List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oLastFALaborHealthIns;

        }


        #endregion


    }
}