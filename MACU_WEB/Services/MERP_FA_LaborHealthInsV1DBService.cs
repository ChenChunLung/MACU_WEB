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
using MACU_WEB.Areas.MERP_TCF000.ViewModels;

namespace MACU_WEB.Services
{
    public class MERP_FA_LaborHealthInsV1DBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_LaborHealthInsV1";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_LaborHealthInsV1

        //1.
        public IEnumerable<FA_LaborHealthInsV1> FA_LaborHealthInsV1_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_LaborHealthInsV1
                             where DataVals.IsValid == 1
                             select DataVals;

            return l_oRtnData;
            //return db.FA_LaborHealthIns.ToList();
        }

        //2.
        public FA_LaborHealthInsV1 FA_LaborHealthInsV1_GetDataById(int p_iId)
        {


            FA_LaborHealthInsV1 l_oFindItem = db.FA_LaborHealthInsV1.
                    Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }

        //3.
        public void FA_LaborHealthInsV1_DBCreate(
                                 string p_sDepartName, string p_sPlusInsCompany, string p_sMemberName, int p_iLaborIns,
                                 int p_iHealthIns, string p_sDependents, string p_sOnJobDate,
                                 string p_sResignDate, int p_sLHInsType
                                 )
        {
            FA_LaborHealthInsV1 l_oNewItem = new FA_LaborHealthInsV1();

            l_oNewItem.DepartName = p_sDepartName;
            l_oNewItem.PlusInsCompany = p_sPlusInsCompany;
            l_oNewItem.MemberName = p_sMemberName;
            l_oNewItem.LaborIns = p_iLaborIns.ToString();
            l_oNewItem.HealthIns = p_iHealthIns.ToString();
            l_oNewItem.Dependents = p_sDependents;
            l_oNewItem.OnJobDate = p_sOnJobDate;
            l_oNewItem.ResignDate = p_sResignDate;
            l_oNewItem.LHInsType = p_sLHInsType;
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_LaborHealthInsV1.Add(l_oNewItem);

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
        public void FA_LaborHealthInsV1_DBDeleteByID(int p_iItemID)
        {
            FA_LaborHealthInsV1 l_oDelItem = db.FA_LaborHealthInsV1.Find(p_iItemID);
            db.FA_LaborHealthInsV1.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //5.
        public void FA_LaborHealthInsV1_DBUpdate(int p_iItemID, FA_LaborHealthInsV1 p_oNewUpdItem)
        {
            FA_LaborHealthInsV1 l_oUpdItem = db.FA_LaborHealthInsV1.Find(p_iItemID);

            l_oUpdItem.DepartName = p_oNewUpdItem.DepartName;
            l_oUpdItem.PlusInsCompany = p_oNewUpdItem.PlusInsCompany;
            l_oUpdItem.MemberName = p_oNewUpdItem.MemberName;
            l_oUpdItem.LaborIns = p_oNewUpdItem.LaborIns;
            l_oUpdItem.HealthIns = p_oNewUpdItem.HealthIns;
            l_oUpdItem.Dependents = p_oNewUpdItem.Dependents;
            l_oUpdItem.OnJobDate = p_oNewUpdItem.OnJobDate;
            l_oUpdItem.ResignDate = p_oNewUpdItem.ResignDate;
            l_oUpdItem.LHInsType = p_oNewUpdItem.LHInsType;
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
        public IEnumerable<FA_LaborHealthInsV1> FA_LaborHealthInsV1_GetDataByYearMon(int p_iYear, int p_iMonth)
        {
            var l_oRtnData = from DataVals in db.FA_LaborHealthInsV1
                             where DataVals.IsValid == 1 && DataVals.DataYear == p_iYear.ToString()
                              && DataVals.DataMonth == p_iMonth.ToString()
                             select DataVals;

            return l_oRtnData;
            
        }

        public Boolean FA_LaborHealthInsV1_ChkDataByYearMon(int p_iYear, int p_iMonth)
        {

            bool l_bIsExist = db.FA_LaborHealthInsV1.ToList().Exists(m => (m.IsValid == 1) &&
                                                      (m.DataYear == p_iYear.ToString()) &&
                                                      (m.DataMonth == p_iMonth.ToString()));

            return l_bIsExist;

        }

        public void FA_LaborHealthInsV1_DBDeleteByYearMon(int p_iYear, int p_iMonth)
        {
            List<FA_LaborHealthInsV1> l_oDelItems = db.FA_LaborHealthInsV1.ToList()
                                          .Where(m => (m.IsValid == 1) &&
                                          (m.DataYear == p_iYear.ToString()) &&
                                          (m.DataMonth == p_iMonth.ToString())).ToList();
            foreach (FA_LaborHealthInsV1 Item in l_oDelItems)
            {
                db.FA_LaborHealthInsV1.Remove(Item);
            }

            db.SaveChanges();
        }

        public void FA_LaborHealthInsV1_SqlDBDeleteByYearMon(int p_iYear, int p_iMonth)
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
                                         p_iMonth.ToString());

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
                    throw new ArgumentException("FA_LaborHealthInsV1資料刪除發生錯誤!!");
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


        public void FA_LaborHealthInsV1_SqlDBCreate(IXLTable p_oNewTable, int p_iYear, int p_iMonth)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();
           

            string l_sExeSqlCmd = "";
            string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
                                   DepartName, PlusInsCompany, MemberName, LaborIns, 
                                   HealthIns, Dependents, OnJobDate, 
                                   ResignDate, DataYear, DataMonth, 
                                   LHInsType, 
                                   CreateTime, UpdateTime)
                                   VALUES (
                                     N'{0}', N'{1}', N'{2}', N'{3}', 
                                     N'{4}', N'{5}', N'{6}', N'{7}', 
                                     N'{8}', N'{9}', N'{10}', N'{11}', 
                                     N'{12}'
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
            String l_sOnJobDate = "";
            String l_sResignDate = "";
            int l_iLHInsType = 0;

            foreach (IXLRangeRow l_row in p_oNewTable.Rows())
            {
                ++l_iIndex;
                if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //從Excel匯進來的要去掉分:秒
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                l_sOnJobDate = l_row.Cell(7).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                if(!string.IsNullOrEmpty(l_sOnJobDate))
                    l_sOnJobDate = l_sOnJobDate.Substring(0, l_sOnJobDate.IndexOf(' '));

                l_sResignDate = l_row.Cell(8).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                if (!string.IsNullOrEmpty(l_sResignDate))
                    l_sResignDate = l_sResignDate.Substring(0, l_sResignDate.IndexOf(' '));

                //20210125 CCl+ 加上備註
                if(!string.IsNullOrEmpty(l_row.Cell(9).Value.ToString()))
                {
                    l_iLHInsType = Convert.ToInt32(l_row.Cell(9).Value.ToString());
                } else
                {
                    l_iLHInsType = 0;
                }
                
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
                                                l_sOnJobDate,
                                                l_sResignDate,                                                                                                                                       
                                                l_iYear.ToString(),
                                                l_iMonth.ToString(),
                                                l_iLHInsType,
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
                        throw new ArgumentException("FA_LaborHealthInsV1資料新增發生錯誤!!");
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


        public List<FA_LaborHealthInsV1> FA_LaborHealthInsV1_GetDataByYearMonthPage(string p_sYear,
                                                          string p_sMonth, int p_iPageing)
        {
            //分頁傳回Paging Data
            const int PAGE_COUNT = 50;
            int l_iShowRange = PAGE_COUNT * p_iPageing;

            //去除之前的範圍
            List<FA_LaborHealthInsV1> l_oFALaborHealthIns = db.FA_LaborHealthInsV1.Where(m => (m.DataYear == p_sYear) &&
                                                                                     (m.DataMonth == p_sMonth) &&
                                                                                     (m.IsValid == 1))
                                                                        .OrderBy(m => m.Id).Skip(l_iShowRange).ToList();
            //再從剩下的傳回300筆
            List<FA_LaborHealthInsV1> l_oLastFALaborHealthIns = l_oFALaborHealthIns.Take(PAGE_COUNT).ToList();

            //List<FA_FaJournal> l_oFaJournals = db.FA_FaJournal.Where(m => m.AccountPeriod == p_sVal).ToList();

            return l_oLastFALaborHealthIns;

        }

        public List<FA_LaborHealthInsV1> FA_LaborHealthInsV1_DataTableTo_FALHInsV1List(DataSet p_oDataSet)
        {
            string l_sGuidStr = "";
            List<FA_LaborHealthInsV1> l_oRtnList = new List<FA_LaborHealthInsV1>();
            FA_LaborHealthInsV1 l_oTmpFARow = null;


            foreach (DataRow l_row in p_oDataSet.Tables[0].Rows)
            {
                //Guid l_oGuid = new Guid(l_row.Field<Guid>(7).ToString());
                //GuidConverter l_oCov = new GuidConverter();
                //Guid l_oGuid = (Guid)l_oCov.ConvertFromString(l_row.Field<string>(7));
                //l_sGuidStr = GUIDStringProcess.GuidProcess();

                l_oTmpFARow = new FA_LaborHealthInsV1();
                l_oTmpFARow.Id = l_row.Field<int>(0);
                l_oTmpFARow.DepartName = l_row.Field<string>(1);
                l_oTmpFARow.PlusInsCompany = l_row.Field<string>(2);
                l_oTmpFARow.MemberName = l_row.Field<string>(3);
                l_oTmpFARow.LaborIns = l_row.Field<string>(4);
                l_oTmpFARow.HealthIns = l_row.Field<string>(5);
                l_oTmpFARow.Dependents = l_row.Field<string>(6);
                l_oTmpFARow.OnJobDate = l_row.Field<string>(7);
                l_oTmpFARow.ResignDate = l_row.Field<string>(8);
                l_oTmpFARow.DataYear = l_row.Field<string>(9);
                l_oTmpFARow.DataMonth = l_row.Field<string>(10);
                l_oTmpFARow.LHInsType = l_row.Field<int>(11);
                l_oTmpFARow.IsValid = l_row.Field<int>(12);
                l_oTmpFARow.CreateTime = l_row.Field<DateTime>(13);
                l_oTmpFARow.UpdateTime = l_row.Field<DateTime>(14);

                l_oRtnList.Add(l_oTmpFARow);
            }

            return l_oRtnList;
        }


        //20210114 CCL+ Sql方式////////////////////////////////////////////////////////////
        public DataSet FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(MERP_TCF004_JournalsOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = "";


            l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                    WHERE  DataYear = N'{0}' AND DataMonth = N'{1}' 
                                            AND IsValid = 1 
                                     ORDER BY PlusInsCompany, DepartName, Id
                                  ";


            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_LaborHealthInsV1 l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();


            l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                         p_oOption.m_sDataYear,
                         p_oOption.m_sDataMonth);



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
                    throw new ArgumentException("FA_LaborHealthInsV1資料新增發生錯誤!!");
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



        //20210114 CCL+ Sql方式////////////////////////////////////////////////////////////
        public DataSet FA_LaborHealthInsV1_SqlGetDataListByOptions(MERP_TCF004_JournalsOptions p_oOption)
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = "";

            if (!string.IsNullOrEmpty(p_oOption.m_sResignDate))
            {            

                //20210118 CCl-
                //l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                //                    WHERE  DataYear = N'{0}' AND DataMonth = N'{1}' AND 
                //                           CONVERT(DATE, OnJobDate) >= CONVERT(DATE, N'{2}') 
                //                           AND 
                //                           ( 
                //                               CASE  
                //                                    WHEN ResignDate = '' THEN                                                        
                //                                        CONVERT(DATE, N'9999/12/30')
                //                                    ELSE                                                
                //                                        CONVERT(DATE, ResignDate) 
                //                               END                                                    
                //                           ) <= CONVERT(DATE, N'{3}') 
                //                            AND IsValid = 1
                //                     ORDER BY Id
                //                  ";

                l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                    WHERE  DataYear = N'{0}' AND DataMonth = N'{1}' AND 
                                           CONVERT(DATE, OnJobDate) >= CONVERT(DATE, N'{2}') 
                                           AND 
                                           ( 
                                               CASE  
                                                    WHEN ResignDate = '' THEN                                                        
                                                        CONVERT(DATE, N'9999/12/30')
                                                    ELSE                                                
                                                        CONVERT(DATE, ResignDate) 
                                               END                                                    
                                           ) <= CONVERT(DATE, N'{3}') 
                                            AND IsValid = 1 
                                     ORDER BY PlusInsCompany, DepartName, Id
                                  ";
            } else
            {
                //跑出所有未離職的 && 到職日符合條件的
                //l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                //                WHERE  DataYear = N'{0}' AND DataMonth = N'{1}' AND 
                //                       CONVERT(DATE, OnJobDate) >= CONVERT(DATE, N'{2}')  
                //                        AND ResignDate = ''  
                //                        AND IsValid = 1
                //                 ORDER BY Id
                //              ";

                l_sSelSqlCmd = @"SELECT * FROM " + TB_NAME + @"
                                WHERE  DataYear = N'{0}' AND DataMonth = N'{1}' AND 
                                       CONVERT(DATE, OnJobDate) >= CONVERT(DATE, N'{2}')  
                                        AND ResignDate = ''  
                                        AND IsValid = 1                                  
                                 ORDER BY PlusInsCompany, DepartName, Id
                              ";

            }




            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            FA_LaborHealthInsV1 l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            if (!string.IsNullOrEmpty(p_oOption.m_sResignDate))
            {
                l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                             p_oOption.m_sDataYear,
                             p_oOption.m_sDataMonth,
                             p_oOption.m_sOnJobDate,
                             p_oOption.m_sResignDate);

            } else
            {
                l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
                             p_oOption.m_sDataYear,
                             p_oOption.m_sDataMonth,
                             p_oOption.m_sOnJobDate);

            }

 

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
                    throw new ArgumentException("FA_LaborHealthInsV1資料新增發生錯誤!!");
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


        #endregion

    }
}