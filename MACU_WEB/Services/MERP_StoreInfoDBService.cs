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
    public class MERP_StoreInfoDBService
    {

        private const int BATCH_COUNT = 100;
        private const string SRC_TB_NAME = "StoreInfo";
        private const string TB_NAME = "StoreInfo";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region StoreInfo
        public List<StoreInfo> StoreInfo_GetDataList()
        {
            var l_oRtnData = from DataVals in db.StoreInfo
                             where DataVals.Disabled == false
                             select DataVals;

            //return (List<StoreInfo>)l_oRtnData.ToList();
            return l_oRtnData.ToList();
            //return db.StoreInfo.ToList();
        }

        //20210107 CCL+ 
        public IEnumerable<dynamic> StoreInfo_GetDataGroupByBranch()
        {
            //以BranchArea_ID來分群
            var l_oRtnShop = from DataVals in db.StoreInfo
                             group DataVals by DataVals.BranchArea_ID into grp
                             select new
                             {
                                 BranchID = grp.Key,
                                 BranchShopCount = grp.Count(),
                                 GrpObj = grp
                             };

            /*
            foreach (var AllGroup in l_oRtnShop)
            {
                Trace.WriteLine("BranchID: {0}; ", AllGroup.BranchID);
                Trace.WriteLine("BranchShopCount: {0}; ", AllGroup.BranchShopCount.ToString());
                foreach (var grpobj in AllGroup.GrpObj)
                {
                    Trace.WriteLine("   SID: {0}", grpobj.SID);
                    Trace.WriteLine("   Name: {1}", grpobj.Name);
                }
            }
            */

            //return l_oRtnShop;

            //傳回匿名型別List
            return (IEnumerable<dynamic>)l_oRtnShop.ToList();
        }


        public IEnumerable<dynamic> StoreInfo_GetDataGroupBySID()
        {
            //以SID開頭來分群 N->北, S->南, C->中
            var l_oRtnShop = from DataVals in db.StoreInfo
                             group DataVals by DataVals.SID.Substring(0,1) into grp
                             select new
                             {
                                 SIDTopChar = grp.Key,
                                 GroupShopsCount = grp.Count(),
                                 GrpObj = grp
                             };

            /*
            foreach (var AllGroup in l_oRtnShop)
            {
                Trace.WriteLine("SIDTopChar: {0}; ", AllGroup.SIDTopChar);
                Trace.WriteLine("GroupShopsCount: {0}; ", AllGroup.GroupShopsCount.ToString());
                foreach (var grpobj in AllGroup.GrpObj)
                {
                    Trace.WriteLine("   SID: {0}", grpobj.SID);
                    Trace.WriteLine("   Name: {1}", grpobj.Name);
                }
            }
            */

            //傳回匿名型別List
            return (IEnumerable<dynamic>)l_oRtnShop.ToList();
        }


        public StoreInfo StoreInfo_GetDataById(int p_iId)
        {

            StoreInfo l_oFindItem = db.StoreInfo.Where(m => m.Disabled == false)
                                        .ToList().Find( m => m.Id == p_iId);
            return l_oFindItem;
        }

        public DataSet StoreInfo_GetDataListFromWebHQCenter()
        {
            DataSet l_oRtnDataSet = null;

            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["WebHQCenterSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }


            string l_sExeSqlCmd = "";
            string l_sSelSqlCmd = @"SELECT ID, Name,Kind, Memo,Group_ID, BranchArea_ID, OrderArea_ID,
                                     PrePaidScaleGroup_ID, ConveyKind_ID, BeginDate, TelPhone,
                                     FaxPhone, CellPhone, Contact, Address, Balance, OptimisticLockField,
                                     GCRecord, MobileStoreID, CreateDate, UpdateDate 
                                     FROM " + SRC_TB_NAME + @"
                                     WHERE  Disabled = N'0'  
                                     ORDER BY ID
                                   ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            StoreInfo l_oNewRow = null;
            int l_iIndex = 0, l_iRtnCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            l_sExeSqlCmd = l_sSelSqlCmd;
            //l_sExeSqlCmd = string.Format(l_sSelSqlCmd,
            //                                p_oOption.m_sAccountPeriod,
            //                                p_oOption.m_sStartDate,
            //                                p_oOption.m_sEndDate);

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

            //return db.StoreInfo.ToList();
        }


        public bool StoreInfo_SqlDBChkUpdate()
        {

            List<StoreInfo> l_oStoreItem = StoreInfo_GetDataList();
            if (l_oStoreItem != null && l_oStoreItem.Count() > 0)
            {
                //刪除舊的,加新的
                StoreInfo_SqlDBDelete();
                //從WebHQCenter抓資訊存入LocalDB
                DataSet L_oTmpDS = StoreInfo_GetDataListFromWebHQCenter();
                StoreInfo_SqlDBCreate(L_oTmpDS);

                return true;
            } else
            {
                //從WebHQCenter抓資訊存入LocalDB
                DataSet L_oTmpDS = StoreInfo_GetDataListFromWebHQCenter();
                StoreInfo_SqlDBCreate(L_oTmpDS);
            }

            return false;
        }

        public void StoreInfo_SqlDBDelete()
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sDelSqlCmd = @"DELETE FROM " + TB_NAME + @" 
                                     WHERE Disabled = N'0'
                                    ";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            StoreInfo l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //DateTimeConverter l_oDT = new DateTimeConverter();
            String l_sSubpDate = "";
            int l_iYear = 0, l_iMonth = 0;

          
            l_sExeSqlCmd = l_sDelSqlCmd;


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
                    throw new ArgumentException("StoreInfo資料新增發生錯誤!!");
                }


            }
            catch (Exception ex)
            {
                l_oSqlTrans.Rollback();
                l_oSqlConn.Close();
                Trace.WriteLine("Err: " + l_iRtnInsCount);
            }


            //成功寫入
            l_oSqlTrans.Commit();
            //成功關閉DB Conn          
            l_oSqlConn.Close();

            sw.Stop();
            Trace.WriteLine(sw.ElapsedMilliseconds);

        }



        public void StoreInfo_SqlDBCreate(DataSet p_oSrcDataSet)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();


            string l_sExeSqlCmd = "";
            string l_sInsSqlCmd = @"INSERT INTO " + TB_NAME + @"(
                                   SID, Name, Kind, Memo, 
                                   Group_ID, BranchArea_ID, OrderArea_ID, PrePaidScaleGroup_ID, 
                                   ConveyKind_ID, BeginDate, TelPhone, FaxPhone,  
                                   CellPhone, Contact, Address, Balance,
                                   OptimisticLockField, GCRecord, MobileStoreID, ManageKind,
                                   CreateDate, UpdateDate)
                                   VALUES (
                                     N'{0}', N'{1}', {2}, N'{3}', 
                                     {4}, N'{5}', N'{6}', {7}, 
                                     {8}, N'{9}', N'{10}', N'{11}', 
                                     N'{12}', N'{13}', N'{14}', {15},
                                     {16}, {17}, N'{18}', '',
                                     N'{19}', N'{20}'
                                   )";

            //l_sInsSqlCmd = string.Format(l_sInsSqlCmd, );

            //SqlDataAdapter l_oSqlAd = new SqlDataAdapter();

            ////////// DB Save /////////////////////////////////////////////////////
            StoreInfo l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從Excel匯入的IXLTable轉成的Datable存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            //DateTimeConverter l_oDT = new DateTimeConverter();
            String l_sSubpDate = "";
            int l_iYear = 0, l_iMonth = 0;

            foreach (DataRow l_row in p_oSrcDataSet.Tables[0].Rows)
            {
                ++l_iIndex;
                //if (l_iIndex == 1) continue; //跳過第一行中文欄位

                //共41行
                //Guid.Parse(l_row.Cell(7).Value.ToString()),
                ///l_sSubpDate = l_row.Cell(1).Value.ToString(); //去掉後面的" 上午 HH:MM:SS"
                ///l_sSubpDate = l_sSubpDate.Substring(0, l_sSubpDate.IndexOf(' '));
                //var l_oDatetime = Convert.IsDBNull(l_row["BeginDate"]) ? l_row.Field<DateTime?>("BeginDate") : Convert.ToDateTime(l_row["BeginDate"]);

                l_sExeSqlCmd = string.Format(l_sInsSqlCmd,
                                                l_row["ID"],
                                                l_row["Name"],
                                                l_row["Kind"] != null ? l_row["Kind"] : 0,
                                                l_row["Memo"],
                                                l_row["Group_ID"] != null ? 0 : l_row["Group_ID"],
                                                l_row["BranchArea_ID"],
                                                l_row["OrderArea_ID"],
                                                l_row["PrePaidScaleGroup_ID"] != null ? 0 : l_row["PrePaidScaleGroup_ID"],
                                                l_row["ConveyKind_ID"] != null ? 0 : l_row["ConveyKind_ID"],
                                                l_row["BeginDate"],
                                                l_row["TelPhone"],
                                                l_row["FaxPhone"],
                                                l_row["CellPhone"],
                                                l_row["Contact"],
                                                l_row["Address"],
                                                l_row["Balance"],
                                                l_row["OptimisticLockField"] != null ? 0 : l_row["OptimisticLockField"],
                                                l_row["GCRecord"] != null ? 0 : l_row["GCRecord"],
                                                l_row["MobileStoreID"],
                                                l_row["CreateDate"],
                                                l_row["UpdateDate"]
                                                //DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                                                //DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")                                               
                                                );

                //l_sExeSqlCmd = string.Format(l_sInsSqlCmd,                                                 
                //                                l_row.Field<String>("ID"),
                //                                l_row.Field<String>("Name"), 
                //                                l_row.Field<Int32?>("Kind").Value, 
                //                                l_row.Field<String>("Memo"), 
                //                                l_row.Field<Int32?>("Group_ID").Value, 
                //                                l_row.Field<String>("BranchArea_ID"), 
                //                                l_row.Field<String>("OrderArea_ID"), 
                //                                l_row.Field<Int32?>("PrePaidScaleGroup_ID").Value, 
                //                                l_row.Field<Int32?>("ConveyKind_ID").Value,
                //                                l_row.Field<DateTime?>("BeginDate"),
                //                                l_row.Field<String>("TelPhone"),
                //                                l_row.Field<String>("FaxPhone"),
                //                                l_row.Field<String>("CellPhone"),
                //                                l_row.Field<String>("Contact"),
                //                                l_row.Field<String>("Address"),
                //                                l_row.Field<Single?>("Balance").Value,
                //                                l_row.Field<Int32?>("OptimisticLockField").Value,
                //                                l_row.Field<Int32?>("GCRecord").Value,
                //                                l_row.Field<String>("MobileStoreID"),
                //                                l_row.Field<String>("ManageKind"),
                //                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                //                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")                                               
                //                                );

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
                        throw new ArgumentException("StoreInfo資料新增發生錯誤!!");
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



        public List<StoreInfo> StoreInfo_GetDataListByManageKind(string p_sManageKind)
        {

            //根據上下載Type,和程式分類目錄找出File List
            //return db.FileContent.Where(m => (m.DirType == p_sDirType) && (m.ProgCatg == p_sProgCat)).ToList();

            var l_oRtnData = from DataVals in db.StoreInfo
                             where DataVals.ManageKind == p_sManageKind.Trim() && DataVals.Disabled == false
                             select DataVals;

            return (List<StoreInfo>)l_oRtnData;

        }

        //public void StoreInfo_DBCreate(string p_sSID, string p_sName, int p_iKind, string p_sMemo,
        //                       int p_iGroupID, string p_sBranchAreaID, string p_sOrderAreaID,
        //                       int p_iPrePaidScaleGroupID, int p_iConveyKindID, string p_sBeginDate,
        //                       string p_sTelPhone, string p_sFaxPhone, string p_sCellPhone,
        //                       string p_sContact, string p_sAddress, float p_fBalance,
        //                       int p_iOptimisticLockField, int p_iGCRecord, string p_sMobileStoreID,
        //                       string p_sManageKind)
        public void StoreInfo_DBCreate(string p_sSID, string p_sName, int p_iKind, string p_sMemo, 
                                       int p_iGroupID, string p_sBranchAreaID, string p_sOrderAreaID, 
                                        string p_sBeginDate,
                                       string p_sTelPhone, string p_sFaxPhone, string p_sCellPhone, 
                                       string p_sContact, string p_sAddress, 
                                        string p_sMobileStoreID,
                                       string p_sManageKind)
        {
            StoreInfo l_oNewItem = new StoreInfo();
            string l_sBeginDate = DateStringProcess.Del_MonthDayZero(p_sBeginDate, "-", "/");

            l_oNewItem.SID = p_sSID;
            l_oNewItem.Name = p_sName;
            l_oNewItem.Kind = p_iKind;
            l_oNewItem.Memo = p_sMemo;
            l_oNewItem.Group_ID = p_iGroupID;
            l_oNewItem.BranchArea_ID = p_sBranchAreaID;
            l_oNewItem.OrderArea_ID = p_sOrderAreaID;
            l_oNewItem.PrePaidScaleGroup_ID = 0;
            l_oNewItem.ConveyKind_ID = 0;
            l_oNewItem.BeginDate = l_sBeginDate; //開店日期
            l_oNewItem.TelPhone = p_sTelPhone;
            l_oNewItem.FaxPhone = p_sFaxPhone;
            l_oNewItem.CellPhone = p_sCellPhone;
            l_oNewItem.Contact = p_sContact; //聯絡人
            l_oNewItem.Address = p_sAddress;
            l_oNewItem.Balance = 0F;
            l_oNewItem.OptimisticLockField = 0;
            l_oNewItem.GCRecord = 0;
            l_oNewItem.MobileStoreID = p_sMobileStoreID;
            l_oNewItem.ManageKind = p_sManageKind;
            l_oNewItem.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            l_oNewItem.UpdateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            l_oNewItem.CreateUser = "admin";
            l_oNewItem.UpdateUser = "admin";
            l_oNewItem.Disabled = false;

            //l_oNewItem.SID = p_sSID;
            //l_oNewItem.Name = p_sName;
            //l_oNewItem.Kind = p_iKind;
            //l_oNewItem.Memo = p_sMemo;
            //l_oNewItem.Group_ID = p_iGroupID;
            //l_oNewItem.BranchArea_ID = p_sBranchAreaID;
            //l_oNewItem.OrderArea_ID = p_sOrderAreaID;
            //l_oNewItem.PrePaidScaleGroup_ID = p_iPrePaidScaleGroupID;
            //l_oNewItem.ConveyKind_ID = p_iConveyKindID;
            //l_oNewItem.BeginDate = p_sBeginDate; //開店日期
            //l_oNewItem.TelPhone = p_sTelPhone;
            //l_oNewItem.FaxPhone = p_sFaxPhone;
            //l_oNewItem.CellPhone = p_sCellPhone;
            //l_oNewItem.Contact = p_sContact; //聯絡人
            //l_oNewItem.Address = p_sAddress;
            //l_oNewItem.Balance = p_fBalance;
            //l_oNewItem.OptimisticLockField = p_iOptimisticLockField;
            //l_oNewItem.GCRecord = p_iGCRecord;
            //l_oNewItem.MobileStoreID = p_sMobileStoreID;
            //l_oNewItem.ManageKind = p_sManageKind;
            //l_oNewItem.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            //l_oNewItem.UpdateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            ///l_oNewItem.CreateDate = DateTime.Now;
            ///l_oNewItem.UpdateDate = DateTime.Now;


            db.StoreInfo.Add(l_oNewItem);

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


        public void StoreInfo_DBDeleteByID(int p_iItemID)
        {
            StoreInfo l_oDelItem = db.StoreInfo.Find(p_iItemID);
            db.StoreInfo.Remove(l_oDelItem);
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

        //public void StoreInfo_DBUpdate(int p_iId, string p_sSID, string p_sName, int p_iKind, string p_sMemo,
        //                       int p_iGroupID, string p_sBranchAreaID, string p_sOrderAreaID,
        //                       int p_iPrePaidScaleGroupID, int p_iConveyKindID, string p_sBeginDate,
        //                       string p_sTelPhone, string p_sFaxPhone, string p_sCellPhone,
        //                       string p_sContact, string p_sAddress, float p_fBalance,
        //                       int p_iOptimisticLockField, int p_iGCRecord, string p_sMobileStoreID,
        //                       string p_sManageKind)
        public void StoreInfo_DBUpdate(int p_iId, string p_sSID, string p_sName, int p_iKind, string p_sMemo,
                               int p_iGroupID, string p_sBranchAreaID, string p_sOrderAreaID,
                                string p_sBeginDate,
                               string p_sTelPhone, string p_sFaxPhone, string p_sCellPhone,
                               string p_sContact, string p_sAddress, string p_sMobileStoreID,
                               string p_sManageKind)
        {
            //StoreInfo l_oNewItem = new StoreInfo();
            StoreInfo l_oFindItem = db.StoreInfo.Find(p_iId);
            string l_sBeginDate = DateStringProcess.Del_MonthDayZero(p_sBeginDate, "-", "/");

            l_oFindItem.SID = p_sSID;
            l_oFindItem.Name = p_sName;
            l_oFindItem.Kind = p_iKind;
            l_oFindItem.Memo = p_sMemo;
            l_oFindItem.Group_ID = p_iGroupID;
            l_oFindItem.BranchArea_ID = p_sBranchAreaID;
            l_oFindItem.OrderArea_ID = p_sOrderAreaID;
            l_oFindItem.PrePaidScaleGroup_ID = 0;
            l_oFindItem.ConveyKind_ID = 0;
            l_oFindItem.BeginDate = l_sBeginDate; //開店日期
            l_oFindItem.TelPhone = p_sTelPhone;
            l_oFindItem.FaxPhone = p_sFaxPhone;
            l_oFindItem.CellPhone = p_sCellPhone;
            l_oFindItem.Contact = p_sContact; //聯絡人
            l_oFindItem.Address = p_sAddress;
            l_oFindItem.Balance = 0F;
            l_oFindItem.OptimisticLockField = 0;
            l_oFindItem.GCRecord = 0;
            l_oFindItem.MobileStoreID = p_sMobileStoreID;
            l_oFindItem.ManageKind = p_sManageKind;            
            l_oFindItem.UpdateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            l_oFindItem.UpdateUser = "admin";
            l_oFindItem.Disabled = false;

            //l_oFindItem.SID = p_sSID;
            //l_oFindItem.Name = p_sName;
            //l_oFindItem.Kind = p_iKind;
            //l_oFindItem.Memo = p_sMemo;
            //l_oFindItem.Group_ID = p_iGroupID;
            //l_oFindItem.BranchArea_ID = p_sBranchAreaID;
            //l_oFindItem.OrderArea_ID = p_sOrderAreaID;
            //l_oFindItem.PrePaidScaleGroup_ID = p_iPrePaidScaleGroupID;
            //l_oFindItem.ConveyKind_ID = p_iConveyKindID;
            //l_oFindItem.BeginDate = p_sBeginDate; //開店日期
            //l_oFindItem.TelPhone = p_sTelPhone;
            //l_oFindItem.FaxPhone = p_sFaxPhone;
            //l_oFindItem.CellPhone = p_sCellPhone;
            //l_oFindItem.Contact = p_sContact; //聯絡人
            //l_oFindItem.Address = p_sAddress;
            //l_oFindItem.Balance = p_fBalance;
            //l_oFindItem.OptimisticLockField = p_iOptimisticLockField;
            //l_oFindItem.GCRecord = p_iGCRecord;
            //l_oFindItem.MobileStoreID = p_sMobileStoreID;
            //l_oFindItem.ManageKind = p_sManageKind;
            //l_oFindItem.UpdateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

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


        public void StoreInfo_SqlDBUpdateGroup_IDBySID(string p_sSIDs, int p_iSelGroupID)
        {
            SqlCommand l_oSqlCmdObj = null;

            string l_sDBConnStr = WebConfigurationManager.ConnectionStrings["MERPSqlDB"].ConnectionString;
            SqlConnection l_oSqlConn = new SqlConnection(l_sDBConnStr);
            l_oSqlConn.Open(); //記得要先連接, 最後finally{ Close() }
            SqlTransaction l_oSqlTrans = l_oSqlConn.BeginTransaction();

            string[] l_arySIDs = null;

            if(!string.IsNullOrEmpty(p_sSIDs))
            {
                l_arySIDs = p_sSIDs.Split(',');
            }

            string l_sExeSqlCmd = "";
            string l_sUpdSqlCmd = @"UPDATE " + TB_NAME + @" SET                                    
                                   Group_ID = {0}, UpdateDate = N'{1}' 
                                   WHERE SID LIKE N'{2}' 
                                   AND Disabled = false ";

           
            ////////// DB Save /////////////////////////////////////////////////////
            StoreInfo l_oNewRow = null;
            int l_iIndex = 0, l_iRtnInsCount = 0;
            //從參數傳來的SID和Group_ID存入DB內
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();           

            foreach (string l_sSID in l_arySIDs)
            {
                ++l_iIndex;

                l_sExeSqlCmd = string.Format(l_sUpdSqlCmd,
                                                p_iSelGroupID,
                                                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                                                l_sSID                                                                                              
                                                );


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
                        throw new ArgumentException("StoreInfo資料Sql更新發生錯誤!!");
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

        #endregion

    }
}