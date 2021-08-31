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


namespace MACU_WEB.Services
{
    public class MERP_StoreGroupSetDBService
    {
        private const int BATCH_COUNT = 100;        
        private const string TB_NAME = "StoreGroupSet";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region StoreGroupSet
        public List<StoreGroupSet> StoreGroupSet_GetDataList()
        {
            var l_oRtnData = from DataVals in db.StoreGroupSet
                             where DataVals.IsValid == 1
                             select DataVals;

           
            return l_oRtnData.ToList();
            
        }


        public StoreGroupSet StoreGroupSet_GetDataByGroupNo(int p_sGroupNo )
        {
            //以StoreGroupNo來
            if(db.StoreGroupSet.Count() == 0)
                return null;

            var l_oRtnData = from DataVals in db.StoreGroupSet
                             where DataVals.IsValid == 1 &&
                                    DataVals.StoreGroupNo == p_sGroupNo
                             select DataVals;


            //傳回
            if (l_oRtnData != null && l_oRtnData.Count() > 0)
                return l_oRtnData.First();
            else
                return null;
        }

        //20210225 CCL+ 增加區域Type
        public StoreGroupSet StoreGroupSet_GetDataByGroupNoType(int p_sGroupNo, string p_sGroupType)
        {
            //以StoreGroupNo來
            if (db.StoreGroupSet.Count() == 0)
                return null;

            var l_oRtnData = from DataVals in db.StoreGroupSet
                             where DataVals.IsValid == 1 &&
                                    DataVals.StoreGroupNo == p_sGroupNo && 
                                    DataVals.StoreGroupType == p_sGroupType
                             select DataVals;


            //傳回
            if (l_oRtnData != null && l_oRtnData.Count() > 0)
                return l_oRtnData.First();
            else
                return null;
        }

        /*
        public StoreGroupSet StoreGroupSet_GetDataById(int p_iId)
        {

            StoreGroupSet l_oFindItem = db.StoreGroupSet.Where(m => m.IsValid == 1)
                                        .ToList().Find(m => m.Id == p_iId);
            return l_oFindItem;
        }
        */


        public void StoreGroupSet_DBCreate(int p_iStoreGroupNo,
                         string p_sStoreGroupType,
                         string p_sStoreGroupName,
                         string p_sStoreGroupDesc,
                          string p_sStoreGroupSIDList)
        {
            StoreGroupSet l_oNewItem = new StoreGroupSet();

            l_oNewItem.StoreGroupType = p_sStoreGroupType.ToString();
            l_oNewItem.StoreGroupNo = p_iStoreGroupNo;
            l_oNewItem.StoreGroupName = p_sStoreGroupName;
            l_oNewItem.StoreGroupDesc = p_sStoreGroupDesc;
            l_oNewItem.StoreGroupSIDList = p_sStoreGroupSIDList.Trim();
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.StoreGroupSet.Add(l_oNewItem);

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
        

        

        /*
        public void StoreGroupSet_DBDeleteByID(int p_iItemID)
        {
            StoreGroupSet l_oDelItem = db.StoreGroupSet.Find(p_iItemID);
            db.StoreGroupSet.Remove(l_oDelItem);

            try
            {
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }

        }
        */

        //GroupNo
        public void StoreGroupSet_DBDeleteByGroupNo(int p_iItemNo)
        {
            //要注意TB內Primary Key是GroupNo
            StoreGroupSet l_oDelItem = db.StoreGroupSet.Find(p_iItemNo);
            db.StoreGroupSet.Remove(l_oDelItem);

            try
            {
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }

        }

        //20210225 CCL+ 增加區域Type
        public void StoreGroupSet_DBDeleteByGroupNoType(int p_iItemNo, string p_sItemType)
        {
            //要注意TB內Primary Key是GroupNo
            var l_oFindData = from DataVals in db.StoreGroupSet
                              where DataVals.IsValid == 1 &&
                              DataVals.StoreGroupNo == p_iItemNo &&
                              DataVals.StoreGroupType == p_sItemType
                              select DataVals;

            StoreGroupSet l_oDelItem = l_oFindData.First();
            //StoreGroupSet l_oDelItem = db.StoreGroupSet.Find(p_iItemNo);
            db.StoreGroupSet.Remove(l_oDelItem);

            try
            {
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }

        }


        /* 20210225 CCL-
        public void StoreGroupSet_DBUpdate(int p_iItemNo, StoreGroupSet p_oUpdItem)
        {          

            StoreGroupSet l_oFindItem = db.StoreGroupSet.Find(p_iItemNo); ;

            l_oFindItem.StoreGroupNo = p_oUpdItem.StoreGroupNo;
            l_oFindItem.StoreGroupType = p_oUpdItem.StoreGroupType;            
            l_oFindItem.StoreGroupName = p_oUpdItem.StoreGroupName;
            //20210225 CCL- 不更新 l_oFindItem.StoreGroupDesc = p_oUpdItem.StoreGroupDesc;
            l_oFindItem.StoreGroupSIDList = p_oUpdItem.StoreGroupSIDList;
            l_oFindItem.IsValid = 1;            
            l_oFindItem.UpdateTime = DateTime.Now;            

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

        //20210225 CCL*
        public void StoreGroupSet_DBUpdate(int p_iItemNo, StoreGroupSet p_oUpdItem)
        {

            //StoreGroupSet l_oFindItem = db.StoreGroupSet.Find(p_iItemNo); 
            StoreGroupSet l_oFindItem = 
                StoreGroupSet_GetDataByGroupNoType(p_oUpdItem.StoreGroupNo, p_oUpdItem.StoreGroupType);

            l_oFindItem.StoreGroupNo = p_oUpdItem.StoreGroupNo;
            l_oFindItem.StoreGroupType = p_oUpdItem.StoreGroupType;
            //20210225 CCL- 不更新 l_oFindItem.StoreGroupName = p_oUpdItem.StoreGroupName;
            //20210225 CCL- 不更新 l_oFindItem.StoreGroupDesc = p_oUpdItem.StoreGroupDesc;
            l_oFindItem.StoreGroupSIDList = p_oUpdItem.StoreGroupSIDList;
            l_oFindItem.IsValid = 1;
            l_oFindItem.UpdateTime = DateTime.Now;

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



        #endregion

    }
}