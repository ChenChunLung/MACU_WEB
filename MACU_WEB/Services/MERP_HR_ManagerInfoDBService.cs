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
    public class MERP_HR_ManagerInfoDBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "HR_ManagerInfo";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_JournalV1
        public List<HR_ManagerInfo> HR_ManagerInfo_GetDataList()
        {
            return db.HR_ManagerInfo.ToList().Where(m => m.IsValid == 1).ToList();
            //return db.HR_ManagerInfo.ToList();
        }

        public HR_ManagerInfo HR_ManagerInfo_GetDataById(int p_iId)
        {

            //HR_ManagerInfo l_oFindItem = db.HR_ManagerInfo.Where(m => (m.IsValid == 1) &&
            //                                                          (m.Id == p_iId)).First();
            HR_ManagerInfo l_oFindItem = db.HR_ManagerInfo.Find(p_iId);
            return l_oFindItem;
        }


        public void HR_ManagerInfo_DBCreate(string p_sManagerID, string p_sManagerName,
                                    string p_sManagerNickNa,
                                    string p_sManagerTelPhone,
                                    string p_sManagerMobiPhone,
                                    string p_sManageBranchID,
                                    string p_sManageShopList)
        {
            HR_ManagerInfo l_oNewItem = new HR_ManagerInfo();

            l_oNewItem.ManagerID = p_sManagerID;
            l_oNewItem.ManagerName = p_sManagerName;
            l_oNewItem.ManagerNickNa = p_sManagerNickNa;            
            l_oNewItem.ManagerTelPhone = p_sManagerTelPhone;
            l_oNewItem.ManagerMobiPhone = p_sManagerMobiPhone;
            l_oNewItem.ManageBranchID = p_sManageBranchID;
            l_oNewItem.ManageShopList = p_sManageShopList;
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;



            db.HR_ManagerInfo.Add(l_oNewItem);

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


        public void HR_ManagerInfo_DBDeleteByID(int p_iItemID)
        {
            HR_ManagerInfo l_oDelItem = db.HR_ManagerInfo.Find(p_iItemID);
            db.HR_ManagerInfo.Remove(l_oDelItem);
            db.SaveChanges();
        }

        public void HR_ManagerInfo_DBUpdate(int id, string p_sManagerID, string p_sManagerName,
                            string p_sManagerNickNa,
                            string p_sManagerTelPhone,
                            string p_sManagerMobiPhone,
                            string p_sManageBranchID,
                            string p_sManageShopList)
        {
            HR_ManagerInfo l_oFindItem = db.HR_ManagerInfo.Find(id);

            l_oFindItem.ManagerID = p_sManagerID;
            l_oFindItem.ManagerName = p_sManagerName;
            l_oFindItem.ManagerNickNa = p_sManagerNickNa;
            l_oFindItem.ManagerTelPhone = p_sManagerTelPhone;
            l_oFindItem.ManagerMobiPhone = p_sManagerMobiPhone;
            l_oFindItem.ManageBranchID = p_sManageBranchID;
            l_oFindItem.ManageShopList = p_sManageShopList;
            l_oFindItem.IsValid = 1;            
            l_oFindItem.UpdateTime = DateTime.Now;



            //db.HR_ManagerInfo.Add(l_oNewItem);

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