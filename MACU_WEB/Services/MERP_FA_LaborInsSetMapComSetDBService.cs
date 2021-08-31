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
    public class MERP_FA_LaborInsSetMapComSetDBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_LaborInsSetMapComSet";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_LaborInsSetMapComSet
        //
        public List<FA_LaborInsSetMapComSet> FA_LaborInsSetMapComSet_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_LaborInsSetMapComSet
                             where DataVals.IsValid == 1
                             select DataVals;

            return l_oRtnData.ToList();

        }

        //
        public FA_LaborInsSetMapComSet FA_LaborInsSetMapComSet_GetDataById(int p_iId)
        {


            FA_LaborInsSetMapComSet l_oFindItem = db.FA_LaborInsSetMapComSet
                    .Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }


        public FA_LaborInsSetMapComSet FA_LaborInsSetMapComSet_GetDataByPlusInsCom(string p_sPlusInsCompany)
        {

            var l_oRtnData = from DataVals in db.FA_LaborInsSetMapComSet
                             where DataVals.IsValid == 1 &&
                                   DataVals.PlusInsCompany == p_sPlusInsCompany.Trim()
                             select DataVals;

            FA_LaborInsSetMapComSet l_oFindItem = l_oRtnData.First();

            return l_oFindItem;
        }

        //直接取出相對應的勞保設定 20210129 CCL+
        public FA_LaborInsSet FA_LaborInsSetMapComSet_GetDataLInsSetByPlusInsCom(string p_sPlusInsCompany)
        {

            var l_oRtnData = from DataVals in db.FA_LaborInsSetMapComSet
                             from Vals in db.FA_LaborInsSet
                             where DataVals.IsValid == 1 &&
                                   DataVals.PlusInsCompany == p_sPlusInsCompany.Trim() &&
                                   DataVals.LabInsSetNo == Vals.LabInsSetNo
                             select Vals;


            FA_LaborInsSet l_oFindItem = l_oRtnData.First();

            return l_oFindItem;
        }

        //
        public void FA_LaborInsSetMapComSet_DBCreate(
                                 string p_dPlusInsCompany, string p_sLaborInsSetNo)
        {
            FA_LaborInsSetMapComSet l_oNewItem = new FA_LaborInsSetMapComSet();

            l_oNewItem.PlusInsCompany = p_dPlusInsCompany;
            l_oNewItem.LabInsSetNo = p_sLaborInsSetNo;  //勞保設定               
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_LaborInsSetMapComSet.Add(l_oNewItem);

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

        //
        public void FA_LaborInsSetMapComSet_DBDeleteByID(int p_iItemID)
        {
            FA_LaborInsSetMapComSet l_oDelItem = db.FA_LaborInsSetMapComSet.Where(m => m.IsValid == 1)
                                                         .ToList().Find(m => m.Id == p_iItemID);
            db.FA_LaborInsSetMapComSet.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //
        public void FA_LaborInsSetMapComSet_DBUpdate(int p_iItemID, FA_LaborInsSetMapComSet p_oNewUpdItem)
        {
            FA_LaborInsSetMapComSet l_oUpdItem = db.FA_LaborInsSetMapComSet.Find(p_iItemID);


            l_oUpdItem.PlusInsCompany = p_oNewUpdItem.PlusInsCompany;
            l_oUpdItem.LabInsSetNo = p_oNewUpdItem.LabInsSetNo;
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

        public void FA_LaborInsSetMapComSet_DBUpdate(int p_iItemID, string p_sLaborInsSetNO)
        {
            FA_LaborInsSetMapComSet l_oUpdItem = db.FA_LaborInsSetMapComSet.Find(p_iItemID);

           
            l_oUpdItem.LabInsSetNo = p_sLaborInsSetNO;
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


        #endregion


    }
}