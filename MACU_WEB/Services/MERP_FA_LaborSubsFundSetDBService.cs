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
    //Labor Insurace Fund 勞保代墊基金設定 By 各公司不同
    public class MERP_FA_LaborSubsFundSetDBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_LaborSubsFundSet";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_LaborSubsFundSet
        //
        public List<FA_LaborSubsFundSet> FA_LaborSubsFundSet_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_LaborSubsFundSet
                             where DataVals.IsValid == 1
                             select DataVals;

            return l_oRtnData.ToList();
            
        }

        //
        public FA_LaborSubsFundSet FA_LaborSubsFundSet_GetDataById(int p_iId)
        {


            FA_LaborSubsFundSet l_oFindItem = db.FA_LaborSubsFundSet
                    .Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }


        public FA_LaborSubsFundSet FA_LaborSubsFundSet_GetDataByPlusInsCom(string p_sPlusInsCompany)
        {

            var l_oRtnData = from DataVals in db.FA_LaborSubsFundSet
                             where DataVals.IsValid == 1 && 
                                   DataVals.PlusInsCompany == p_sPlusInsCompany.Trim()
                             select DataVals;

            FA_LaborSubsFundSet l_oFindItem = l_oRtnData.First();

            return l_oFindItem;
        }

        //
        public void FA_LaborSubsFundSet_DBCreate(
                                 string p_dPlusInsCompany, decimal p_dLaborSubsFund)
        {
            FA_LaborSubsFundSet l_oNewItem = new FA_LaborSubsFundSet();

            l_oNewItem.PlusInsCompany = p_dPlusInsCompany;
            l_oNewItem.LaborSubsFund = p_dLaborSubsFund;  //勞保代墊基金               
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_LaborSubsFundSet.Add(l_oNewItem);

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
        public void FA_LaborSubsFundSet_DBDeleteByID(int p_iItemID)
        {
            FA_LaborSubsFundSet l_oDelItem = db.FA_LaborSubsFundSet.Where(m => m.IsValid == 1)
                                                         .ToList().Find(m => m.Id == p_iItemID);
            db.FA_LaborSubsFundSet.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //
        public void FA_LaborSubsFundSet_DBUpdate(int p_iItemID, FA_LaborSubsFundSet p_oNewUpdItem)
        {
            FA_LaborSubsFundSet l_oUpdItem = db.FA_LaborSubsFundSet.Find(p_iItemID);


            l_oUpdItem.PlusInsCompany = p_oNewUpdItem.PlusInsCompany;
            l_oUpdItem.LaborSubsFund = p_oNewUpdItem.LaborSubsFund;            
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