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
    //Health Insurace 健保設定
    public class MERP_FA_HealthInsSetDBService
    {
        
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_HealthInsSet";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_HealthInsSet

        //1.
        public List<FA_HealthInsSet> FA_HealthInsSet_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_HealthInsSet
                             where DataVals.IsValid == 1
                             select DataVals;

            return l_oRtnData.ToList();
            
        }

        //2.
        public FA_HealthInsSet FA_HealthInsSet_GetDataById(int p_iId)
        {


            FA_HealthInsSet l_oFindItem = db.FA_HealthInsSet
                    .Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }

        public FA_HealthInsSet FA_HealthInsSet_GetDataByNewestBeginDate()
        {


            FA_HealthInsSet l_oFindItem = db.FA_HealthInsSet
                    .Where(m => m.IsValid == 1).ToList().OrderByDescending(m => m.Heal_RateBeginDate).First();

            return l_oFindItem;
        }

        //3.
        public void FA_HealthInsSet_DBCreate(
                                 decimal p_dHeal_Rate,
                                 decimal p_dHeal_LaborInsBurdenRatio, decimal p_dHeal_ComInsBurdenRatio,
                                 decimal p_dHeal_GovInsBurdenRatio, decimal p_dHeal_AverhouseholdsNum,
                                 string p_sHeal_RateBeginDate)
        {
            FA_HealthInsSet l_oNewItem = new FA_HealthInsSet();

            l_oNewItem.Heal_Rate = p_dHeal_Rate;
            l_oNewItem.Heal_RateBeginDate = p_sHeal_RateBeginDate;
            l_oNewItem.Heal_LaborInsBurdenRatio = p_dHeal_LaborInsBurdenRatio;
            l_oNewItem.Heal_ComInsBurdenRatio = p_dHeal_ComInsBurdenRatio;
            l_oNewItem.Heal_GovInsBurdenRatio = p_dHeal_GovInsBurdenRatio;
            l_oNewItem.Heal_AverhouseholdsNum = p_dHeal_AverhouseholdsNum;           
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_HealthInsSet.Add(l_oNewItem);

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
        public void FA_HealthInsSet_DBDeleteByID(int p_iItemID)
        {
            FA_HealthInsSet l_oDelItem = db.FA_HealthInsSet.Where(m => m.IsValid == 1)
                                                         .ToList().Find(m => m.Id == p_iItemID);
            db.FA_HealthInsSet.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //5.
        public void FA_HealthInsSet_DBUpdate(int p_iItemID, FA_HealthInsSet p_oNewUpdItem)
        {
            FA_HealthInsSet l_oUpdItem = db.FA_HealthInsSet.Find(p_iItemID);


            l_oUpdItem.Heal_Rate = p_oNewUpdItem.Heal_Rate;
            l_oUpdItem.Heal_RateBeginDate = p_oNewUpdItem.Heal_RateBeginDate;
            l_oUpdItem.Heal_LaborInsBurdenRatio = p_oNewUpdItem.Heal_LaborInsBurdenRatio;
            l_oUpdItem.Heal_ComInsBurdenRatio = p_oNewUpdItem.Heal_ComInsBurdenRatio;
            l_oUpdItem.Heal_GovInsBurdenRatio = p_oNewUpdItem.Heal_GovInsBurdenRatio;
            l_oUpdItem.Heal_AverhouseholdsNum = p_oNewUpdItem.Heal_AverhouseholdsNum;
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