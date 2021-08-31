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
    //Labor Insurace 勞保設定
    public class MERP_FA_LaborInsSetDBService
    {
        private const int BATCH_COUNT = 100;
        private const string TB_NAME = "FA_LaborInsSet";

        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region FA_LaberInsSet

        //1.
        public List<FA_LaborInsSet> FA_LaborInsSet_GetDataList()
        {
            var l_oRtnData = from DataVals in db.FA_LaborInsSet
                             where DataVals.IsValid == 1
                             select DataVals;

            if(l_oRtnData == null)
            {
                return null;
            } else
            {
                return l_oRtnData.ToList();
            }

            
            //return db.FA_LaborHealthIns.ToList();
        }

        //2.
        public FA_LaborInsSet FA_LaborInsSet_GetDataById(int p_iId)
        {


            FA_LaborInsSet l_oFindItem = db.FA_LaborInsSet
                    .Where(m => m.IsValid == 1).ToList().Find(m => m.Id == p_iId);

            return l_oFindItem;
        }

        //20210126 CCL+
        public FA_LaborInsSet FA_LaborInsSet_GetDataByNo(string  p_sNo)
        {


            FA_LaborInsSet l_oFindItem = db.FA_LaborInsSet
                    .Where(m => m.IsValid == 1).ToList().Find(m => m.LabInsSetNo == p_sNo);

            return l_oFindItem;
        }

        public FA_LaborInsSet FA_LaborInsSet_GetDataByNewestBeginDate()
        {


            FA_LaborInsSet l_oFindItem = db.FA_LaborInsSet
                    .Where(m => m.IsValid == 1).ToList().OrderByDescending(m => m.OnBeginDate).First();

            return l_oFindItem;
        }

        //3.
        public void FA_LaborInsSet_DBCreate(
                                 string p_sLabInsSetNo,
                                 decimal p_dOrdAccidentInsRate, decimal p_dEmployInsRate,
                                 decimal p_dPersonalInsRate, decimal p_dLaborBurdenRatio,
                                 decimal p_dComBurdenRatio, decimal p_dGovBurdenRatio, 
                                 decimal p_dCommuteDisaInsRate, decimal p_dIndustryDisaInsRate,
                                 decimal p_dOccuDisaInsRate, decimal p_dOccuDisComBurdenRatio,
                                 decimal p_dLaborSubsFund, decimal p_dLaborSubsFundRate,
                                 decimal p_dLaborRetireRate, 
                                 string p_sOnBeginDate)
        {
            FA_LaborInsSet l_oNewItem = new FA_LaborInsSet();

            l_oNewItem.LabInsSetNo = p_sLabInsSetNo;
            l_oNewItem.OrdAccidentInsRate = p_dOrdAccidentInsRate;
            l_oNewItem.EmployInsRate = p_dEmployInsRate;
            l_oNewItem.PersonalInsRate = p_dPersonalInsRate;
            l_oNewItem.LaborBurdenRatio = p_dLaborBurdenRatio;
            l_oNewItem.ComBurdenRatio = p_dComBurdenRatio;
            l_oNewItem.GovBurdenRatio = p_dGovBurdenRatio;
            l_oNewItem.CommuteDisaInsRate = p_dCommuteDisaInsRate;
            l_oNewItem.IndustryDisaInsRate = p_dIndustryDisaInsRate;
            l_oNewItem.OccuDisaInsRate = p_dOccuDisaInsRate;
            l_oNewItem.OccuDisComBurdenRatio = p_dOccuDisComBurdenRatio;
            l_oNewItem.OnBeginDate = p_sOnBeginDate;
            l_oNewItem.LaborSubsFund = p_dLaborSubsFund;
            l_oNewItem.LaborSubsFundRate = p_dLaborSubsFundRate;
            l_oNewItem.LaborRetireRate = p_dLaborRetireRate;
            l_oNewItem.IsValid = 1;
            l_oNewItem.CreateTime = DateTime.Now;
            l_oNewItem.UpdateTime = DateTime.Now;

            db.FA_LaborInsSet.Add(l_oNewItem);

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
        public void FA_LaborInsSet_DBDeleteByID(int p_iItemID)
        {
            FA_LaborInsSet l_oDelItem = db.FA_LaborInsSet.Where(m => m.IsValid == 1)
                                                         .ToList().Find(m => m.Id == p_iItemID);
            db.FA_LaborInsSet.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //20210126 CCL+
        public void FA_LaborInsSet_DBDeleteByNo(string p_sItemNo)
        {
            FA_LaborInsSet l_oDelItem = db.FA_LaborInsSet.Where(m => m.IsValid == 1)
                                                         .ToList().Find(m => m.LabInsSetNo == p_sItemNo);
            db.FA_LaborInsSet.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //5.
        
        public void FA_LaborInsSet_DBUpdate(int p_iItemID, FA_LaborInsSet p_oNewUpdItem)
        {
            FA_LaborInsSet l_oUpdItem = db.FA_LaborInsSet.Find(p_iItemID);

            l_oUpdItem.LabInsSetNo = p_oNewUpdItem.LabInsSetNo;
            l_oUpdItem.OrdAccidentInsRate = p_oNewUpdItem.OrdAccidentInsRate;
            l_oUpdItem.EmployInsRate = p_oNewUpdItem.EmployInsRate;
            l_oUpdItem.PersonalInsRate = p_oNewUpdItem.PersonalInsRate;
            l_oUpdItem.LaborBurdenRatio = p_oNewUpdItem.LaborBurdenRatio;
            l_oUpdItem.ComBurdenRatio = p_oNewUpdItem.ComBurdenRatio;
            l_oUpdItem.GovBurdenRatio = p_oNewUpdItem.GovBurdenRatio;
            l_oUpdItem.CommuteDisaInsRate = p_oNewUpdItem.CommuteDisaInsRate;
            l_oUpdItem.IndustryDisaInsRate = p_oNewUpdItem.IndustryDisaInsRate;
            l_oUpdItem.OccuDisaInsRate = p_oNewUpdItem.OccuDisaInsRate;
            l_oUpdItem.OccuDisComBurdenRatio = p_oNewUpdItem.OccuDisComBurdenRatio;
            l_oUpdItem.OnBeginDate = p_oNewUpdItem.OnBeginDate;
            l_oUpdItem.LaborSubsFund = p_oNewUpdItem.LaborSubsFund;
            l_oUpdItem.LaborSubsFundRate = p_oNewUpdItem.LaborSubsFundRate;
            l_oUpdItem.LaborRetireRate = p_oNewUpdItem.LaborRetireRate;
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
        

        public void FA_LaborInsSet_DBUpdateByNO(string p_sItemNO, FA_LaborInsSet p_oNewUpdItem)
        {
            //FA_LaborInsSet l_oUpdItem = db.FA_LaborInsSet.Find(p_iItemID);
            FA_LaborInsSet l_oUpdItem = db.FA_LaborInsSet.Find(p_sItemNO);

            l_oUpdItem.LabInsSetNo = p_oNewUpdItem.LabInsSetNo;
            l_oUpdItem.OrdAccidentInsRate = p_oNewUpdItem.OrdAccidentInsRate;
            l_oUpdItem.EmployInsRate = p_oNewUpdItem.EmployInsRate;
            l_oUpdItem.PersonalInsRate = p_oNewUpdItem.PersonalInsRate;
            l_oUpdItem.LaborBurdenRatio = p_oNewUpdItem.LaborBurdenRatio;
            l_oUpdItem.ComBurdenRatio = p_oNewUpdItem.ComBurdenRatio;
            l_oUpdItem.GovBurdenRatio = p_oNewUpdItem.GovBurdenRatio;
            l_oUpdItem.CommuteDisaInsRate = p_oNewUpdItem.CommuteDisaInsRate;
            l_oUpdItem.IndustryDisaInsRate = p_oNewUpdItem.IndustryDisaInsRate;
            l_oUpdItem.OccuDisaInsRate = p_oNewUpdItem.OccuDisaInsRate;
            l_oUpdItem.OccuDisComBurdenRatio = p_oNewUpdItem.OccuDisComBurdenRatio;
            l_oUpdItem.OnBeginDate = p_oNewUpdItem.OnBeginDate;
            l_oUpdItem.LaborSubsFund = p_oNewUpdItem.LaborSubsFund;
            l_oUpdItem.LaborSubsFundRate = p_oNewUpdItem.LaborSubsFundRate;
            l_oUpdItem.LaborRetireRate = p_oNewUpdItem.LaborRetireRate;
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