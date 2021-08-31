using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using MACU_WEB.BIServices;
using MACU_WEB.Areas.MERP_TCC000.ViewModels;
using MACU_WEB.Models._Base;
using System.IO;

namespace MACU_WEB.Areas.MERP_TCF000.Controllers
{
    //部門健保負擔比例設定Settings
    public class MERP_TCF003Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TCF003"; //客製程式
        string strMENU_ID = "MERP_TCF000";

        public MERP_FA_HealthInsSetDBService m_HInsSettingDBService = new MERP_FA_HealthInsSetDBService();
        #endregion


        #region Action_View
        // GET: MERP_TCF000/MERP_TCF003
        public ActionResult Index()
        {
            List<FA_HealthInsSet> l_oRtnData = m_HInsSettingDBService.FA_HealthInsSet_GetDataList();

            return View(l_oRtnData);

           
        }

        // GET: MERP_TCF000/MERP_TCF003/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MERP_TCF000/MERP_TCF003/Create
        public ActionResult Create()
        {
            return View();
        }

        // GET: MERP_TCF000/MERP_TCF003/Delete/5
        public ActionResult Delete(int id)
        {
            m_HInsSettingDBService.FA_HealthInsSet_DBDeleteByID(id);

            return RedirectToAction("Index");
         
        }

        // GET: MERP_TCF000/MERP_TCF003/Edit/5
        public ActionResult Edit(int id)
        {

            FA_HealthInsSet l_oRtnData = m_HInsSettingDBService.FA_HealthInsSet_GetDataById(id);
            //把日期0回復和改成"-"分隔才能正常顯示
            string l_sRateBeginDate = l_oRtnData.Heal_RateBeginDate;
            l_sRateBeginDate = DateStringProcess.ReStore_MonthDayZero(l_sRateBeginDate, "/", "-");
            l_oRtnData.Heal_RateBeginDate = l_sRateBeginDate;

            ViewData["FA_HInsSet"] = l_oRtnData;
            //return View(l_oRtnData);
            return View();
            
        }
        #endregion


        #region Action_DB
        // POST: MERP_TCF000/MERP_TCF003/Create
        [HttpPost]
        public ActionResult Create(FormCollection p_oForm)
        {
            try
            {
                // TODO: Add insert logic here

                string l_sRateBeginDate = "";
                // TODO: Add insert logic here
                if (!string.IsNullOrEmpty(p_oForm["Heal_RateBeginDate"]))
                {
                    l_sRateBeginDate = DateStringProcess.Del_MonthDayZero(p_oForm["Heal_RateBeginDate"], "-", "/");
                }


                m_HInsSettingDBService.FA_HealthInsSet_DBCreate(
                                                    Convert.ToDecimal(p_oForm["Heal_Rate"]),
                                                    Convert.ToDecimal(p_oForm["Heal_LaborInsBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["Heal_ComInsBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["Heal_GovInsBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["Heal_AverhouseholdsNum"]),                                                    
                                                    l_sRateBeginDate
                                                    );

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        // POST: MERP_TCF000/MERP_TCF003/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection p_oForm)
        {
            try
            {
                // TODO: Add update logic here
                string l_sRateBeginDate = "";
                // TODO: Add insert logic here
                if (!string.IsNullOrEmpty(p_oForm["Heal_RateBeginDate"]))
                {
                    l_sRateBeginDate = DateStringProcess.Del_MonthDayZero(p_oForm["Heal_RateBeginDate"], "-", "/");
                }

                FA_HealthInsSet l_oUpdItem = new FA_HealthInsSet();
                l_oUpdItem.Heal_Rate = Convert.ToDecimal(p_oForm["Heal_Rate"]);
                l_oUpdItem.Heal_LaborInsBurdenRatio = Convert.ToDecimal(p_oForm["Heal_LaborInsBurdenRatio"]);
                l_oUpdItem.Heal_ComInsBurdenRatio = Convert.ToDecimal(p_oForm["Heal_ComInsBurdenRatio"]);
                l_oUpdItem.Heal_GovInsBurdenRatio = Convert.ToDecimal(p_oForm["Heal_GovInsBurdenRatio"]);
                l_oUpdItem.Heal_AverhouseholdsNum = Convert.ToDecimal(p_oForm["Heal_AverhouseholdsNum"]);                           
                l_oUpdItem.Heal_RateBeginDate = l_sRateBeginDate;


                m_HInsSettingDBService.FA_HealthInsSet_DBUpdate(id, l_oUpdItem);

                return RedirectToAction("Index");
                
            }
            catch
            {
                return View();
            }
        }


        // POST: MERP_TCF000/MERP_TCF003/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here
               

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        #endregion
    }
}
