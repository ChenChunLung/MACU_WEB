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

namespace MACU_WEB.Areas.MERP_TCC000.Controllers
{
    public class MERP_TCC003Controller : Controller
    {
        //會計科目資訊 Map Table
        #region  Param Initial
        string strPROG_ID = "MERP_TCC003"; //客製程式
        string strMENU_ID = "MERP_TCC000";

        public MERP_AccountInfoDBService m_AccInfoDBService = new MERP_AccountInfoDBService();
        #endregion

        // GET: MERP_TCC000/MERP_TCC003
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出上傳的日記帳檔案
            List<AccountInfo> l_oDataList = m_AccInfoDBService.AccountInfo_GetDataList();

            return View(l_oDataList);
            
        }

        // GET: MERP_TCC000/MERP_TCC003/Insert
        public ActionResult Insert()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

          
            return View();

        }

        // GET: MERP_TCC000/MERP_TCC003/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MERP_TCC000/MERP_TCC003/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC003/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MERP_TCC000/MERP_TCC003/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC003/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MERP_TCC000/MERP_TCC003/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC003/Delete/5
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

        // POST: MERP_TCC000/MERP_TCC003/Insert
        [HttpPost]
        public ActionResult Insert(FormCollection p_oForm)
        {
            string l_sAccNo = p_oForm["AccountNo"];
            string l_sAccName = p_oForm["AccountName"];
            string l_sDetailAccNo = p_oForm["DetailAccNo"];
            string l_sDetailAccName = p_oForm["DetailAccName"];
            string l_sCountFlag = p_oForm["CountFlag"];
            int l_iPrintOrder = Convert.ToInt32(p_oForm["PrintOrder"]);
            int l_iGroupID = Convert.ToInt32(p_oForm["GroupID"]);

            m_AccInfoDBService.AccountInfo_DBCreate(l_sAccNo, l_sAccName,
                                                    l_sDetailAccNo, l_sDetailAccName,
                                                    l_sCountFlag, l_iPrintOrder, l_iGroupID);


            return View();
           
        }
    }
}
