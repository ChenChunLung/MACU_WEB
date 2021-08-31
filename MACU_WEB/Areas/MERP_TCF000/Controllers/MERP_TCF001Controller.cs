using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using MACU_WEB.BIServices;
using MACU_WEB.Models._Base;
using System.IO;
using MACU_WEB.Areas.MERP_TCF000.ViewModels;

namespace MACU_WEB.Areas.MERP_TCF000.Controllers
{
    //部門勞健保薪資計算匯入
    public class MERP_TCF001Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TCF001"; //部門勞健保薪資計算
        string strMENU_ID = "MERP_TCF000";

        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        public MERP_FA_LaborHealthInsDBService m_LHInsDBService = new MERP_FA_LaborHealthInsDBService();
        #endregion

        #region Action_View
        // GET: MERP_TCF000/MERP_TCF001
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出上傳的日記帳檔案
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);


            return View(l_oDataList);
        }

        // GET: MERP_TCF000/MERP_TCF001/Details/5
        public ActionResult Details(int id)
        {
            string l_sLHInsMonth = "";

            FileContent l_oSearchFile = m_FileDBService.FileContent_GetDataById(id);
            //載入上傳的Excel,並且匯入DataBase
            l_sLHInsMonth = MERP_LaborHealthExcelBIService.ImportExcelTo_FA_LaborHealth(l_oSearchFile);

            //顯示本年月的All
            List<FA_LaborHealthIns> l_oLHIns = MERP_LaborHealthExcelBIService.GetImportExcelInDB_YearMonthData(
                                                                    l_oSearchFile.DataYear,
                                                                    l_oSearchFile.DataMonth);


            //顯示本月的各分頁Data(300筆)
            List<FA_LaborHealthIns> l_oFaLHIns = MERP_LaborHealthExcelBIService.GetImportExcelInDB_YearMonthDataPage(l_oSearchFile.DataYear,
                                                                                                                   l_oSearchFile.DataMonth, 1);

            MERP_TCF001_DetailsViewModel l_oDetailsVM = new MERP_TCF001_DetailsViewModel();
            l_oDetailsVM.m_sYear = l_oSearchFile.DataYear;
            l_oDetailsVM.m_sMonth = l_oSearchFile.DataMonth;
            l_oDetailsVM.m_oFALaborHealthInsList = l_oFaLHIns;

            return View(l_oDetailsVM);
            //return View();
        }


        // GET: MERP_TCF000/MERP_TCF001/Create
        public ActionResult Create()
        {
            return View();
        }


        // GET: MERP_TCF000/MERP_TCF001/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }


        // GET: MERP_TCF000/MERP_TCF001/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // GET: MERP_TCC000/MERP_TCC001/Journals/5
        public ActionResult Journals(int id, string year, string month)
        {

            //20210106 CCl Mod

            //string l_sMonth = DateTime.Now.Month.ToString();
            //string l_sYear = DateTime.Now.Year.ToString();
            string l_sMonth = month;
            string l_sYear = year;
            //顯示本月的各分頁Data(300筆)
            List<FA_LaborHealthIns> l_oFaLHIns = MERP_LaborHealthExcelBIService.GetImportExcelInDB_YearMonthDataPage(l_sYear,
                                                                                                               l_sMonth, 0);
            
            MERP_TCF001_DetailsViewModel l_oJournalsVM = new MERP_TCF001_DetailsViewModel();


            l_oJournalsVM.m_sYear = l_sYear;  //Year
            l_oJournalsVM.m_sMonth = l_sMonth; //Month
            l_oJournalsVM.m_oFALaborHealthInsList = l_oFaLHIns;
           
            //l_oJournalsVM.m_oSelShopList = l_RtnShopsData.ToList();

            return View(l_oJournalsVM);
        }

        #endregion

        #region Action_DB
        // POST: MERP_TCF000/MERP_TCF001/Index
        [HttpPost]
        #region 查詢畫面送出(Index) [Submit]
        public ActionResult Index(string year_month, HttpPostedFileBase upload)
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            if(!string.IsNullOrEmpty(year_month))
            {
                DateStringProcess.Del_MonthZero(year_month, "-", "");
                string l_sYear = DateStringProcess.m_Year.ToString();
                string l_sMonth = DateStringProcess.m_Month.ToString();
                    

                //執行上傳檔案           
                MERP_UploadBIService.UploadFile(upload, Server, strPROG_ID, l_sYear, l_sMonth);

                //只查出Upload及本程式分類File
                List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);


                return View(l_oDataList);
            } else
            {
                return View();
            }

        }
        #endregion

        // POST: MERP_TCF000/MERP_TCF001/Create
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

        // POST: MERP_TCF000/MERP_TCF001/Edit/5
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

        // POST: MERP_TCF000/MERP_TCF001/Delete/5
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
