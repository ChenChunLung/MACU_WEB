using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using MACU_WEB.BIServices;
using MACU_WEB.Areas.MERP_TCF000.ViewModels;
using MACU_WEB.Models._Base;
using System.IO;

namespace MACU_WEB.Areas.MERP_TCF000.Controllers
{
    ////部門勞健保薪資計算匯入 新版V1
    public class MERP_TCF004Controller : Controller
    {

        #region  Param Initial
        string strPROG_ID = "MERP_TCF004"; //部門勞健保薪資計算
        string strMENU_ID = "MERP_TCF000";

        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        public MERP_FA_LaborHealthInsV1DBService m_LHInsDBV1Service = new MERP_FA_LaborHealthInsV1DBService();
        #endregion

        #region Action_View
        // GET: MERP_TCF000/MERP_TCF004
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出上傳的日記帳檔案
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);
            
            //抓出目前LHInsDBV1的有存在那些資料年月
            IEnumerable<FA_LaborHealthInsV1> l_oFaLHIns = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataList();

            var l_oRtnData = from DataVals in l_oFaLHIns
                             group DataVals by DataVals.DataYear + "/" + DataVals.DataMonth into grp
                             select new SelectListItem
                             {
                                 Text = grp.Key,
                                 Value = grp.Key
                             };

            SelectList l_oRtnSelListData = new SelectList(l_oRtnData, "Value", "Text");
            ViewData["LHInsV1DBYearMonth"] = l_oRtnSelListData;



            return View(l_oDataList);
        }

        // GET: MERP_TCF000/MERP_TCF004/Details/5
        public ActionResult Details(int id)
        {
            string l_sLHInsMonth = "";

            FileContent l_oSearchFile = m_FileDBService.FileContent_GetDataById(id);
            //載入上傳的Excel,並且匯入DataBase
            l_sLHInsMonth = MERP_LaborHealthExcelV1BIService.ImportExcelTo_FA_LaborHealthV1(l_oSearchFile);

            //顯示本年月的All
            List<FA_LaborHealthInsV1> l_oLHIns = MERP_LaborHealthExcelV1BIService.GetImportExcelInDB_YearMonthData(
                                                                    l_oSearchFile.DataYear,
                                                                    l_oSearchFile.DataMonth);


            //顯示本月的各分頁Data(300筆)
            List<FA_LaborHealthInsV1> l_oFaLHIns = MERP_LaborHealthExcelV1BIService.GetImportExcelInDB_YearMonthDataPage(l_oSearchFile.DataYear,
                                                                                                                   l_oSearchFile.DataMonth, 1);

            MERP_TCF001_Details01ViewModel l_oDetailsVM = new MERP_TCF001_Details01ViewModel();
            l_oDetailsVM.m_sYear = l_oSearchFile.DataYear;
            l_oDetailsVM.m_sMonth = l_oSearchFile.DataMonth;
            l_oDetailsVM.m_oFALaborHealthInsV1List = l_oFaLHIns;

            return View(l_oDetailsVM);
            //return View();
        }


        // GET: MERP_TCF000/MERP_TCF004/Create
        public ActionResult Create()
        {
            return View();
        }


        // GET: MERP_TCF000/MERP_TCF004/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }


        // GET: MERP_TCF000/MERP_TCF004/Delete/5
        public ActionResult Delete(int id)
        {
            int p_iFileID = id;
            //20201214 CCL+
            if (p_iFileID > 0)
            {
                //刪除實體檔
                string p_sDelFileUrl = m_FileDBService.FileContent_GetDataById(p_iFileID).Url;
                MERP_UploadBIService.DeleteFile(p_sDelFileUrl);
                //刪除資料庫FileContent紀錄
                m_FileDBService.FileContent_DBDeleteByID(p_iFileID);
            }

            return RedirectToAction("Index");
           
        }

        // GET: MERP_TCC000/MERP_TCC004/Journals/5
        public ActionResult Journals(int id, int year, int month)
        {

            //20210106 CCl Mod

            //string l_sMonth = DateTime.Now.Month.ToString();
            //string l_sYear = DateTime.Now.Year.ToString();
            string l_sMonth = month.ToString();
            string l_sYear = year.ToString();
            //顯示本月的各分頁Data(300筆)
            List<FA_LaborHealthInsV1> l_oFaLHIns = MERP_LaborHealthExcelV1BIService.GetImportExcelInDB_YearMonthDataPage(l_sYear,
                                                                                                               l_sMonth, 0);

            MERP_TCF001_Details01ViewModel l_oJournalsVM = new MERP_TCF001_Details01ViewModel();


            l_oJournalsVM.m_sYear = l_sYear;  //Year
            l_oJournalsVM.m_sMonth = l_sMonth; //Month
            l_oJournalsVM.m_oFALaborHealthInsV1List = l_oFaLHIns;

            //l_oJournalsVM.m_oSelShopList = l_RtnShopsData.ToList();
            //ViewData["FA_LHInsV1VM"] = l_oJournalsVM;

            return View(l_oJournalsVM);
            //return View();
        }

        public ActionResult DownFileList()
        {
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Dn", strPROG_ID);
            //20201229 CCL+ ReOrder Desc
            List<FileContent> l_oDataListDesc = l_oDataList.OrderByDescending(m => m.CreateTime).ToList();

            return View(l_oDataListDesc);
            //20201229 CCL- return View(l_oDataList);
        }

        public ActionResult DeleteDownFile(int id)
        {
            int p_iFileID = id;
            //20201214 CCL+
            if (p_iFileID > 0)
            {
                //刪除實體檔
                string p_sDelFileUrl = m_FileDBService.FileContent_GetDataById(p_iFileID).Url;
                MERP_UploadBIService.DeleteFile(p_sDelFileUrl);
                //刪除資料庫FileContent紀錄
                m_FileDBService.FileContent_DBDeleteByID(p_iFileID);
            }

            return RedirectToAction("DownFileList");
        }

        public ActionResult Download(int id)
        {
            int p_iFileID = id;
            string l_sDwnFileUrl = "";
            string l_sFileName = "";
            string l_sFileExten = "";
            //20201214 CCL+
            if (p_iFileID > 0)
            {
                FileContent l_oFileFinded = m_FileDBService.FileContent_GetDataById(p_iFileID);
                l_sDwnFileUrl = l_oFileFinded.Url;
                l_sFileName = l_oFileFinded.Name;
                l_sFileExten = l_oFileFinded.Type;
                l_sFileName += "." + l_sFileExten;
                //HttpContext.Response.Headers
                //Uri l_oUri = new Uri()


                try
                {
                    FileStream l_oStream = new FileStream(l_sDwnFileUrl, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return File(l_oStream, "application/octet-stream", l_sFileName); //MME 格式 可上網查 此為通用設定
                }
                catch (System.Exception)
                {
                    return Content("<script>alert('查無此檔案');window.close()</script>");
                }

            }

            return Content("<script>alert('查無此檔案');window.close()</script>");
            //return RedirectToAction("DownFileList");
        }

        // GET: MERP_TCC000/MERP_TCF004/LHInsV1DB_Delete
        public ActionResult LHInsV1DB_Delete()
        {
            //刪除DB
            //抓出目前DB有哪些年月Data
            IEnumerable<FA_LaborHealthInsV1> l_oFaLHIns = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataList();

            var l_oRtnData = from DataVals in l_oFaLHIns
                             group DataVals by DataVals.DataYear + "/" + DataVals.DataMonth into grp
                             select new SelectListItem
                             {
                                 Text = grp.Key,
                                 Value = grp.Key
                                 //Value = grp.Count().ToString()
                             };

            SelectList l_oRtnSelListData = new SelectList(l_oRtnData, "Value", "Text");

            return View(l_oRtnSelListData);
        }

        #endregion

        #region Action_DB
        // POST: MERP_TCF000/MERP_TCF004/Index
        [HttpPost]
        #region 查詢畫面送出(Index) [Submit]
        public ActionResult Index(string year_month, HttpPostedFileBase upload)
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            if (!string.IsNullOrEmpty(year_month))
            {
                DateStringProcess.Del_MonthZero(year_month, "-", "");
                string l_sYear = DateStringProcess.m_Year.ToString();
                string l_sMonth = DateStringProcess.m_Month.ToString();


                //執行上傳檔案           
                MERP_UploadBIService.UploadFile(upload, Server, strPROG_ID, l_sYear, l_sMonth);

                //只查出Upload及本程式分類File
                List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);


                return View(l_oDataList);
            }
            else
            {
                return View();
            }

        }
        #endregion

        // POST: MERP_TCC000/MERP_TCF004/Journals
        [HttpPost]
        #region Journals畫面送出(Journals) [Submit]
        public ActionResult Journals(FormCollection p_oForm)
        {


            int l_iYear = 0, l_iMonth = 0;
            string l_sYear = "", l_sMonth = "";          
            //字串"07"轉為int 7          
            string l_sShops = "";


            MERP_TCF004_JournalsOptions l_oProcessOptions = new MERP_TCF004_JournalsOptions();
            //資料年,月
            l_oProcessOptions.m_sDataYear = p_oForm["DataYear"];
            l_oProcessOptions.m_sDataMonth = p_oForm["DataMonth"];
            //到職日,離職日
            //if(!string.IsNullOrEmpty(p_oForm["OnJobDate"]))
            //    l_oProcessOptions.m_sOnJobDate = DateStringProcess.Del_MonthDayZero(p_oForm["OnJobDate"], "-", "/");
            //if (!string.IsNullOrEmpty(p_oForm["ResignDate"]))
            //    l_oProcessOptions.m_sResignDate = DateStringProcess.Del_MonthDayZero(p_oForm["ResignDate"], "-", "/");
            //開始日,結束日
            if (!string.IsNullOrEmpty(p_oForm["StartDate"]))
                l_oProcessOptions.m_sOnJobDate = DateStringProcess.Del_MonthDayZero(p_oForm["StartDate"], "-", "/");
            if (!string.IsNullOrEmpty(p_oForm["EndDate"]))
                l_oProcessOptions.m_sResignDate = DateStringProcess.Del_MonthDayZero(p_oForm["EndDate"], "-", "/");
            //員工
            l_oProcessOptions.m_sMemberName = p_oForm["MemberName"];
            //部門
            l_oProcessOptions.m_sShopName = p_oForm["ShopName"];
            
            //顯示本月的各分頁Data(300筆)
            List<FA_LaborHealthInsV1> l_oFaLHIns = 
                MERP_LaborHealthExcelV1BIService.ProcessImportExcelFromDB(l_oProcessOptions);

            if ((p_oForm["IsUseDetailSel"] != null) && (p_oForm["IsUseDetailSel"].Contains("on")))
            {
                //輸出明細版本
                //合計
                //MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptions(l_oProcessOptions, strPROG_ID, Server);
                //合計 V1版
                //MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV1(l_oProcessOptions, strPROG_ID, Server);
                //合計 V1_1版 新增勞健保小計
                //MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV1_1(l_oProcessOptions, strPROG_ID, Server);
                //合計 V1_2版 新增勞健保小計+新增勞保墊償小計
                MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV1_2(l_oProcessOptions, strPROG_ID, Server);

            } else
            {
                //輸出總計版本
                //合計 V2版
                //MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV2(l_oProcessOptions, strPROG_ID, Server);
                //合計 V2_1版 新增勞健保小計
                //MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV2_1(l_oProcessOptions, strPROG_ID, Server);
                //合計 V2_2版 新增勞健保小計+新增勞保墊償小計
                MERP_LaborHealthExcelV1BIService.SaveAsExcelByOptionsV2_2(l_oProcessOptions, strPROG_ID, Server);
            }
           

            //MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            //l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            //l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            //20201227 CCL- return View(l_oJournalsVM);
            //20201227 CCL+ 到檔案下載頁面
            return RedirectToAction("DownFileList");
        }
        #endregion

        /*
        // POST: MERP_TCF000/MERP_TCF004/LHInsV1DB_Delete
        [HttpPost]
        public ActionResult LHInsV1DB_Delete(string DataYear, string DataMonth)
        {

            if (!string.IsNullOrEmpty(DataYear) &&
               !string.IsNullOrEmpty(DataMonth))
            {
                int l_iDataYear = Convert.ToInt32(DataYear);
                int l_iDataMonth = Convert.ToInt32(DataMonth);

                m_LHInsDBV1Service.FA_LaborHealthInsV1_DBDeleteByYearMon(l_iDataYear, l_iDataMonth);
            }


            return View();
        }
        */

        // POST: MERP_TCF000/MERP_TCF004/LHInsV1DB_Delete
        [HttpPost]
        public ActionResult LHInsV1DB_Delete(string LHInsV1DBYearMonth)
        {

            string l_sDataYear = "";
            string l_sDataMonth = "";
            if(!string.IsNullOrEmpty(LHInsV1DBYearMonth) )
            {
                l_sDataYear = LHInsV1DBYearMonth.Substring(0, LHInsV1DBYearMonth.IndexOf("/"));
                l_sDataMonth = LHInsV1DBYearMonth.Substring(LHInsV1DBYearMonth.IndexOf("/") + 1);
            }
            

            if (!string.IsNullOrEmpty(l_sDataYear) &&
               !string.IsNullOrEmpty(l_sDataMonth))
            {
                int l_iDataYear = Convert.ToInt32(l_sDataYear);
                int l_iDataMonth = Convert.ToInt32(l_sDataMonth);

                m_LHInsDBV1Service.FA_LaborHealthInsV1_DBDeleteByYearMon(l_iDataYear, l_iDataMonth);
            }

            return RedirectToAction("LHInsV1DB_Delete");
            //return View();
        }


        // POST: MERP_TCF000/MERP_TCF004/Create
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

        // POST: MERP_TCF000/MERP_TCF004/Edit/5
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

        // POST: MERP_TCF000/MERP_TCF004/Delete/5
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
