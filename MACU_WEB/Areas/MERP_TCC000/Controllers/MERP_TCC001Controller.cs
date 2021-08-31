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

namespace MACU_WEB.Areas.MERP_TCC000.Controllers
{
    //將總帳(T)->會計帳簿查詢列印(C)->日記帳(C) 載入本月日記帳到DB內
    public class MERP_TCC001Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TCC001"; //客製程式
        string strMENU_ID = "MERP_TCC000";

        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        //20210106 CCL+
        public MERP_StoreInfoDBService m_StoreInfoDBService = new MERP_StoreInfoDBService();
        public MERP_HR_ManagerInfoDBService m_ManagerDBService = new MERP_HR_ManagerInfoDBService();
        //20210204 CCL+ 直合營設定
        public MERP_StoreGroupSetDBService m_StoreGroupSetDBService = new MERP_StoreGroupSetDBService();
        #endregion


        #region Action_View
        // GET: MERP_TCC000/MERP_TCC001
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出上傳的日記帳檔案
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);
           
            return View(l_oDataList);
        }

        /* 20201223 CCL-
        // GET: MERP_TCC000/MERP_TCC001/Details/5
        public ActionResult Details(int id)
        {
            string l_sAccountPeriod = "";        

            FileContent l_oSearchFile = m_FileDBService.FileContent_GetDataById(id);
            //載入上傳的Excel,並且匯入DataBase
            l_sAccountPeriod = MERP_ExcelBIService.ImportExcelToFA_DayContentDB(l_oSearchFile.Url);

            //顯示本月的All
            //List<FA_FaJournal> l_dFaJournal = MERP_ExcelBIService.GetImportExcelInDB_PeriodData(l_sAccountPeriod);
            //顯示本月的各分頁Data(300筆)
            List<FA_FaJournal> l_dFaJournal = MERP_ExcelBIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 1);

            MERP_TCC001_DetailsViewModel l_oDetailsVM = new MERP_TCC001_DetailsViewModel();
            l_oDetailsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oDetailsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oDetailsVM);
        }
        */

        // GET: MERP_TCC000/MERP_TCC001/Details/5
        public ActionResult Details(int id)
        {
            string l_sAccountPeriod = "";
            //載入的會計年分
            string l_sFiscalYear = "";
            string[] l_aryYearMonth = null;

            FileContent l_oSearchFile = m_FileDBService.FileContent_GetDataById(id);
            //載入上傳的Excel,並且匯入DataBase
            l_aryYearMonth = MERP_ExcelV1BIService.ImportExcelToFA_DayContentDB_V1(l_oSearchFile.Url);
            l_sFiscalYear = l_aryYearMonth[0];
            l_sAccountPeriod = l_aryYearMonth[1];

            //顯示本月的All
            //List<FA_FaJournal> l_dFaJournal = MERP_ExcelBIService.GetImportExcelInDB_PeriodData(l_sAccountPeriod);
            //顯示本月的各分頁Data(300筆)
            //20210204 CCL- List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 1);
            List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_YearPeriodDataPage(l_sFiscalYear, 
                                                                                                        l_sAccountPeriod, 1);

            MERP_TCC001_DetailsViewModel l_oDetailsVM = new MERP_TCC001_DetailsViewModel();
            l_oDetailsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oDetailsVM.m_sFiscalYear = l_sFiscalYear; //20210204 CCL+
            l_oDetailsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oDetailsVM);
        }

        /* 20201223 CCL-
        // GET: MERP_TCC000/MERP_TCC001/Details/5
        public ActionResult Journals(int id)
        {

            string l_sAccountPeriod = "7";
            //顯示本月的各分頁Data(300筆)
            List<FA_FaJournal> l_dFaJournal = MERP_ExcelBIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 0);

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oJournalsVM);
        }
        */

        /*  20210106 CCL-
        // GET: MERP_TCC000/MERP_TCC001/Journals/5
        public ActionResult Journals(int id)
        {

            string l_sAccountPeriod = "7";
            //顯示本月的各分頁Data(300筆)
            List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 0);

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oJournalsVM);
        }
        */

        /* 20210107 CCL-
        // GET: MERP_TCC000/MERP_TCC001/Journals/5
        public ActionResult Journals(int id)
        {
            //20210106 CCl Mod

            string l_sAccountPeriod = DateTime.Now.Month.ToString();
            //顯示本月的各分頁Data(300筆)
            List<FA_JournalV1> l_oFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 0);
            List<HR_ManagerInfo> l_oHRManager = m_ManagerDBService.HR_ManagerInfo_GetDataList(); //督導
            List<StoreInfo> l_oSelShops = m_StoreInfoDBService.StoreInfo_GetDataList(); //部門

            var l_RtnShopsData = from DataVals in l_oSelShops
                                 where DataVals.Disabled == false
                                 select new SelectListItem
                                 {
                                     Text = DataVals.Name,
                                     Value = DataVals.SID
                                 };

            var l_RtnManagerData = from DataVals in l_oHRManager
                                 where DataVals.IsValid == 1
                                 select new SelectListItem
                                 {
                                     Text = DataVals.ManagerNickNa,
                                     Value = DataVals.ManageShopList
                                     //,BranchID = DataVals.ManageBranchID
                                 };

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_oFaJournal;
            l_oJournalsVM.m_oHRManagerList = l_RtnManagerData.ToList();
            l_oJournalsVM.m_oSelShopList = l_RtnShopsData.ToList();

            return View(l_oJournalsVM);
        }
        */

        public void GetStoreInfoGroups(MERP_TCC001_Details01ViewModel p_oRtnData)
        {
            var l_oAllShopsGroup = m_StoreInfoDBService.StoreInfo_GetDataGroupBySID();

            foreach (dynamic grp in l_oAllShopsGroup)
            {
                switch (grp.SIDTopChar)
                {
                    case "N":
                        //l_RtnData.m_NShopKey = grp.SIDTopChar;
                        p_oRtnData.m_NShopKey = "北區";
                        p_oRtnData.m_NShopCount = grp.GroupShopsCount;
                        //依店代碼排序
                        var l_oRtnNShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        p_oRtnData.m_oSelNShopList = l_oRtnNShopSelItems.ToList();                        
                        break;
                    case "C":
                        //l_RtnData.m_CShopKey = grp.SIDTopChar;
                        p_oRtnData.m_CShopKey = "中區";
                        p_oRtnData.m_CShopCount = grp.GroupShopsCount;
                        //依店代碼排序
                        var l_oRtnCShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        p_oRtnData.m_oSelCShopList = l_oRtnCShopSelItems.ToList();
                        break;
                    case "S":
                        //l_RtnData.m_SShopKey = grp.SIDTopChar;
                        p_oRtnData.m_SShopKey = "南區";
                        p_oRtnData.m_SShopCount = grp.GroupShopsCount;
                        //依店代碼排序
                        var l_oRtnSShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        p_oRtnData.m_oSelSShopList = l_oRtnSShopSelItems.ToList();
                        break;
                }
            }

            
        }

        // 20210204 CCL- 改用選擇全部
        //20210107 CCL+ 改分群
        // GET: MERP_TCC000/MERP_TCC001/Journals/5
        public ActionResult Journals(int id)
        {

            //20210106 CCl Mod

            string l_sAccountPeriod = DateTime.Now.Month.ToString();
            //20210204 CCL+ 
            string l_sFiscalYear = DateTime.Now.Year.ToString();

            //顯示本月的各分頁Data(300筆)
            //20210204 CCL- List<FA_JournalV1> l_oFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_PeriodDataPage(l_sAccountPeriod, 0);
            //20210204 CCL* 用年月抓
            List<FA_JournalV1> l_oFaJournal = MERP_ExcelV1BIService.GetImportExcelInDB_YearPeriodDataPage(l_sFiscalYear,
                                                                                                     l_sAccountPeriod, 0);
            List<HR_ManagerInfo> l_oHRManager = m_ManagerDBService.HR_ManagerInfo_GetDataList(); //督導
            //20210204 CCL+ 加上直合營設定List
            List<StoreGroupSet> l_oStoreGroupSetList = m_StoreGroupSetDBService.StoreGroupSet_GetDataList();

            MERP_TCC001_Details01ViewModel l_oJournalsVM = new MERP_TCC001_Details01ViewModel();
            //List<StoreInfo> l_oSelShops = m_StoreInfoDBService.StoreInfo_GetDataList(); //部門
            //var l_RtnShopsData = from DataVals in l_oSelShops
            //                     where DataVals.Disabled == false
            //                     select new SelectListItem
            //                     {
            //                         Text = DataVals.Name,
            //                         Value = DataVals.SID
            //                     };

            //取得StoreInfo List
            GetStoreInfoGroups(l_oJournalsVM);

            //取得ManagerInfo List
            var l_RtnManagerData = from DataVals in l_oHRManager
                                   where DataVals.IsValid == 1
                                   select new SelectListItem
                                   {
                                       Text = DataVals.ManagerNickNa,
                                       Value = DataVals.ManageShopList
                                       //,BranchID = DataVals.ManageBranchID
                                   };

            //20210225 CCL+ 取出直合營設定
            var l_oRtnTmpGroupSetData = from DataVals in l_oStoreGroupSetList
                                       where DataVals.IsValid == 1
                                       select new SelectListItem
                                       {
                                           Text = DataVals.StoreGroupName,
                                           Value = DataVals.StoreGroupSIDList

                                       };
            SelectList l_oRtnGroupSetDataSelList = new SelectList(l_oRtnTmpGroupSetData, "Value", "Text");


            l_oJournalsVM.m_FiscalYear = l_sFiscalYear; //20210204 CCL+
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod; //Month
            l_oJournalsVM.m_FaJournalList = l_oFaJournal;
            l_oJournalsVM.m_oHRManagerList = l_RtnManagerData.ToList(); //督導List           
            //20210204 CCL+ 直合營設定List
            //20210225 CCL- l_oJournalsVM.m_oStoreGroupSetList = l_oStoreGroupSetList;
            l_oJournalsVM.m_oStoreGroupSetSelList = l_oRtnGroupSetDataSelList;

            //l_oJournalsVM.m_oSelShopList = l_RtnShopsData.ToList();

            return View(l_oJournalsVM);
        }
        

       

        // GET: MERP_TCC000/MERP_TCC001/Journals
        public ActionResult GoJournals(string type)
        {
            return RedirectToAction("Journals");
            
            //return View();
        }

        // GET: MERP_TCC000/MERP_TCC001/Create
        public ActionResult Create()
        {
            return View();
        }

        // GET: MERP_TCC000/MERP_TCC001/Delete/5
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

        //20201227 CCL+
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
        #endregion



        #region Action_DB
        [HttpPost]
        #region 查詢畫面送出(Index) [Submit]
        public ActionResult Index(HttpPostedFileBase upload)
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //執行上傳檔案           
            MERP_UploadBIService.UploadFile(upload, Server, strPROG_ID);

            //只查出Upload及本程式分類File
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);


            return View(l_oDataList);
        }
        #endregion

        /* 20201223 CCL-
        // POST: MERP_TCC000/MERP_TCC001/Journals
        [HttpPost]
        #region Journals畫面送出(Journals) [Submit]
        public ActionResult Journals(FormCollection p_oForm)
        {
            int l_iAccountPeriod = 0;
            string l_sAccountPeriod = "";
            //string l_sAccountPeriod = p_oForm["na_StartDate"];
            //字串"07"轉為int 7
            //l_iAccountPeriod = Convert.ToInt32(l_sAccountPeriod.Substring(l_sAccountPeriod.IndexOf('-')+1, 2));



            MERP_ProcessExcelOptions l_oProcessOptions = new MERP_ProcessExcelOptions();
            l_oProcessOptions.m_sStartDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_StartDate"],"-","/");
            l_iAccountPeriod = DateStringProcess.m_Month;
            l_sAccountPeriod = l_iAccountPeriod.ToString();
            l_oProcessOptions.m_sEndDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_EndDate"], "-", "/");
            //l_oProcessOptions.m_sStartDate = p_oForm["na_StartDate"].Replace('-', '/');
            //l_oProcessOptions.m_sEndDate = p_oForm["na_EndDate"].Replace('-', '/');
            l_oProcessOptions.m_sAccountPeriod = l_sAccountPeriod;
            l_oProcessOptions.m_sShop = "SD001"; //"ND004";

            //顯示本月的各分頁Data(300筆)
            List<FA_FaJournal> l_dFaJournal = MERP_ExcelBIService.ProcessImportExcelFromDB(l_oProcessOptions);
            MERP_ExcelBIService.SaveAsExcelByOptions(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions2(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions3(l_oProcessOptions, strPROG_ID, Server);

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oJournalsVM);


        }
        */

        /* 20201227 CCL-
        // POST: MERP_TCC000/MERP_TCC001/Journals
        [HttpPost]
        #region Journals畫面送出(Journals) [Submit]
        public ActionResult Journals(FormCollection p_oForm)
        {
            int l_iAccountPeriod = 0;
            string l_sAccountPeriod = "";
            //string l_sAccountPeriod = p_oForm["na_StartDate"];
            //字串"07"轉為int 7
            //l_iAccountPeriod = Convert.ToInt32(l_sAccountPeriod.Substring(l_sAccountPeriod.IndexOf('-')+1, 2));



            MERP_ProcessExcelOptions l_oProcessOptions = new MERP_ProcessExcelOptions();
            l_oProcessOptions.m_sStartDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_StartDate"], "-", "/");
            l_iAccountPeriod = DateStringProcess.m_Month;
            l_sAccountPeriod = l_iAccountPeriod.ToString();
            l_oProcessOptions.m_sEndDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_EndDate"], "-", "/");
            //l_oProcessOptions.m_sStartDate = p_oForm["na_StartDate"].Replace('-', '/');
            //l_oProcessOptions.m_sEndDate = p_oForm["na_EndDate"].Replace('-', '/');
            l_oProcessOptions.m_sAccountPeriod = l_sAccountPeriod;
            l_oProcessOptions.m_sShop = "SD007";//"SD023";//"SD002";//"SD010";//"SD024";//"SD022";//"SD004";//"SD012";//"SD017"; //"ND004";

            //顯示本月的各分頁Data(300筆)
            List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.ProcessImportExcelFromDB(l_oProcessOptions);
            MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions6(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions5(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions2(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions3(l_oProcessOptions, strPROG_ID, Server);

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            return View(l_oJournalsVM);

        }
        */

        /* 20210106 CCL-
        // POST: MERP_TCC000/MERP_TCC001/Journals
        [HttpPost]
        #region Journals畫面送出(Journals) [Submit]
        public ActionResult Journals(FormCollection p_oForm)
        {
            int l_iAccountPeriod = 0;
            string l_sAccountPeriod = "";
            //string l_sAccountPeriod = p_oForm["na_StartDate"];
            //字串"07"轉為int 7
            //l_iAccountPeriod = Convert.ToInt32(l_sAccountPeriod.Substring(l_sAccountPeriod.IndexOf('-')+1, 2));



            MERP_ProcessExcelOptions l_oProcessOptions = new MERP_ProcessExcelOptions();
            l_oProcessOptions.m_sStartDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_StartDate"], "-", "/");
            l_iAccountPeriod = DateStringProcess.m_Month;
            l_sAccountPeriod = l_iAccountPeriod.ToString();
            l_oProcessOptions.m_sEndDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_EndDate"], "-", "/");
            //l_oProcessOptions.m_sStartDate = p_oForm["na_StartDate"].Replace('-', '/');
            //l_oProcessOptions.m_sEndDate = p_oForm["na_EndDate"].Replace('-', '/');
            l_oProcessOptions.m_sAccountPeriod = l_sAccountPeriod;
            l_oProcessOptions.m_sShop = p_oForm["na_ShopsSel"];//"SD007";//"SD023";//"SD002";//"SD010";//"SD024";//"SD022";//"SD004";//"SD012";//"SD017"; //"ND004";
            l_oProcessOptions.m_sManager = p_oForm["SelManagerNo"];

            //顯示本月的各分頁Data(300筆)
            List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.ProcessImportExcelFromDB(l_oProcessOptions);
            //科目名稱List唯一 ,加上合計
            MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V5(l_oProcessOptions, strPROG_ID, Server);
            //科目名稱List唯一
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V4(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V3(l_oProcessOptions, strPROG_ID, Server);
            //不顯示AccountNo 版本,改標頭&置中
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V2(l_oProcessOptions, strPROG_ID, Server);
            //不顯示AccountNo 版本
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V1(l_oProcessOptions, strPROG_ID, Server);
            //顯示AccountNo 版本
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7(l_oProcessOptions, strPROG_ID, Server);

            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions6(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelV1BIService.SaveAsExcelV1ByOptions5(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions2(l_oProcessOptions, strPROG_ID, Server);
            //MERP_ExcelBIService.SaveAsExcelByOptions3(l_oProcessOptions, strPROG_ID, Server);

            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            //20201227 CCL- return View(l_oJournalsVM);
            //20201227 CCL+ 到檔案下載頁面
            return RedirectToAction("DownFileList");
        }
        */

        //20210106 CCL+
        // POST: MERP_TCC000/MERP_TCC001/Journals
        [HttpPost]
        #region Journals畫面送出(Journals) [Submit]
        public ActionResult Journals(FormCollection p_oForm)
        {


            int l_iAccountPeriod = 0;
            string l_sAccountPeriod = "";
            //20210204 CCL+
            int l_iFiscalYear = 0;
            string l_sFiscalYear = ""; 
            //string l_sAccountPeriod = p_oForm["na_StartDate"];
            //字串"07"轉為int 7
            //l_iAccountPeriod = Convert.ToInt32(l_sAccountPeriod.Substring(l_sAccountPeriod.IndexOf('-')+1, 2));
            string l_sShops = "";

            if ((p_oForm["IsUseManSel"] != null) && (p_oForm["IsUseManSel"].Contains("on")))
            {
                //使用督導
                l_sShops = p_oForm["CheckedManagerItems"];
            } else
            {
                //選擇單店家
                l_sShops = p_oForm["CheckedShopItems"];
            }



            MERP_ProcessExcelOptions l_oProcessOptions = new MERP_ProcessExcelOptions();
            l_oProcessOptions.m_sStartDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_StartDate"], "-", "/");
            l_iAccountPeriod = DateStringProcess.m_Month;
            l_sAccountPeriod = l_iAccountPeriod.ToString();
            //20210204 CCL+ 以年月抓取
            l_iFiscalYear = DateStringProcess.m_Year;
            l_sFiscalYear = l_iFiscalYear.ToString();

            l_oProcessOptions.m_sEndDate = DateStringProcess.Del_MonthDayZero(p_oForm["na_EndDate"], "-", "/");
            //l_oProcessOptions.m_sStartDate = p_oForm["na_StartDate"].Replace('-', '/');
            //l_oProcessOptions.m_sEndDate = p_oForm["na_EndDate"].Replace('-', '/');
            l_oProcessOptions.m_sAccountPeriod = l_sAccountPeriod;
            l_oProcessOptions.m_sFiscalYear = l_sFiscalYear; //20210204 CCL+ 以年月抓取
            l_oProcessOptions.m_sShop = l_sShops;
            //"SD007";//"SD023";//"SD002";//"SD010";//"SD024";//"SD022";//"SD004";//"SD012";//"SD017"; //"ND004";
            l_oProcessOptions.m_sManager = p_oForm["CheckedManagerNames"]; //p_oForm["SelManagerNo"];

            //顯示本月的各分頁Data(300筆)
            List<FA_JournalV1> l_dFaJournal = MERP_ExcelV1BIService.ProcessImportExcelFromDB(l_oProcessOptions);
            //科目名稱List唯一 ,加上合計
            MERP_ExcelV1BIService.SaveAsExcelV1ByOptions7V5(l_oProcessOptions, strPROG_ID, Server);
           
            MERP_TCC001_DetailsViewModel l_oJournalsVM = new MERP_TCC001_DetailsViewModel();
            l_oJournalsVM.m_AccountPeroid = l_sAccountPeriod;
            l_oJournalsVM.m_FaJournalList = l_dFaJournal;

            //20201227 CCL- return View(l_oJournalsVM);
            //20201227 CCL+ 到檔案下載頁面
            return RedirectToAction("DownFileList");
        }

        #endregion

        // POST: MERP_TCC000/MERP_TCC001/Create
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

        // GET: MERP_TCC000/MERP_TCC001/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC001/Edit/5
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


        // POST: MERP_TCC000/MERP_TCC001/Delete/5
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
