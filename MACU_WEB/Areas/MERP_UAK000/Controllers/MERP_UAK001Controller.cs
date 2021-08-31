using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using MACU_WEB.BIServices;
using MACU_WEB.Areas.MERP_UAK000.ViewModels;
using MACU_WEB.Models._Base;
using System.IO;
using System.Diagnostics;


namespace MACU_WEB.Areas.MERP_UAK000.Controllers
{
    //督導Manager
    public class MERP_UAK001Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_UAK001"; //客製程式
        string strMENU_ID = "MERP_UAK000";

        public MERP_HR_ManagerInfoDBService m_ManagerDBService = new MERP_HR_ManagerInfoDBService();
        public MERP_StoreInfoDBService m_StoreInfoDBService = new MERP_StoreInfoDBService();
        #endregion


        #region Action_View
        // GET: MERP_UAK000/MERP_UAK001
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出督導Manager List
            List<HR_ManagerInfo> l_oDataList = m_ManagerDBService.HR_ManagerInfo_GetDataList();

            return View(l_oDataList);
        }

        /* 20210107 CCL-
        // GET: MERP_UAK000/MERP_UAK001/Edit/5
        public ActionResult Edit(int id)
        {
            //找出欲編輯督導
            HR_ManagerInfo l_oFindItem = m_ManagerDBService.HR_ManagerInfo_GetDataById(id);
            List<StoreInfo> l_oStoreInfos = m_StoreInfoDBService.StoreInfo_GetDataList();

            var l_oRtnShopSelItems = from DataVals in l_oStoreInfos                                    
                                     select new SelectListItem
                                     {
                                         Value = DataVals.SID,
                                         Text = DataVals.Name
                                     };

            MERP_UAK001_EditViewModel l_oManagerInfoVM = new MERP_UAK001_EditViewModel();
            l_oManagerInfoVM.m_oHRManager = l_oFindItem;
            l_oManagerInfoVM.m_oSelShopList = l_oRtnShopSelItems.ToList();

            return View(l_oManagerInfoVM);
        }
        */

        // GET: MERP_UAK000/MERP_UAK001/Edit/5
        public ActionResult Edit(int id)
        {
            //找出欲編輯督導
            HR_ManagerInfo l_oFindItem = m_ManagerDBService.HR_ManagerInfo_GetDataById(id);
            MERP_UAK001_Edit01ViewModel l_RtnData = new MERP_UAK001_Edit01ViewModel();

            l_RtnData.m_oHRManager = l_oFindItem; //督導
            var l_oAllShopsGroup = m_StoreInfoDBService.StoreInfo_GetDataGroupBySID();

            foreach (dynamic grp in l_oAllShopsGroup)
            {
                switch (grp.SIDTopChar)
                {
                    case "N":
                        //l_RtnData.m_NShopKey = grp.SIDTopChar;
                        l_RtnData.m_NShopKey = "北區";
                        l_RtnData.m_NShopCount = grp.GroupShopsCount;
                        var l_oRtnNShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelNShopList = l_oRtnNShopSelItems.ToList();
                        break;
                    case "C":
                        //l_RtnData.m_CShopKey = grp.SIDTopChar;
                        l_RtnData.m_CShopKey = "中區";
                        l_RtnData.m_CShopCount = grp.GroupShopsCount;
                        var l_oRtnCShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelCShopList = l_oRtnCShopSelItems.ToList();
                        break;
                    case "S":
                        //l_RtnData.m_SShopKey = grp.SIDTopChar;
                        l_RtnData.m_SShopKey = "南區";
                        l_RtnData.m_SShopCount = grp.GroupShopsCount;
                        var l_oRtnSShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelSShopList = l_oRtnSShopSelItems.ToList();
                        break;
                }
            }

            return View(l_RtnData);




            //List<StoreInfo> l_oStoreInfos = m_StoreInfoDBService.StoreInfo_GetDataList();

            //var l_oRtnShopSelItems = from DataVals in l_oStoreInfos
            //                         select new SelectListItem
            //                         {
            //                             Value = DataVals.SID,
            //                             Text = DataVals.Name
            //                         };

            //MERP_UAK001_EditViewModel l_oManagerInfoVM = new MERP_UAK001_EditViewModel();
            //l_oManagerInfoVM.m_oHRManager = l_oFindItem;
            //l_oManagerInfoVM.m_oSelShopList = l_oRtnShopSelItems.ToList();

            //return View(l_oManagerInfoVM);
        }

        // GET: MERP_UAK000/MERP_UAK001/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        /* 20210107 CCL-
        // GET: MERP_UAK000/MERP_UAK001/Create
        public ActionResult Create()
        {
            //List<HR_ManagerInfo> l_oDataList = m_ManagerDBService.HR_ManagerInfo_GetDataList();
            List<StoreInfo> l_oStoreInfos = m_StoreInfoDBService.StoreInfo_GetDataList();
            //l_oStoreInfos = l_oStoreInfos.GroupBy(m => m.SID.Substring(0,1)).ToList();

            var l_oRtnShopSelItems = from DataVals in l_oStoreInfos                                     
                                     select new SelectListItem
                                     {
                                         
                                         Value = DataVals.SID,
                                         Text = DataVals.Name                                           
                                     };

            List<SelectListItem> l_oViewDataList = l_oRtnShopSelItems.ToList();

            //SelectListItem l_oSelItems = new SelectListItem();
            //l_oSelItems.Text = "";
            //l_oSelItems.Value = "";

            return View(l_oViewDataList);
        }
        */

        // GET: MERP_UAK000/MERP_UAK001/Create
        public ActionResult Create()
        {
            //List<HR_ManagerInfo> l_oDataList = m_ManagerDBService.HR_ManagerInfo_GetDataList();
            ///List<StoreInfo> l_oStoreInfos = m_StoreInfoDBService.StoreInfo_GetDataList();
            //l_oStoreInfos = l_oStoreInfos.GroupBy(m => m.SID.Substring(0,1)).ToList();
            MERP_UAK001_Edit01ViewModel l_RtnData = new MERP_UAK001_Edit01ViewModel();

            var l_oAllShopsGroup = m_StoreInfoDBService.StoreInfo_GetDataGroupBySID();
            
            foreach(dynamic grp in l_oAllShopsGroup)
            {
                switch(grp.SIDTopChar)
                {
                    case "N":
                        //l_RtnData.m_NShopKey = grp.SIDTopChar;
                        l_RtnData.m_NShopKey = "北區";
                        l_RtnData.m_NShopCount = grp.GroupShopsCount;
                        var l_oRtnNShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID 
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelNShopList = l_oRtnNShopSelItems.ToList();
                        break;
                    case "C":
                        //l_RtnData.m_CShopKey = grp.SIDTopChar;
                        l_RtnData.m_CShopKey = "中區";
                        l_RtnData.m_CShopCount = grp.GroupShopsCount;
                        var l_oRtnCShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelCShopList = l_oRtnCShopSelItems.ToList();
                        break;
                    case "S":
                        //l_RtnData.m_SShopKey = grp.SIDTopChar;
                        l_RtnData.m_SShopKey = "南區";
                        l_RtnData.m_SShopCount = grp.GroupShopsCount;
                        var l_oRtnSShopSelItems = from DataVals in (IEnumerable<StoreInfo>)grp.GrpObj
                                                  orderby DataVals.SID
                                                  select new SelectListItem
                                                  {
                                                      Text = DataVals.Name,
                                                      Value = DataVals.SID
                                                  };
                        l_RtnData.m_oSelSShopList = l_oRtnSShopSelItems.ToList();
                        break;
                }
            }

            return View(l_RtnData);

            ///var l_oRtnShopSelItems = from DataVals in l_oStoreInfos
            ///                         select new SelectListItem
            ///                         {

                                      ///                             Value = DataVals.SID,
                                      ///                             Text = DataVals.Name
                                      ///                         };

                                      ///List<SelectListItem> l_oViewDataList = l_oRtnShopSelItems.ToList();

                                      //SelectListItem l_oSelItems = new SelectListItem();
                                      //l_oSelItems.Text = "";
                                      //l_oSelItems.Value = "";

            //return View(l_oViewDataList);
        }

        // GET: MERP_UAK000/MERP_UAK001/Delete/5
        public ActionResult Delete(int id)
        {
            m_ManagerDBService.HR_ManagerInfo_DBDeleteByID(id);

            //return View();
            return RedirectToAction("Index");
        }

        //20210205 CCL+ Str Testing ///////////////////////////////////////////////
        public JsonResult QueryTable()
        {
            //取得所有資料
            //列出督導Manager List
            var l_oDataList = m_ManagerDBService.HR_ManagerInfo_GetDataList();

            //組成jqGrid要讀取的資料
            var jsonData = new
            {
                rows = l_oDataList   //jqGrid呈現在表格中的資料
            };

            //回傳
            return Json(jsonData, JsonRequestBehavior.AllowGet);

        }
        //20210205 CCL+ End Testing ///////////////////////////////////////////////

        #endregion

        #region Action_DB
        // POST: MERP_UAK000/MERP_UAK001/Create
        [HttpPost]
        public ActionResult Create(FormCollection p_oForm)
        {
            try
            {
                /*
                string l_sShops = "";
                List<string> l_oTmpStr;
                string[] l_aryShops = p_oForm["SelShopItem"].Trim().Split(',');
                l_oTmpStr = l_aryShops.ToList().Where(m => m != "false").ToList();
                
                foreach (string item in l_oTmpStr)
                {
                    l_sShops += item + ",";
                    
                }
                l_sShops = l_sShops.Substring(0, l_sShops.Length - 1);
                */

                string l_sShops = p_oForm["CheckedItems"];

                // TODO: Add insert logic here
                m_ManagerDBService.HR_ManagerInfo_DBCreate(
                   p_oForm["ManagerID"].Trim(),
                   p_oForm["ManagerName"].Trim(),
                   p_oForm["ManagerNickNa"].Trim(),
                   p_oForm["ManagerTelPhone"].Trim(),
                   p_oForm["ManagerMobiPhone"].Trim(),
                   p_oForm["ManageBranchID"].Trim(),
                   l_sShops.Trim()
                   );

                //return RedirectToAction("Create");
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }



        // POST: MERP_UAK000/MERP_UAK001/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection p_oForm)
        {
            try
            {
                // TODO: Add update logic here
                string l_sShops = "";
               
                //選擇單店家
                l_sShops = p_oForm["CheckedShopItems"];
                l_sShops = l_sShops.Trim();

                m_ManagerDBService.HR_ManagerInfo_DBUpdate(id,
                                                    p_oForm["ManagerID"].Trim(),
                                                    p_oForm["ManagerName"].Trim(),
                                                    p_oForm["ManagerNickNa"].Trim(),
                                                    p_oForm["ManagerTelPhone"].Trim(),
                                                    p_oForm["ManagerMobiPhone"].Trim(),
                                                    p_oForm["ManageBranchID"].Trim(),
                                                    l_sShops);

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }




        // POST: MERP_UAK000/MERP_UAK001/Delete/5
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
