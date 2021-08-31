using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using MACU_WEB.BIServices;
using MACU_WEB.Areas.MERP_TCC000.ViewModels;
using MACU_WEB.Areas.MERP_UAJ000.ViewModels;
using MACU_WEB.Models._Base;
using System.IO;
using System.Data;
using System.Diagnostics;

namespace MACU_WEB.Areas.MERP_UAJ000.Controllers
{
    public class MERP_UAJ001Controller : Controller
    {
        //部門建立管理作業

        #region  Param Initial
        string strPROG_ID = "MERP_UAJ001"; //客製程式
        string strMENU_ID = "MERP_UAJ000";

        public MERP_StoreInfoDBService m_StoreDBService = new MERP_StoreInfoDBService();
        public MERP_StoreGroupSetDBService m_StoreGroupSetDBService = new MERP_StoreGroupSetDBService();
        #endregion

        #region Action_View
        // GET: MERP_UAJ000/MERP_UAJ001
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //列出部門資訊
            List<StoreInfo> l_oDataList = m_StoreDBService.StoreInfo_GetDataList();


            return View(l_oDataList);
        }

        // GET: MERP_UAJ000/MERP_UAJ001/UpdFromWebHQCenter
        //public ActionResult UpdFromWebHQCenter()
        //{

        //}

        // GET: MERP_UAJ000/MERP_UAJ001/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MERP_UAJ000/MERP_UAJ001/Create
        public ActionResult Create()
        {
            

            return View();
        }

        // GET: MERP_UAJ000/MERP_UAJ001/Edit/5
        public ActionResult Edit(int id)
        {
            StoreInfo l_oRtnData =  m_StoreDBService.StoreInfo_GetDataById(id);
            //20210107 CCL Mod 日期要改回 XXXX-X-X
            string l_sBeginDate = l_oRtnData.BeginDate;
            //l_sBeginDate = DateStringProcess.Del_MonthDayZero(l_sBeginDate, "/", "-");
            l_sBeginDate = DateStringProcess.ReStore_MonthDayZero(l_sBeginDate,"/","-");
            l_oRtnData.BeginDate = l_sBeginDate;


            return View(l_oRtnData);
        }

        // GET: MERP_UAJ000/MERP_UAJ001/Delete/5
        public ActionResult Delete(int id)
        {
            m_StoreDBService.StoreInfo_DBDeleteByID(id);

            return RedirectToAction("Index");
        }


        public void GetStoreInfoGroups(MERP_UAJ001_GroupSetViewModel p_oRtnData)
        {
            var l_oAllShopsGroup = m_StoreDBService.StoreInfo_GetDataGroupBySID();

            foreach (dynamic grp in l_oAllShopsGroup)
            {
                switch (grp.SIDTopChar)
                {
                    case "N":
                        //l_RtnData.m_NShopKey = grp.SIDTopChar;
                        p_oRtnData.m_sNShopKey = "北區";
                        p_oRtnData.m_iNShopCount = grp.GroupShopsCount;
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
                        p_oRtnData.m_sCShopKey = "中區";
                        p_oRtnData.m_iCShopCount = grp.GroupShopsCount;
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
                        p_oRtnData.m_sSShopKey = "南區";
                        p_oRtnData.m_iSShopCount = grp.GroupShopsCount;
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



        //設定直營,合營,加盟 Group_ID = 0:一般(未分類) 1:直營, 2:合營, 3:加盟
        // GET: MERP_UAJ000/MERP_UAJ001/GroupSet_Create
        public ActionResult GroupSet_Create()
        {

           
            //取出所有設定
            List<StoreGroupSet> l_oSGSetList = m_StoreGroupSetDBService.StoreGroupSet_GetDataList();

            MERP_UAJ001_GroupSetViewModel l_oJournalsVM = new MERP_UAJ001_GroupSetViewModel();
            //分群
            GetStoreInfoGroups(l_oJournalsVM);
           
            //取出直合營設定SelectItem
            SelectList l_oSelGroupList = MERP_StoreInfoGroups.GetStoreGroupSetSelList();
            //20210225 CCL+ 取出直合營設定Type區域 SelectItem
            SelectList l_oSelGroupTypeList = MERP_StoreInfoGroups.GetStoreGroupSetTypeSelList();

            //取出直合營設定SelectItem
            l_oJournalsVM.m_oStoreInfoGroup = l_oSelGroupList;
            //20210225 CCL+ 取出直合營設定Type區域 SelectItem
            l_oJournalsVM.m_oStoreInfoGroupSetType = l_oSelGroupTypeList;

            l_oJournalsVM.m_oStoreGroupSetList = l_oSGSetList;
            

            //return View();
            return View(l_oJournalsVM);
        }

        //設定直營,合營,加盟 Group_ID = 0:一般(未分類) 1:直營, 2:合營, 3:加盟
        // GET: MERP_UAJ000/MERP_UAJ001/GroupSet_Delete
        //20210225 CCL-  public ActionResult GroupSet_Delete(int no)
        public ActionResult GroupSet_Delete(int no, string type)
        {

            try
            {

                //m_StoreGroupSetDBService.StoreGroupSet_DBDeleteByID(id);
                //Primary Key改用SGNo, 要改用DeleteBySGNo
                //20210225 CCL- m_StoreGroupSetDBService.StoreGroupSet_DBDeleteByGroupNo(no);
                //20210225 CCL+
                m_StoreGroupSetDBService.StoreGroupSet_DBDeleteByGroupNoType(no, type);

            }
            catch (Exception ex)
            {

                string l_sErrMsg = ex.Message;
                return Content("<script>confirm('錯誤訊息: " + l_sErrMsg + "');window.close()</script>");

            }

            //return RedirectToAction("GroupSet_Create");
            return RedirectToAction("GroupSet_Index");
        }

        //設定直營,合營,加盟 Group_ID = 0:一般(未分類) 1:直營, 2:合營, 3:加盟
        // GET: MERP_UAJ000/MERP_UAJ001/GroupSet_Index
        public ActionResult GroupSet_Index()
        {
            //列表
            //取出所有直合營加盟List
            List<StoreGroupSet> l_oStoreGrpSetList = m_StoreGroupSetDBService.StoreGroupSet_GetDataList();


            return View(l_oStoreGrpSetList);
        }

        //設定直營,合營,加盟 Group_ID = 0:一般(未分類) 1:直營, 2:合營, 3:加盟
        // GET: MERP_UAJ000/MERP_UAJ001/GroupSet_Edit
        //public ActionResult GroupSet_Edit(int no)
        public ActionResult GroupSet_Edit(int no, string type)
        {
            //找出該Item
            //StoreGroupSet l_oFindItem = m_StoreGroupSetDBService.StoreGroupSet_GetDataById(id);
            //20210225 CCL- StoreGroupSet l_oFindItem = m_StoreGroupSetDBService.StoreGroupSet_GetDataByGroupNo(no);
            StoreGroupSet l_oFindItem = m_StoreGroupSetDBService.StoreGroupSet_GetDataByGroupNoType(no, type);

            MERP_UAJ001_GroupSetViewModel l_oJournalsVM = new MERP_UAJ001_GroupSetViewModel();
            //店家 分群
            GetStoreInfoGroups(l_oJournalsVM);

            //取出直合營設定SelectItem
            SelectList l_oSelGroupList = MERP_StoreInfoGroups.GetStoreGroupSetSelList();
            //20210225 CCL+ 取出直合營設定Type區域 SelectItem
            SelectList l_oSelGroupTypeList = MERP_StoreInfoGroups.GetStoreGroupSetTypeSelList();

            l_oJournalsVM.m_oStoreInfoGroup = l_oSelGroupList;
            //20210225 CCL+ 取出直合營設定Type區域 SelectItem
            l_oJournalsVM.m_oStoreInfoGroupSetType = l_oSelGroupTypeList;
            l_oJournalsVM.m_oToEditItem = l_oFindItem;


            return View(l_oJournalsVM);
        }

        #endregion


        #region Action_DB
        // POST: MERP_UAJ000/MERP_UAJ001
        [HttpPost]
        public ActionResult Index(FormCollection p_oForm)
        {

            //從WebHQCenter抓資訊存入LocalDB
            m_StoreDBService.StoreInfo_SqlDBChkUpdate();
            

            return RedirectToAction("Index");
            
        }

        // POST: MERP_UAJ000/MERP_UAJ001/Create
        [HttpPost]
        public ActionResult Create(FormCollection p_oForm)
        {
            try
            {
                // TODO: Add insert logic here
                m_StoreDBService.StoreInfo_DBCreate(
                                                   p_oForm["SID"].Trim(),
                                                   p_oForm["Name"].Trim(),
                                                   Convert.ToInt32(p_oForm["Kind"]),
                                                   p_oForm["Memo"].Trim(),
                                                   Convert.ToInt32(p_oForm["Group_ID"]),
                                                   p_oForm["BranchArea_ID"].Trim(),
                                                   p_oForm["OrderArea_ID"].Trim(),                                                  
                                                   p_oForm["BeginDate"].Trim(),
                                                   p_oForm["TelPhone"].Trim(),
                                                   p_oForm["FaxPhone"].Trim(),
                                                   p_oForm["CellPhone"].Trim(),
                                                   p_oForm["Contact"].Trim(),
                                                   p_oForm["Address"].Trim(),                                                  
                                                   p_oForm["MobileStoreID"].Trim(),
                                                   p_oForm["ManageKind"].Trim()
                                                   );


                return RedirectToAction("Index");
            }
            catch(Exception ex)
            {
                string l_sErrMsg = ex.Message;
                return View();
            }
        }



        // POST: MERP_UAJ000/MERP_UAJ001/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection p_oForm)
        {
            try
            {
                // TODO: Add update logic here
                m_StoreDBService.StoreInfo_DBUpdate(id,
                                                   p_oForm["SID"].Trim(),
                                                   p_oForm["Name"].Trim(),
                                                   Convert.ToInt32(p_oForm["Kind"]),
                                                   p_oForm["Memo"].Trim(),
                                                   Convert.ToInt32(p_oForm["Group_ID"]),
                                                   p_oForm["BranchArea_ID"].Trim(),
                                                   p_oForm["OrderArea_ID"].Trim(),                                                  
                                                   p_oForm["BeginDate"].Trim(),
                                                   p_oForm["TelPhone"].Trim(),
                                                   p_oForm["FaxPhone"].Trim(),
                                                   p_oForm["CellPhone"].Trim(),
                                                   p_oForm["Contact"].Trim(),
                                                   p_oForm["Address"].Trim(),                                                   
                                                   p_oForm["MobileStoreID"].Trim(),
                                                   p_oForm["ManageKind"].Trim()
                                                   );

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                string l_sErrMsg = ex.Message;
                return View();
            }
        }


        // POST: MERP_UAJ000/MERP_UAJ001/GroupSet_Create/5
        [HttpPost]
        public ActionResult GroupSet_Create(FormCollection p_oForm)
        {
            string l_sSelGroup = p_oForm["GroupSet"]; //分類ID
            int l_iSelGroupNo = Convert.ToInt32(l_sSelGroup); //選擇分類ID

            //20210225 CCL+ 區域TypeID
            string l_sSelGroupType = p_oForm["GroupSetType"]; //區域Type ID


            //string l_sSelUpdSIDs = p_oForm[""].Trim();

            //20210225 CCL+ 區域TypeID
            string l_sSelGroupSetTypeName = p_oForm["SelGroupSetTypeName"].Trim(); //區域名稱              
            string l_sSelGroupSetName = p_oForm["SelGroupSetName"].Trim(); //分類名稱
            string l_sSelGroupSetFullName = l_sSelGroupSetTypeName + l_sSelGroupSetName;
            string l_sSelGroupSetFullDesc = l_sSelGroupSetTypeName + "," + l_sSelGroupSetName;

            string l_sSelShopIDSList = p_oForm["CheckedShopItems"].Trim(); //店列表Str
            string l_sInputShopIDSList = p_oForm["ShopsGroupSIDS"].Trim();

            try
            {
                //改存在一個分類Grouping Table
                //m_StoreDBService.StoreInfo_SqlDBUpdateGroup_IDBySID(l_sSelUpdSIDs, l_iSelGroupVal);
                //如果有存在的話改成更新,否則新增設定
                //20210225 CCL- StoreGroupSet l_oTmpItem = m_StoreGroupSetDBService.StoreGroupSet_GetDataByGroupNo(l_iSelGroupNo);
                //20210225 CCL+ 區域TypeID
                StoreGroupSet l_oTmpItem = m_StoreGroupSetDBService.StoreGroupSet_GetDataByGroupNoType(l_iSelGroupNo, l_sSelGroupType);
                if (l_oTmpItem != null)
                {
                    //更新
                    Trace.WriteLine(l_oTmpItem.StoreGroupNo);
                } else
                {
                    //新增
                    if(p_oForm["IsUseSIDSIInput"] != null &&  p_oForm["IsUseSIDSIInput"].Contains("on"))
                    {
                        //m_StoreGroupSetDBService.StoreGroupSet_DBCreate(l_iSelGroupNo,
                        //                        "", l_sSelGroupSetName, "", l_sInputShopIDSList);

                        //20210225 CCL+ 區域TypeID
                        m_StoreGroupSetDBService.StoreGroupSet_DBCreate(l_iSelGroupNo,
                                               l_sSelGroupType.Trim(), l_sSelGroupSetFullName,
                                               l_sSelGroupSetFullDesc, l_sInputShopIDSList);
                    } else
                    {
                        //m_StoreGroupSetDBService.StoreGroupSet_DBCreate(l_iSelGroupNo,
                        //                        "", l_sSelGroupSetName, "", l_sSelShopIDSList);

                        //20210225 CCL+ 區域TypeID
                        m_StoreGroupSetDBService.StoreGroupSet_DBCreate(l_iSelGroupNo,
                                               l_sSelGroupType.Trim(), l_sSelGroupSetFullName,
                                               l_sSelGroupSetFullDesc, l_sSelShopIDSList);

                    }
                    
                    
                }
                //

            }
            catch (Exception ex)
            {
                string l_sErrMsg = ex.Message;
                return Content("<script>confirm('錯誤訊息: " + l_sErrMsg + "');window.close()</script>");
            }


            //return View();
            return RedirectToAction("GroupSet_Create");
        }

        //設定直營,合營,加盟 Group_ID = 0:一般(未分類) 1:直營, 2:合營, 3:加盟
        // POST: MERP_UAJ000/MERP_UAJ001/GroupSet_Edit
        [HttpPost]
        public ActionResult GroupSet_Edit(int id, FormCollection p_oForm)        
        {

            try
            {


                //string l_sShopIDSStr = p_oForm["CheckedShopItems"].Trim();
                string l_sSelShopIDSList = p_oForm["CheckedShopItems"].Trim(); //店列表Str
                string l_sInputShopIDSList = p_oForm["ShopsGroupSIDS"].Trim();

                StoreGroupSet l_oUpdItem = new StoreGroupSet();

                if (p_oForm["IsUseSIDSIInput"] != null && p_oForm["IsUseSIDSIInput"].Contains("on"))
                {
                    l_oUpdItem.StoreGroupSIDList = l_sInputShopIDSList;
                } else
                {
                    l_oUpdItem.StoreGroupSIDList = l_sSelShopIDSList;
                }
               
                l_oUpdItem.StoreGroupNo = Convert.ToInt32(p_oForm["SelGroupSetNo"]);
                //l_oUpdItem.StoreGroupType = "";
                l_oUpdItem.StoreGroupType = p_oForm["SelGroupSetType"]; //20210225 CCL+
                //l_oUpdItem.StoreGroupName = p_oForm["SelGroupSetName"];
                //l_oUpdItem.StoreGroupDesc = "";
                //l_oUpdItem.StoreGroupSIDList = l_sShopIDSStr;

                m_StoreGroupSetDBService.StoreGroupSet_DBUpdate(id, l_oUpdItem);
            }
            catch (Exception ex)
            {

                string l_sErrMsg = ex.Message;
                return Content("<script>confirm('錯誤訊息: " + l_sErrMsg + "');window.close()</script>");

            }

            return RedirectToAction("GroupSet_Index");
        }


        // POST: MERP_UAJ000/MERP_UAJ001/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                string l_sErrMsg = ex.Message;
                return View();
            }
        }

        #endregion
    }
}
