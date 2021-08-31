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

    //部門勞保負擔比例設定Settings
    public class MERP_TCF002Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TCF002"; //客製程式
        string strMENU_ID = "MERP_TCF000";

        const string LISETNO_PREFIX = "LISET";
       
        public MERP_FA_LaborInsSetDBService m_LInsSettingDBService = new MERP_FA_LaborInsSetDBService();
        public MERP_FA_LaborHealthInsV1DBService m_LHInsDBV1Service = new MERP_FA_LaborHealthInsV1DBService();
        public MERP_FA_LaborSubsFundSetDBService m_LSubsFundSetDBService = new MERP_FA_LaborSubsFundSetDBService();
        public MERP_FA_LaborInsSetMapComSetDBService m_LInsSetMapComSetDBService =
                                                                        new MERP_FA_LaborInsSetMapComSetDBService();

        #endregion

        #region Action_View
        // GET: MERP_TCF000/MERP_TCF002
        public ActionResult Index()
        {
            List<FA_LaborInsSet> l_oRtnData = m_LInsSettingDBService.FA_LaborInsSet_GetDataList();

            return View(l_oRtnData);
            //return View();
        }

        // GET: MERP_TCF000/MERP_TCF002/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MERP_TCF000/MERP_TCF002/Create
        public ActionResult Create()
        {
            //找出目前資料庫中,的LInsSetNo 流水號 + 1
            String l_sLaborInsSetNo = "";
            List<FA_LaborInsSet> l_oRtnData = m_LInsSettingDBService.FA_LaborInsSet_GetDataList();
            if(l_oRtnData != null )
            {

                if(l_oRtnData.Count() > 0 )
                {
                    //抓出最大流水號
                    FA_LaborInsSet l_oTmpData = l_oRtnData.OrderByDescending(m => m.LabInsSetNo).First();
                    string l_sTheMaxNo = l_oTmpData.LabInsSetNo;
                    string l_sNo = l_sTheMaxNo.Substring(LISETNO_PREFIX.Length+1);
                    double l_dNewNoNum = Convert.ToDouble(l_sNo) + 1;
                    //l_sLaborInsSetNo = LISETNO_PREFIX + l_dNewNoNum.ToString();
                    l_sLaborInsSetNo = l_dNewNoNum.ToString();
                    for (int i=0; i< 5; i++)
                    {
                        if(l_sLaborInsSetNo.ToString().Length < 5)
                        {
                            l_sLaborInsSetNo = "0" + l_sLaborInsSetNo;
                        } else
                        {
                            l_sLaborInsSetNo = LISETNO_PREFIX + l_sLaborInsSetNo;
                            break;
                        }
                    }

                    ViewData["LabInsSetNo"] = l_sLaborInsSetNo;

                } else if(l_oRtnData.Count() == 0)
                {
                    //DB是空的 流水號 = "00001"
                    l_sLaborInsSetNo = LISETNO_PREFIX + "00001";
                    ViewData["LabInsSetNo"] = l_sLaborInsSetNo;
                }
            }

            return View();
            //return View(l_sLaborInsSetNo);
        }

        /*
        // GET: MERP_TCF000/MERP_TCF002/Delete/5
        public ActionResult Delete(int id)
        {
            m_LInsSettingDBService.FA_LaborInsSet_DBDeleteByID(id);

            return RedirectToAction("Index");
        }
        */

        
        // GET: MERP_TCF000/MERP_TCF002/Delete/5
        public ActionResult Delete(string no)
        {
            m_LInsSettingDBService.FA_LaborInsSet_DBDeleteByNo(no);

            return RedirectToAction("Index");
        }
        
        
        /*
        // GET: MERP_TCF000/MERP_TCF002/Edit/5
        public ActionResult Edit(int id)
        {
            FA_LaborInsSet l_oRtnData = m_LInsSettingDBService.FA_LaborInsSet_GetDataById(id);
            //把日期0回復和改成"-"分隔才能正常顯示
            string l_sOnBeginDate = l_oRtnData.OnBeginDate;
            l_sOnBeginDate = DateStringProcess.ReStore_MonthDayZero(l_sOnBeginDate, "/", "-");
            l_oRtnData.OnBeginDate = l_sOnBeginDate;

            ViewData["FA_LInsSet"] = l_oRtnData ;
            //return View(l_oRtnData);
            return View();
        }
        */

        
        //20210126 CCL+
        public ActionResult Edit(string no)
        {
            FA_LaborInsSet l_oRtnData = m_LInsSettingDBService.FA_LaborInsSet_GetDataByNo(no);
            //把日期0回復和改成"-"分隔才能正常顯示
            string l_sOnBeginDate = l_oRtnData.OnBeginDate;
            l_sOnBeginDate = DateStringProcess.ReStore_MonthDayZero(l_sOnBeginDate, "/", "-");
            l_oRtnData.OnBeginDate = l_sOnBeginDate;

            ViewData["FA_LInsSet"] = l_oRtnData;
            //return View(l_oRtnData);
            return View();
        }
        

        // GET: MERP_TCF000/MERP_TCF002/LaborSubsFund_Create
        public ActionResult LaborSubsFund_Create()
        {
            //抓出所有公司名稱
            //IEnumerable<FA_LaborHealthInsV1> l_oLHInsV1 = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataByYearMon(2021, 1);
            IEnumerable<FA_LaborHealthInsV1> l_oLHInsV1 = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataList();

            
            var l_oTmpData = from DataVals in l_oLHInsV1
                             group DataVals by DataVals.PlusInsCompany into grp
                             select new {
                                 PlusInsCom = grp.Key,
                                 PlusInsComCount = grp.Count()
                             };

            var l_oItemDatas = from DataVals in l_oTmpData
                              select new SelectListItem
                             {
                                 Text = DataVals.PlusInsCom,
                                 Value = DataVals.PlusInsCom
                             };
                      

            l_oItemDatas.Where(m => m.Text == "集朗有限公司").First().Selected = true;

            SelectList l_oRtnSelListData = new SelectList(l_oItemDatas, "Value", "Text");
            //List<SelectListItem> l_oRtnDate = l_oItemData.ToList();

            //抓出勞保代墊基金所有設定
            List<FA_LaborSubsFundSet> l_oExistedLSFundSets = m_LSubsFundSetDBService.FA_LaborSubsFundSet_GetDataList();

            MERP_TCF002_LaborSubsFund_CreateViewModel l_oRtnDataVM = new MERP_TCF002_LaborSubsFund_CreateViewModel();
            l_oRtnDataVM.m_oPlusComInsList = l_oRtnSelListData;
            l_oRtnDataVM.m_oExistedLSFundSetList = l_oExistedLSFundSets;

            return View(l_oRtnDataVM);
        }


        // GET: MERP_TCF000/MERP_TCF002/LabInsSetMapCom_Create
        public ActionResult LabInsSetMapCom_Create()
        {
            //抓出所有公司名稱
            //IEnumerable<FA_LaborHealthInsV1> l_oLHInsV1 = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataByYearMon(2021, 1);
            IEnumerable<FA_LaborHealthInsV1> l_oLHInsV1 = m_LHInsDBV1Service.FA_LaborHealthInsV1_GetDataList();


            var l_oTmpData = from DataVals in l_oLHInsV1
                             group DataVals by DataVals.PlusInsCompany into grp
                             select new
                             {
                                 PlusInsCom = grp.Key,
                                 PlusInsComCount = grp.Count()
                             };

            var l_oItemDatas = from DataVals in l_oTmpData
                               select new SelectListItem
                               {
                                   Text = DataVals.PlusInsCom,
                                   Value = DataVals.PlusInsCom
                               };


            l_oItemDatas.Where(m => m.Text == "集朗有限公司").First().Selected = true;

            SelectList l_oRtnSelListData = new SelectList(l_oItemDatas, "Value", "Text");
            //List<SelectListItem> l_oRtnDate = l_oItemData.ToList();

            //取得所有勞保設定
            List<FA_LaborInsSet> l_oLInsSetList = m_LInsSettingDBService.FA_LaborInsSet_GetDataList();

            var l_oLInsSetSelList = from DataVals in l_oLInsSetList
                                    select new SelectListItem
                                    {
                                        Text = DataVals.LabInsSetNo,
                                        Value = DataVals.LabInsSetNo
                                    };

            SelectList l_oLInsSetSelListData = new SelectList(l_oLInsSetSelList, "Value", "Text");

            //取得所有LaborInsSet Map PlusCompany
            List<FA_LaborInsSetMapComSet> l_oExistedLabInsSetMapComSetList =
                                        m_LInsSetMapComSetDBService.FA_LaborInsSetMapComSet_GetDataList();


            MERP_TCF002_LabInsSetMapCom_CreateViewModel l_oRtnDataVM = new MERP_TCF002_LabInsSetMapCom_CreateViewModel();
            l_oRtnDataVM.m_oPlusComInsList = l_oRtnSelListData;
            l_oRtnDataVM.m_oLaborInsSettings = l_oLInsSetList;
            l_oRtnDataVM.m_oLaborInsSetList = l_oLInsSetSelListData;
            l_oRtnDataVM.m_oExistedLabInsMapPlusComSetList = l_oExistedLabInsSetMapComSetList;

            //ViewData["LabInsSetMapCom"] = l_oRtnDataVM;

            //return View();
            return View(l_oRtnDataVM);
        }

        // GET: MERP_TCF000/MERP_TCF002/LabInsSetMapCom_Edit/1
        public ActionResult LabInsSetMapCom_Edit(int id)
        {
            //編輯更新對應設定
            //取得要編輯的LaborInsSet Map PlusCompany
            FA_LaborInsSetMapComSet l_oLInsSetMapData = 
                m_LInsSetMapComSetDBService.FA_LaborInsSetMapComSet_GetDataById(id);

            //抓出目前選擇的勞保設定
            string l_sLaborInsMapComSetNo = l_oLInsSetMapData.LabInsSetNo;
            string l_sToEditPlusCompany = l_oLInsSetMapData.PlusInsCompany;

            //取得所有勞保設定
            List<FA_LaborInsSet> l_oLInsSetList = m_LInsSettingDBService.FA_LaborInsSet_GetDataList();

            var l_oLInsSetSelList = from DataVals in l_oLInsSetList
                                    select new SelectListItem
                                    {
                                        Text = DataVals.LabInsSetNo,
                                        Value = DataVals.LabInsSetNo
                                    };
                                                                        

            //設定selected
            //l_oLInsSetSelList.Where(m => m.Value == l_sLaborInsMapComSetNo).First().Selected = true;
            

            SelectList l_oLInsSetSelListData = new SelectList(l_oLInsSetSelList, "Value", "Text");

            MERP_TCF002_LabInsSetMapCom_EditViewModel l_oRtnDataVM = new MERP_TCF002_LabInsSetMapCom_EditViewModel();
            l_oRtnDataVM.m_sPlusCompany = l_sToEditPlusCompany;
            l_oRtnDataVM.m_sOrgLaborInsSetNo = l_sLaborInsMapComSetNo;
            l_oRtnDataVM.m_oLaborInsSetList = l_oLInsSetSelListData;
            l_oRtnDataVM.m_oLaborInsSettings = l_oLInsSetList;



            return View(l_oRtnDataVM);
        }

        // GET: MERP_TCF000/MERP_TCF002/LabInsSetMapCom_Create
        public ActionResult LabInsSetMapCom_Delete(int id)
        {

            try
            {
                m_LInsSetMapComSetDBService.FA_LaborInsSetMapComSet_DBDeleteByID(id);
            }
            catch (Exception ex)
            {
                string l_sErrMsg = ex.Message;
                return Content("<script>confirm('錯誤訊息: " + l_sErrMsg + "');window.close()</script>");
            }

            return RedirectToAction("LabInsSetMapCom_Create");
        }

        #endregion


        #region Action_DB
        // POST: MERP_TCF000/MERP_TCF002/Create
        [HttpPost]
        public ActionResult Create(FormCollection p_oForm)
        {
            try
            {
                string l_sOnBeginDate = "";
                // TODO: Add insert logic here
                if (!string.IsNullOrEmpty(p_oForm["OnBeginDate"]))
                {
                    l_sOnBeginDate = DateStringProcess.Del_MonthDayZero(p_oForm["OnBeginDate"], "-", "/");
                }


                m_LInsSettingDBService.FA_LaborInsSet_DBCreate(
                                                    p_oForm["LabInsSetNo"],
                                                    Convert.ToDecimal(p_oForm["OrdAcciInsRate"]),
                                                    Convert.ToDecimal(p_oForm["EmployInsRate"]),
                                                    Convert.ToDecimal(p_oForm["PersonalInsRate"]),
                                                    Convert.ToDecimal(p_oForm["LaborBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["ComBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["GovBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["CommuteDisaInsRate"]),
                                                    Convert.ToDecimal(p_oForm["IndustryDisaInsRate"]),
                                                    Convert.ToDecimal(p_oForm["OccuDisaInsRate"]),
                                                    Convert.ToDecimal(p_oForm["OccuDisComBurdenRatio"]),
                                                    Convert.ToDecimal(p_oForm["LaborSubsFund"]),
                                                    Convert.ToDecimal(p_oForm["LaborSubsFundRate"]),
                                                    Convert.ToDecimal(p_oForm["LaborRetireRate"]),
                                                    l_sOnBeginDate
                                                    );
                
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        /*
        // POST: MERP_TCF000/MERP_TCF002/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection p_oForm)
        {
            try
            {
                // TODO: Add update logic here
                string l_sOnBeginDate = "";
                // TODO: Add insert logic here
                if (!string.IsNullOrEmpty(p_oForm["OnBeginDate"]))
                {
                    l_sOnBeginDate = DateStringProcess.Del_MonthDayZero(p_oForm["OnBeginDate"], "-", "/");
                }

                FA_LaborInsSet l_oUpdItem = new FA_LaborInsSet();

                l_oUpdItem.LabInsSetNo = p_oForm["LabInsSetNo"];
                l_oUpdItem.OrdAccidentInsRate = Convert.ToDecimal(p_oForm["OrdAcciInsRate"]);
                l_oUpdItem.EmployInsRate = Convert.ToDecimal(p_oForm["EmployInsRate"]);
                l_oUpdItem.PersonalInsRate = Convert.ToDecimal(p_oForm["PersonalInsRate"]);
                l_oUpdItem.LaborBurdenRatio = Convert.ToDecimal(p_oForm["LaborBurdenRatio"]);
                l_oUpdItem.ComBurdenRatio = Convert.ToDecimal(p_oForm["ComBurdenRatio"]);
                l_oUpdItem.GovBurdenRatio = Convert.ToDecimal(p_oForm["GovBurdenRatio"]);
                l_oUpdItem.CommuteDisaInsRate = Convert.ToDecimal(p_oForm["CommuteDisaInsRate"]);
                l_oUpdItem.IndustryDisaInsRate = Convert.ToDecimal(p_oForm["IndustryDisaInsRate"]);
                l_oUpdItem.OccuDisaInsRate = Convert.ToDecimal(p_oForm["OccuDisaInsRate"]);
                l_oUpdItem.OccuDisComBurdenRatio = Convert.ToDecimal(p_oForm["OccuDisComBurdenRatio"]);
                l_oUpdItem.LaborSubsFund = Convert.ToDecimal(p_oForm["LaborSubsFund"]);
                l_oUpdItem.LaborSubsFundRate = Convert.ToDecimal(p_oForm["LaborSubsFundRate"]);
                l_oUpdItem.LaborRetireRate = Convert.ToDecimal(p_oForm["LaborRetireRate"]);
                l_oUpdItem.OnBeginDate = l_sOnBeginDate;


                m_LInsSettingDBService.FA_LaborInsSet_DBUpdate(id, l_oUpdItem);

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        */

        // POST: MERP_TCF000/MERP_TCF002/Edit/5
        [HttpPost]
        public ActionResult Edit(string LabInsSetNo, FormCollection p_oForm)
        {
            try
            {
                // TODO: Add update logic here
                string l_sOnBeginDate = "";
                // TODO: Add insert logic here
                if (!string.IsNullOrEmpty(p_oForm["OnBeginDate"]))
                {
                    l_sOnBeginDate = DateStringProcess.Del_MonthDayZero(p_oForm["OnBeginDate"], "-", "/");
                }

                FA_LaborInsSet l_oUpdItem = new FA_LaborInsSet();

                l_oUpdItem.LabInsSetNo = p_oForm["LabInsSetNo"];
                l_oUpdItem.OrdAccidentInsRate = Convert.ToDecimal(p_oForm["OrdAcciInsRate"]);
                l_oUpdItem.EmployInsRate = Convert.ToDecimal(p_oForm["EmployInsRate"]);
                l_oUpdItem.PersonalInsRate = Convert.ToDecimal(p_oForm["PersonalInsRate"]);
                l_oUpdItem.LaborBurdenRatio = Convert.ToDecimal(p_oForm["LaborBurdenRatio"]);
                l_oUpdItem.ComBurdenRatio = Convert.ToDecimal(p_oForm["ComBurdenRatio"]);
                l_oUpdItem.GovBurdenRatio = Convert.ToDecimal(p_oForm["GovBurdenRatio"]);
                l_oUpdItem.CommuteDisaInsRate = Convert.ToDecimal(p_oForm["CommuteDisaInsRate"]);
                l_oUpdItem.IndustryDisaInsRate = Convert.ToDecimal(p_oForm["IndustryDisaInsRate"]);
                l_oUpdItem.OccuDisaInsRate = Convert.ToDecimal(p_oForm["OccuDisaInsRate"]);
                l_oUpdItem.OccuDisComBurdenRatio = Convert.ToDecimal(p_oForm["OccuDisComBurdenRatio"]);
                l_oUpdItem.LaborSubsFund = Convert.ToDecimal(p_oForm["LaborSubsFund"]);
                l_oUpdItem.LaborSubsFundRate = Convert.ToDecimal(p_oForm["LaborSubsFundRate"]);
                l_oUpdItem.LaborRetireRate = Convert.ToDecimal(p_oForm["LaborRetireRate"]);
                l_oUpdItem.OnBeginDate = l_sOnBeginDate;


                m_LInsSettingDBService.FA_LaborInsSet_DBUpdateByNO(LabInsSetNo, l_oUpdItem);

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        // POST: MERP_TCF000/MERP_TCF002/LaborSubsFund_Create
        [HttpPost]
        public ActionResult LaborSubsFund_Create(string LaborSubsFund, string PlusInsCompany)
        {

            decimal l_dLaborSubsFund = 0;
            if(!string.IsNullOrEmpty(LaborSubsFund))
            {
                l_dLaborSubsFund = Convert.ToDecimal(LaborSubsFund);
            }

            m_LSubsFundSetDBService.FA_LaborSubsFundSet_DBCreate(PlusInsCompany, l_dLaborSubsFund);

            return RedirectToAction("LaborSubsFund_Create");
        }


        // POST: MERP_TCF000/MERP_TCF002/Delete/5
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

        // POST: MERP_TCF000/MERP_TCF002/LabInsSetMapCom_Create
        [HttpPost]
        public ActionResult LabInsSetMapCom_Create(string LaborInsSetNO, string PlusInsCompany)
        {
            try
            {

                m_LInsSetMapComSetDBService.FA_LaborInsSetMapComSet_DBCreate(PlusInsCompany, LaborInsSetNO);

                //m_LInsSetMapComSetDBService
            } catch (Exception ex)
            {

                return View();
            }

            return RedirectToAction("LabInsSetMapCom_Create");
        }

        // POST: MERP_TCF000/MERP_TCF002/LabInsSetMapCom_Edit/1
        [HttpPost]
        public ActionResult LabInsSetMapCom_Edit(int id, string LaborInsSetNO)
        {
            //更新
            try
            {
                                
                m_LInsSetMapComSetDBService.FA_LaborInsSetMapComSet_DBUpdate(id, LaborInsSetNO);
            }
            catch (Exception ex)
            {

                return View();
            }

            return RedirectToAction("LabInsSetMapCom_Create");
        }

        #endregion
    }
}
