using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Services;
using MACU_WEB.Models;
using System.IO;
using MACU_WEB.BIServices;


namespace MACU_WEB.Areas.MERP_TDQ000.Controllers
{
    //上傳多部門比較損益表Excel,並做各式處理
    public class MERP_TDQ001Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TDQ001"; //客製程式
        string strMENU_ID = "MERP_TDQ000";

        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        #endregion


        #region Action_View
        // GET: MERP_TDQ000/MERP_TDQ001
        public ActionResult Index()
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);

            return View(l_oDataList);
        }

        // GET: MERP_TDQ000/MERP_TDQ001/Details/5
        public ActionResult Details()
        {
            return View();
        }

        // GET: MERP_TDQ000/MERP_TDQ001/Create
        public ActionResult Create()
        {
            return View();
        }

       
        // Get: MERP_TDQ000/MERP_TDQ001/Delete
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

        #endregion





        #region Action_DB
        [HttpPost]
        #region 查詢畫面送出(Index) [Submit]
        public ActionResult Index(HttpPostedFileBase upload)
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            //執行上傳檔案
            //UploadFile(upload);

            MERP_UploadBIService.UploadFile(upload, Server, strPROG_ID);

            //只查出
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Up", strPROG_ID);

            
            //return RedirectToAction("UploadFile", "Upload", new { area = "FrameWork" });
            //return RedirectToRoute("FrameWork/Upload/UploadFile");

            //Redirect("FrameWork/UploadController/UploadFile");
            //RedirectToRoute("FrameWork/UploadController/UploadFile");
            //string l_sUrl = Url.Action("UploadFile", "Upload", new { area = "FrameWork" });
            //return Redirect(l_sUrl);
            return View(l_oDataList);
        }
        #endregion

        // POST: MERP_TDQ000/MERP_TDQ001/Create
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

        // GET: MERP_TDQ000/MERP_TDQ001/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MERP_TDQ000/MERP_TDQ001/Edit/5
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



        // POST: MERP_TDQ000/MERP_TDQ001/Delete/5
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

        #region Private Method

        /*
        public ActionResult UploadFile(HttpPostedFileBase p_oUpload)
        {

            // var result = new Result<string>();

            //var l_oFilessss = HttpContext.Request.Files;

            try
            {

                HttpPostedFileBase l_oFiles = p_oUpload;
                //HttpPostedFileBase l_oFiles = Request.Files["file"];
                FileInfo l_oFileInfo = new FileInfo(l_oFiles.FileName);//獲取文件名稱
                string type = l_oFileInfo.Extension.ToLower();//獲取副檔名
                string name = Guid.NewGuid().ToString();//創建文件名稱
                //string filepath = Server.MapPath("~/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd"));
                //ViewData["PROG_ID"]
                //ViewData["MENU_ID"]
                string l_sFilePath = Server.MapPath("~/Up_Data/" + ViewData["PROG_ID"] + "/" + System.DateTime.Now.ToString("yyyyMMdd"));

                //文件夹不存在创建文件夹
                if (!Directory.Exists(l_sFilePath))
                {
                    Directory.CreateDirectory(l_sFilePath);
                }


                //l_oFiles.SaveAs(l_sFilePath + "/" + name + type);//保存文件

                string l_sUrl = l_sFilePath + "/" + name + type;
                l_oFiles.SaveAs(l_sUrl);//保存文件
                //string l_sUrl = HttpContext.Request.Url.Host;
                int l_iPort = HttpContext.Request.Url.Port;

                //利用DBService將檔案資訊存入資料庫
                m_FileDBService.FileContent_DBCreate(l_oFiles.FileName, l_sUrl, l_oFiles.ContentLength, l_oFiles.ContentType);

                if (l_iPort == 80)
                {
                    //result.msg = "/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd") + "/" + name + type;//返回文件路径
                }
                else
                {
                    //result.msg = "/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd") + "/" + name + type;//返回文件路径
                }
                //result.flag = true;
            }
            catch (Exception ex)
            {
                //Log(ex);
                //result.msg = ex.Message;
            }
            //return Json(result, JsonRequestBehavior.AllowGet); 
            //return Json("", JsonRequestBehavior.AllowGet); 
            return RedirectToAction("Index");
        }
        */

    }
    #endregion
}

