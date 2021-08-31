using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Services;
using MACU_WEB.Models;
using MACU_WEB.BIServices;
using System.IO;

namespace MACU_WEB.Areas.MERP_TDQ000.Controllers
{
    //顯示要處理Excel的設定
    public class MERP_TDQ002Controller : Controller
    {
        #region  Param Initial
        string strPROG_ID = "MERP_TDQ002"; //客製程式
        string strMENU_ID = "MERP_TDQ000";

        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        #endregion

        #region Action_View
        // GET: MERP_TDQ000/MERP_TDQ002
        public ActionResult Index(int id)
        {
            ViewData["PROG_ID"] = strPROG_ID;
            ViewData["MENU_ID"] = strMENU_ID;

            FileContent l_oSearchFile = m_FileDBService.FileContent_GetDataById(id);
            //載入上傳目錄內的Excel檔
            MERP_ExcelBIService.ImportExcel(l_oSearchFile.Url);

            return View(l_oSearchFile);
            
        }

        public ActionResult DownFileList()
        {
            List<FileContent> l_oDataList = m_FileDBService.FileContent_GetDataListByDirType_ProgCat("Dn", strPROG_ID);

            return View(l_oDataList);
        }

        // GET: MERP_TDQ000/MERP_TDQ002/Details/5
        public ActionResult Details(int id)
        {


            return View();
        }

        // GET: MERP_TDQ000/MERP_TDQ002/Create
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
        // POST: MERP_TDQ000/MERP_TDQ002/Index
        [HttpPost]
        #region 查詢畫面送出(Index) [Submit]
        public ActionResult Index(FormCollection p_oForm)
        {
            //CheckBox都要使用Contain()來判斷,因為Post來的值是"true,false", MVC會自動在網頁生成一個Hidden Field,
            //所以會有兩個值的字串
            //if (p_oForm["ChkIsDelAllZero"].Contains("true"))
            if (p_oForm["ChkIsDelAllZero"].Contains("on"))
            {
                //如果勾選去除所有欄位Cell皆為0的Row
                MERP_ExcelBIService.Process_DelAllZeroRowsExcel();
                MERP_ExcelBIService.SaveAsExcel(strPROG_ID, Server);
            }
            //return View();
            return RedirectToAction("DownFileList");
        }
        #endregion

        // POST: MERP_TDQ000/MERP_TDQ002/Create
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

        // GET: MERP_TDQ000/MERP_TDQ002/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MERP_TDQ000/MERP_TDQ002/Edit/5
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

      
        // POST: MERP_TDQ000/MERP_TDQ002/Delete/5
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
