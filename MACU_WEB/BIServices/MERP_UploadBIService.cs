using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using MACU_WEB.Services;

namespace MACU_WEB.BIServices
{
    public static class MERP_UploadBIService
    {
        public static MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();

        public static void UploadFile(HttpPostedFileBase p_oUpload, HttpServerUtilityBase p_oServer, string p_sProgID)
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
                string l_sFilePath = p_oServer.MapPath("~/Up_Data/" + p_sProgID + "/" + System.DateTime.Now.ToString("yyyyMMdd"));

                //文件夹不存在创建文件夹
                if (!Directory.Exists(l_sFilePath))
                {
                    Directory.CreateDirectory(l_sFilePath);
                }


                //l_oFiles.SaveAs(l_sFilePath + "/" + name + type);//保存文件

                //20201223 CCL- string l_sUrl = l_sFilePath + "/" + name + type;
                string l_sUrl = l_sFilePath + "\\" + name + type;

                l_oFiles.SaveAs(l_sUrl);//保存文件
                //string l_sUrl = HttpContext.Request.Url.Host;
                //int l_iPort = HttpContext.Request.Url.Port;

                //利用DBService將檔案資訊存入資料庫
                m_FileDBService.FileContent_DBCreate(l_oFiles.FileName, l_sUrl, l_oFiles.ContentLength, l_oFiles.ContentType, "Up", p_sProgID);

                //if (l_iPort == 80)
                //{
                    //result.msg = "/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd") + "/" + name + type;//返回文件路径
                //}
                //else
                //{
                    //result.msg = "/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd") + "/" + name + type;//返回文件路径
                //}
                //result.flag = true;
            }
            catch (Exception ex)
            {
                //Log(ex);
                //result.msg = ex.Message;
            }
            //return Json(result, JsonRequestBehavior.AllowGet); 
            //return Json(l_sUrl, JsonRequestBehavior.AllowGet); 
            //return RedirectToAction("Index");
        }

        public static bool DeleteFile(string p_sUrl)
        {
            if(p_sUrl != null)
            {
                FileInfo l_oFileInfo = new FileInfo(p_sUrl);
                if(l_oFileInfo.Exists)
                {
                    l_oFileInfo.Delete();
                    return true;
                }
            }

            return false;
            
        }

        //20210111 CCL+ 以年月分區別,因為勞健保上傳資料內容無法分辨年月
        public static void UploadFile(HttpPostedFileBase p_oUpload,
                                    HttpServerUtilityBase p_oServer, string p_sProgID,
                                    string p_sYear, string p_sMonth)
        {
          
            try
            {

                HttpPostedFileBase l_oFiles = p_oUpload;
                //HttpPostedFileBase l_oFiles = Request.Files["file"];
                FileInfo l_oFileInfo = new FileInfo(l_oFiles.FileName);//獲取文件名稱
                string type = l_oFileInfo.Extension.ToLower();//獲取副檔名
                string name = Guid.NewGuid().ToString();//創建文件名稱
                //string year_month = p_sYear.Trim() + p_sMonth.Trim();
                //string filepath = Server.MapPath("~/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd"));
                //ViewData["PROG_ID"]
                //ViewData["MENU_ID"]
                string l_sFilePath = p_oServer.MapPath("~/Up_Data/" + p_sProgID + "/" + System.DateTime.Now.ToString("yyyyMMdd"));

                //文件夹不存在创建文件夹
                if (!Directory.Exists(l_sFilePath))
                {
                    Directory.CreateDirectory(l_sFilePath);
                }
          
                string l_sUrl = l_sFilePath + "\\" + name + type; //保存文件

                l_oFiles.SaveAs(l_sUrl);//保存文件
                //string l_sUrl = HttpContext.Request.Url.Host;
                //int l_iPort = HttpContext.Request.Url.Port;

                //利用DBService將檔案資訊存入資料庫
                //m_FileDBService.FileContent_DBCreate(l_oFiles.FileName, l_sUrl, l_oFiles.ContentLength, l_oFiles.ContentType, "Up", p_sProgID);
                m_FileDBService.FileContent_DBCreate(l_oFiles.FileName, l_sUrl, l_oFiles.ContentLength, 
                                                     l_oFiles.ContentType, "Up", p_sProgID,
                                                     p_sYear.Trim(), p_sMonth.Trim());

            }
            catch (Exception ex)
            {
                //Log(ex);
                //result.msg = ex.Message;
            }           
            //return Json(l_sUrl, JsonRequestBehavior.AllowGet); 
      
        }

    }
}