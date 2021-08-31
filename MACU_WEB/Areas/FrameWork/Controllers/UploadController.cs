using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Services;

namespace MACU_WEB.Areas.FrameWork.Controllers
{
    public class UploadController : Controller
    {
        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();

        // GET: FrameWork/Upload
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        //public JsonResult UploadFile()
        public ActionResult UploadFile()
        {
           
           // var result = new Result<string>();

            //var l_oFilessss = HttpContext.Request.Files;

            try
            {
                

                HttpPostedFileBase l_oFiles = Request.Files["file"];
                FileInfo l_oFileInfo = new FileInfo(l_oFiles.FileName);//獲取文件名稱
                string type = l_oFileInfo.Extension.ToLower();//獲取副檔名
                string name = Guid.NewGuid().ToString();//創建文件名稱
                //string filepath = Server.MapPath("~/Image/Upload/" + System.DateTime.Now.ToString("yyyyMMdd"));
                //ViewData["PROG_ID"]
                //ViewData["MENU_ID"]
                string l_sFilePath = Server.MapPath("~/Up_Data/" + ViewData["PROG_ID"] + "/" +  System.DateTime.Now.ToString("yyyyMMdd"));

                //文件夹不存在创建文件夹
                if (!Directory.Exists(l_sFilePath))
                {
                    Directory.CreateDirectory(l_sFilePath);
                }

                l_oFiles.SaveAs(l_sFilePath + "/" + name + type);//保存文件

                string l_sUrl = HttpContext.Request.Url.Host;
                int l_iPort = HttpContext.Request.Url.Port;

                //利用DBService將檔案資訊存入資料庫
                m_FileDBService.FileContent_DBCreate(l_oFiles.FileName, l_sUrl, l_oFiles.ContentLength, l_oFiles.ContentType, "Up", "MERP_TDQ001");

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
            //return Json(result, JsonRequestBehavior.AllowGet); ;
            //return Json("", JsonRequestBehavior.AllowGet); 
            return RedirectToAction("Index");
        }
    }
}