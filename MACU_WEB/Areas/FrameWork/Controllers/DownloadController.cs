using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;
using MACU_WEB.Services;
using System.IO;


namespace MACU_WEB.Areas.FrameWork.Controllers
{
    public class DownloadController : Controller
    {
        public MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();

        // GET: FrameWork/Download
        //public ActionResult Index(int p_iId)
        //{
        //    FileContent l_oDownloadFile = m_FileDBService.FileContent_GetDataById(p_iId);

        //    return View(l_oDownloadFile);
        //}


        public ActionResult DownloadFile(int p_iId)
        {
            FileContent l_oDownloadFile = m_FileDBService.FileContent_GetDataById(p_iId);

            if(l_oDownloadFile != null)
            {
                //將檔案讀成串流
                Stream l_oStream = new FileStream(l_oDownloadFile.Url, FileMode.Open, FileAccess.Read, FileShare.Read);
                //回傳出檔案
                return File(l_oStream, l_oDownloadFile.Type, l_oDownloadFile.Name);
            } else
            {
                return JavaScript("alert(\"無此檔案\")");

            }

        }
    }
}