using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data.SqlClient;

namespace MACU_WEB.Services
{
    public class MERP_FileContentDBService
    {
        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;
       

        #region FileContent 
        public List<FileContent> FileContent_GetDataList()
        {
            //IQueryable<FileContent> l_oSearchData = GetAllDataList();

            return db.FileContent.ToList();
        }

        //public IQueryable<FileContent> GetAllDataList()
        //{

        //}
        public List<FileContent> FileContent_GetDataListByDirType_ProgCat(string p_sDirType, string p_sProgCat)
        {
            //IQueryable<FileContent> l_oSearchData = GetAllDataList();

            //根據上下載Type,和程式分類目錄找出File List
            return db.FileContent.Where(m => (m.DirType == p_sDirType) && (m.ProgCatg == p_sProgCat) ).ToList();

            /*
            return db.FileContent.ToList().Select(m => new FileContent() 
            {
                Id = m.Id,
                Name = m.Name,
                Type = m.Type,
                Url = m.Url,
                Size = m.Size,
                CreateTime = m.CreateTime,
                IsValid = m.IsValid,
                DirType = p_sDirType
            }).ToList();
            */
        }


        public FileContent FileContent_GetDataById(int p_iId)
        {
            //FileContent l_oFindFile = new FileContent();
            //l_oFindFile.Id = p_iId;
            //return db.FileContent.Find(l_oFindFile);
            FileContent l_oFindFile = db.FileContent.Find(p_iId);
            return l_oFindFile;
        }

        public void FileContent_DBCreate(string p_sName, string p_sUrl, int p_sSize, string p_sType, string p_sDir, string p_sProgCatg)
        {
            FileContent l_oNewFile = new FileContent();

            l_oNewFile.Name = p_sName;
            l_oNewFile.Url = p_sUrl;
            l_oNewFile.Size = p_sSize;
            l_oNewFile.Type = p_sType;
            l_oNewFile.DirType = p_sDir;
            l_oNewFile.ProgCatg = p_sProgCatg;
            l_oNewFile.IsValid = 1;
            l_oNewFile.CreateTime = DateTime.Now;
            l_oNewFile.UpdateTime = DateTime.Now;



            db.FileContent.Add(l_oNewFile);

            //Log l_oLog = new Log();
            //l_oLog.LogCount = l_oLog.LogCount + 1;
            //db.Log.Add(l_oLog);
            try
            {
                l_iDBStatus = db.SaveChanges();
            } catch(Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }
            

            
        }

        public void FileContent_DBDeleteByID(int p_iFileID)
        {
            FileContent l_oDelFile = db.FileContent.Find(p_iFileID);
            db.FileContent.Remove(l_oDelFile);
            db.SaveChanges();
        }


        //20210111 CCL+ 勞健保資料內無欄位可判斷年月,只能由上傳時手動給入 ///////////////////////////////////////
        public void FileContent_DBCreate(string p_sName, string p_sUrl, int p_sSize, string p_sType, 
                                         string p_sDir, string p_sProgCatg, 
                                         string p_sYear, string p_sMonth)
        {
            FileContent l_oNewFile = new FileContent();

            l_oNewFile.Name = p_sName;
            l_oNewFile.Url = p_sUrl;
            l_oNewFile.Size = p_sSize;
            l_oNewFile.Type = p_sType;
            l_oNewFile.DirType = p_sDir;
            l_oNewFile.ProgCatg = p_sProgCatg;
            l_oNewFile.DataYear = p_sYear.Trim();
            l_oNewFile.DataMonth = p_sMonth.Trim();
            l_oNewFile.IsValid = 1;
            l_oNewFile.CreateTime = DateTime.Now;
            l_oNewFile.UpdateTime = DateTime.Now;


            db.FileContent.Add(l_oNewFile);

            //Log l_oLog = new Log();
            //l_oLog.LogCount = l_oLog.LogCount + 1;
            //db.Log.Add(l_oLog);
            try
            {
                l_iDBStatus = db.SaveChanges();
            }
            catch (Exception ex)
            {
                int l_iError = l_iDBStatus;
                string errmsg = ex.Message.ToString();
            }



        }



        /////////////////////////////////////////////////////////////////////////////////////////////////////////

        #endregion

    }
}