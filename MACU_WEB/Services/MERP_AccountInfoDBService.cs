using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data.SqlClient;

namespace MACU_WEB.Services
{
    public class MERP_AccountInfoDBService
    {
        //操作DB
        public MACU_WEB.Models.MERPEntities db = new MERPEntities();
        public int l_iDBStatus = 0;

        #region AccountInfo
        //1.
        public List<AccountInfo> AccountInfo_GetDataList()
        {
            //IQueryable<FileContent> l_oSearchData = GetAllDataList();
            List<AccountInfo> l_oRtnlist = db.AccountInfo.Where(m => m.IsValid == 1).ToList();
            return l_oRtnlist;
            //return db.AccountInfo.ToList();
        }

        //2.
        public AccountInfo AccountInfo_GetDataById(int p_iId)
        {
           
            AccountInfo l_oFindItem = db.AccountInfo.Find(p_iId);
            return l_oFindItem;
        }

        //
        public AccountInfo AccountInfo_GetDataByAccNoDtlAccNo(string p_sAccNo, string p_sDtlAccNo)
        {
            //找出科目名稱 AND 明細科目名稱 都相同的 AccountInfo
            string l_sAccNo = p_sAccNo.Trim();
            string l_sDtlAccNo = p_sDtlAccNo.Trim();
            //只找回找到的            
            if (db.AccountInfo.Where(m => (m.AccountNo == l_sAccNo) &&
                                          (m.DetailAccNo == l_sDtlAccNo) &&
                                          (m.IsValid == 1)).Count() > 0)
            {
                AccountInfo l_oFindItem = db.AccountInfo.Where(m => (m.AccountNo == l_sAccNo) &&
                                                                (m.DetailAccNo == l_sDtlAccNo) &&
                                                                (m.IsValid == 1)).First();

                return l_oFindItem;
            } else
            {
                //throw new Exception("Find Nothing");
                return null;
            }


            
        }

        //20201230 CCL+ //加入多用名稱查找 ////////////////////////////////////////////////////////////
        public AccountInfo AccountInfo_GetDataByAccNoDtlAccNo(string p_sAccNo, string p_sDtlAccNo, string p_sAccName)
        {
            //找出科目名稱 AND 明細科目名稱 都相同的 AccountInfo
            string l_sAccNo = p_sAccNo.Trim();
            string l_sDtlAccNo = p_sDtlAccNo.Trim();
            string l_sAccName = p_sAccName.Trim();
            //只找回找到的            
            if (db.AccountInfo.Where(m => (m.AccountNo == l_sAccNo) &&
                                          (m.DetailAccNo == l_sDtlAccNo) &&
                                          (m.AccountName == l_sAccName) && 
                                          (m.IsValid == 1)).Count() > 0)
            {
                AccountInfo l_oFindItem = db.AccountInfo.Where(m => (m.AccountNo == l_sAccNo) &&
                                                                (m.DetailAccNo == l_sDtlAccNo) &&
                                                                (m.AccountName == l_sAccName) && 
                                                                (m.IsValid == 1)).First();

                return l_oFindItem;
            }
            else
            {
                //throw new Exception("Find Nothing");
                return null;
            }



        }
        /// ///////////////////////////////////////////////////////////////////////////////////////////


        //3.
        public void AccountInfo_DBCreate(string p_sAccNo, string p_sAccName, string p_sDetailAccNo, string p_sDetailAccName, 
                                         string p_sCountFlag, int p_iPrintOrder, int p_iGroupID)
        {
            AccountInfo l_oNewFile = new AccountInfo();

            l_oNewFile.AccountNo = p_sAccNo;
            l_oNewFile.AccountName = p_sAccName;
            l_oNewFile.DetailAccNo = p_sDetailAccNo;
            l_oNewFile.DetailAccName = p_sDetailAccName;
            l_oNewFile.CountFlag = p_sCountFlag; //C:貸, D:借, S:加總
            l_oNewFile.PrintOrder = p_iPrintOrder; //Excel列印顯示順序
            l_oNewFile.GroupID = p_iGroupID; //群組ID
            l_oNewFile.IsValid = 1;
            l_oNewFile.CreateTime = DateTime.Now;
            l_oNewFile.UpdateTime = DateTime.Now;


            db.AccountInfo.Add(l_oNewFile);

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

        //4.
        public void AccountInfo_DBDeleteByID(int p_iItemID)
        {
            AccountInfo l_oDelItem = db.AccountInfo.Find(p_iItemID);
            db.AccountInfo.Remove(l_oDelItem);
            db.SaveChanges();
        }

        //5.
        public void AccountInfo_DBUpdate(int p_iItemID, AccountInfo p_oNewUpdItem)
        {
            AccountInfo l_oUpdItem = db.AccountInfo.Find(p_iItemID);
            
            l_oUpdItem.AccountNo = p_oNewUpdItem.AccountNo;
            l_oUpdItem.AccountName = p_oNewUpdItem.AccountName;
            l_oUpdItem.DetailAccNo = p_oNewUpdItem.DetailAccNo;
            l_oUpdItem.DetailAccName = p_oNewUpdItem.DetailAccName;
            l_oUpdItem.CountFlag = p_oNewUpdItem.CountFlag;
            l_oUpdItem.GroupID = p_oNewUpdItem.GroupID;
            l_oUpdItem.PrintOrder = p_oNewUpdItem.PrintOrder;
            l_oUpdItem.UpdateTime = DateTime.Now;

            

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

        #endregion
    }
}