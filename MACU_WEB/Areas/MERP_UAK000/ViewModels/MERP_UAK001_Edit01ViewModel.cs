using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data;
using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_UAK000.ViewModels
{
    public class MERP_UAK001_Edit01ViewModel
    {
        //20210107 CCL+改分群     
        //北
        public string m_NShopKey { get; set; }
        public int m_NShopCount { get; set; }
        public List<SelectListItem> m_oSelNShopList { get; set; }
        //中
        public string m_CShopKey { get; set; }
        public int m_CShopCount { get; set; }
        public List<SelectListItem> m_oSelCShopList { get; set; }
        //南
        public string m_SShopKey { get; set; }
        public int m_SShopCount { get; set; }
        public List<SelectListItem> m_oSelSShopList { get; set; }

        //public List<SelectListItem> m_oHRManagerList { get; set; }
        public HR_ManagerInfo m_oHRManager { get; set; }
        //public Object m_oHRManagerList { get; set; }

    }
}