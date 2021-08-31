using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data;
using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TCC000.ViewModels
{
    public class MERP_TCC001_Details01ViewModel
    {
        public string m_AccountPeroid { get; set; }

        public string m_FiscalYear { get; set; }

        //20201223 CCL- public List<FA_FaJournal> m_FaJournalList { get; set; }
        public List<FA_JournalV1> m_FaJournalList { get; set; }

        //20210106 CCL+
        //public List<SelectListItem> m_oSelShopList { get; set; }

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

        public List<SelectListItem> m_oHRManagerList { get; set; }
        //public Object m_oHRManagerList { get; set; }

        //20210204 CCL+ 直合營設定List
        //20210225 CCL- public List<StoreGroupSet> m_oStoreGroupSetList { get; set; }
        //20210225 CCL+ 改用選擇
        public SelectList m_oStoreGroupSetSelList { get; set; }

    }
}