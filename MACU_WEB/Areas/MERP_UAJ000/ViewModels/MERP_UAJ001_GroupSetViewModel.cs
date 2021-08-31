using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Web.Mvc;
using MACU_WEB.BIServices;

namespace MACU_WEB.Areas.MERP_UAJ000.ViewModels
{
    public class MERP_UAJ001_GroupSetViewModel
    {
        //年/月
        public string m_sFiscalYear { get; set; }
        public string m_sAccountPeroid { get; set; }

        //取出直合營設定SelectItem
        public SelectList m_oStoreInfoGroup { get; set; }

        //20210225 CCL+ 取出直合營設定Type區域 SelectItem
        public SelectList m_oStoreInfoGroupSetType { get; set; }

        //public List<FA_JournalV1> m_oFaJournalList { get; set; }

        //北
        public string m_sNShopKey { get; set; }
        public int m_iNShopCount { get; set; }
        public List<SelectListItem> m_oSelNShopList { get; set; }
        //中
        public string m_sCShopKey { get; set; }
        public int m_iCShopCount { get; set; }
        public List<SelectListItem> m_oSelCShopList { get; set; }
        //南
        public string m_sSShopKey { get; set; }
        public int m_iSShopCount { get; set; }
        public List<SelectListItem> m_oSelSShopList { get; set; }

        //20210202 CCL+
        public StoreGroupSet m_oToEditItem { get; set; }

        public List<StoreGroupSet> m_oStoreGroupSetList { get; set; }



    }
}