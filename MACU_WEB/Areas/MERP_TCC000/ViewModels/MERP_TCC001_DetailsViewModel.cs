using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data;
using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TCC000.ViewModels
{
    public class MERP_TCC001_DetailsViewModel
    {
        public string m_AccountPeroid { get; set; }
        //20210204 CCL+ 改用年月抓
        public string m_sFiscalYear { get; set; }

        //20201223 CCL- public List<FA_FaJournal> m_FaJournalList { get; set; }
        public List<FA_JournalV1> m_FaJournalList { get; set; }

        //20210106 CCL+
        public List<SelectListItem> m_oSelShopList { get; set; }

        public List<SelectListItem> m_oHRManagerList { get; set; }
        //public Object m_oHRManagerList { get; set; }
    }
}