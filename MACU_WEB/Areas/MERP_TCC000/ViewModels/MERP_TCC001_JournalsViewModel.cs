using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using MACU_WEB.Models._Base;

namespace MACU_WEB.Areas.MERP_TCC000.ViewModels
{
    public class MERP_TCC001_JournalsViewModel
    {
        public string m_AccountPeroid { get; set; }

        public List<FA_FaJournal> m_FaJournalList { get; set; }

        //分頁功能用
        public ForPaging m_Paging { get; set; }
        //搜尋功能用
        public ForSearch m_Search { get; set; }
    }
}