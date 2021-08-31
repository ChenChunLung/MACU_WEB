using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using System.Data;
using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_UAK000.ViewModels
{
    public class MERP_UAK001_EditViewModel
    {             

        //20210106 CCL+
        public List<SelectListItem> m_oSelShopList { get; set; }

        //public List<SelectListItem> m_oHRManagerList { get; set; }
        public HR_ManagerInfo m_oHRManager { get; set; }
        //public Object m_oHRManagerList { get; set; }

    }
}