using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF002_LabInsSetMapCom_CreateViewModel
    {
        //加保公司
        public SelectList m_oPlusComInsList { get; set; }

        //所有勞保設定SelectList
        public SelectList m_oLaborInsSetList { get; set; }

        //所有勞保設定
        public List<FA_LaborInsSet> m_oLaborInsSettings { get; set; }

        //已存在在Db中的LaborInsSetMapComSet 設定
        public List<FA_LaborInsSetMapComSet> m_oExistedLabInsMapPlusComSetList { get; set; }



    }
}