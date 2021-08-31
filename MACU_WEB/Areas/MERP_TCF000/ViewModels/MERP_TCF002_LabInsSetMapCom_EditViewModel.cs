using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF002_LabInsSetMapCom_EditViewModel
    {

        //加保公司
        public string m_sPlusCompany { get; set; }

        //原先選擇的勞保設定編號
        public string m_sOrgLaborInsSetNo { get; set; }

        //所有勞保設定SelectList
        public SelectList m_oLaborInsSetList { get; set; }

        //所有勞保設定 顯示列表
        public List<FA_LaborInsSet> m_oLaborInsSettings { get; set; }

       

    }
}