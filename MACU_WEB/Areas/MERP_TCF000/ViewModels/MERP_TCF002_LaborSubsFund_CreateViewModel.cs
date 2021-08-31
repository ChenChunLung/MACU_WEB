using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MACU_WEB.Models;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF002_LaborSubsFund_CreateViewModel
    {

        public SelectList m_oPlusComInsList { get; set; }

        //已存在在Db中的FundSet 設定
        public List<FA_LaborSubsFundSet> m_oExistedLSFundSetList { get; set; }

    }
}