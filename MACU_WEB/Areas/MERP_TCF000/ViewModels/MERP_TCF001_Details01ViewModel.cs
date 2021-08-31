using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF001_Details01ViewModel
    {
        public string m_sYear { get; set; }

        public string m_sMonth { get; set; }

        public List<FA_LaborHealthInsV1> m_oFALaborHealthInsV1List { get; set; }
    }
}