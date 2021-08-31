using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF001_DetailsViewModel
    {
        public string m_sYear { get; set; }

        public string m_sMonth { get; set; }

        public List<FA_LaborHealthIns> m_oFALaborHealthInsList { get; set; }

    }
}