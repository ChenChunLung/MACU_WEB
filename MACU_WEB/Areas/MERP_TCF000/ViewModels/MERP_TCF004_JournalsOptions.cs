using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Areas.MERP_TCF000.ViewModels
{
    public class MERP_TCF004_JournalsOptions
    {

        //屬性: 當Controller Action收到View來的處理時   

        
        //public int m_iDataCount { get; set; }

        public string m_sDataYear { get; set; }
        public string m_sDataMonth { get; set; }

        //區間: 開始 就職日
        public string m_sOnJobDate { get; set; }

        //區間: 結束 離職日
        public string m_sResignDate { get; set; }

        public string m_sMemberName { get; set; } //員工姓名

        public string m_sShopName { get; set; } //部門


    }
}