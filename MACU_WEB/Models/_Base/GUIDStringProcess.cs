using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public static class GUIDStringProcess
    {
        public static string m_GuidStr { get; set; }

        public static string GuidProcess(string p_sGuidStr)
        {
            string l_sRtnStr = "";

            if(p_sGuidStr != "")
            {
                l_sRtnStr = p_sGuidStr.Replace("{", "");
                l_sRtnStr = l_sRtnStr.Replace("}", "");
            }
           

            return l_sRtnStr;
        }

    }
}