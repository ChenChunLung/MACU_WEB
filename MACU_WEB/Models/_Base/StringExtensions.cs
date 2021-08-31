using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public static class StringExtensions
    {
        //Modify MultiLangResx.Resources.Resource to your [Namespace].[Resource name]
        public static String ToAutoMultiLang(this String source)
        {
            if (string.IsNullOrEmpty(source)) return "zzz";
            //return MultiLangResx.Resources.Resource.ResourceManager.GetString(source) ?? source;
            return "";
        }

        public static String ToLocalMultiLang(this String source)
        {
            //return MSP.MVC.Model.Models.Local.ResourceManager.GetString(source) ?? source;
            return "";
        }
    }
}