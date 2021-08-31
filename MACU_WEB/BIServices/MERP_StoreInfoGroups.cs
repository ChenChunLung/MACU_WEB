using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Diagnostics;

namespace MACU_WEB.BIServices
{
    public static class MERP_StoreInfoGroups
    {
        //NORMAL:一般, DIRECTLY:直營, COMMON:合營, JOIN:加盟
        public enum STOREINFO_GROUPS
        {
            NORMAL = 0,
            DIRECTLY = 1,
            COMMON = 2,
            JOIN = 3,
            DIRECOMMON = 4,
            DIRECOMJOIN = 5
        }; //20210225 CCL+ DIRCOMMON = 4


        /// <summary>
        ///     回傳 SelectListItem
        /// </summary>
        /// <returns> 
        ///     IEnumerable<SelectListItem>
        /// </returns>
        public static IEnumerable<SelectListItem> GetStoreGroupSetListItem()
        {
            //MERP_StoreInfoGroups
            int l_iTmp = 0;
            

            IEnumerable<SelectListItem> l_oRtnData = new List<SelectListItem> {
                //new SelectListItem { Text = "一般",
                //                       Value = (l_iTmp = 0).ToString(),
                //                     Selected = true},
                new SelectListItem { Text = "直營",
                                       Value = (l_iTmp = 1).ToString() },
                new SelectListItem { Text = "合營",
                                       Value = (l_iTmp = 2).ToString() },
                new SelectListItem { Text = "加盟",
                                       Value = (l_iTmp = 3).ToString() },
                //20210225 CCL+ 直合營
                new SelectListItem { Text = "直合營",
                                       Value = (l_iTmp = 4).ToString() }
               
            };


            return l_oRtnData;
        }

        public static SelectList GetStoreGroupSetSelList()
        {
            //MERP_StoreInfoGroups
            IEnumerable<SelectListItem> l_oTmpData = GetStoreGroupSetListItem();

            SelectList l_oSelGroupList = new SelectList(l_oTmpData, "Value", "Text");
            //SelectList l_oSelGroupList = new SelectList(l_oTmpData, "Text", "Value");

            return l_oSelGroupList;
        }


        //20210225 CCL+ 增加區域Type 全區(不分區) 北區 中區 南區 離島 外國
        public static IEnumerable<SelectListItem> GetStoreGroupSetTypeListItem()
        {
            

            IEnumerable<SelectListItem> l_oRtnData = new List<SelectListItem> {

                new SelectListItem { Text = "不分區",
                                       Value =  "A" },
                new SelectListItem { Text = "北區",
                                       Value =  "N" },
                new SelectListItem { Text = "中區",
                                       Value =  "C" },
                new SelectListItem { Text = "南區",
                                       Value =  "S" },               
                new SelectListItem { Text = "離島",
                                       Value =  "I" },
                new SelectListItem { Text = "外國",
                                       Value =  "F" }
            };


            return l_oRtnData;
        }

        public static SelectList GetStoreGroupSetTypeSelList()
        {
            //MERP_StoreInfoGroups
            IEnumerable<SelectListItem> l_oTmpData = GetStoreGroupSetTypeListItem();

            SelectList l_oSelGroupList = new SelectList(l_oTmpData, "Value", "Text");
            //SelectList l_oSelGroupList = new SelectList(l_oTmpData, "Text", "Value");

            return l_oSelGroupList;
        }



    }
}