using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public class ForSearch
    {
        //儲存,搜尋條件
        public string m_sSearch { get; set; }
        //public string m_sCondition { get; set; }

        public const char DELIMITER = ',';
        //public const string OR = "OR";
        //public const string AND = "AND";

        public List<string> m_oSearchList = null;

        public int m_iSearchTokenCount = 0;

        public ForSearch()
        {
            this.m_sSearch = "";
            //this.m_sCondition = "";
            m_oSearchList = new List<string>();

        }

        public ForSearch(string p_sSearch)
        {
            //紀錄原始
            this.m_sSearch = p_sSearch;
            
            if (!String.IsNullOrEmpty(p_sSearch))
            {
                if(p_sSearch.Contains(DELIMITER))
                {
                    this.m_oSearchList = GetSearchDataList(p_sSearch);
                    this.m_iSearchTokenCount = this.m_oSearchList.Count;
                }
                else
                {
                    m_oSearchList = new List<string>();
                    this.m_iSearchTokenCount = 0;
                }
                
            }
        }

        /*
        public ForSearch(string p_sSearch, string p_sCondition)
        {
            //紀錄原始
            this.m_sSearch = p_sSearch;
            this.m_sCondition = p_sCondition;

            if (!String.IsNullOrEmpty(p_sSearch))
            {
                this.m_oSearchList = GetSearchDataList(p_sSearch);
            }

            if (!String.IsNullOrEmpty(p_sSearch))
            {
                this.m_oSearchList = GetSearchDataList(p_sSearch);
            }
        }
        */

        public List<string> GetSearchDataList(string p_sSearch)
        {
            m_oSearchList = new List<string>();
            String[] l_oStrArry = p_sSearch.Split(DELIMITER);
            return l_oStrArry.ToList();
        }


    }
}