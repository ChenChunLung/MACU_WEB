using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.BIServices
{
    public class MERP_ProcessExcelOptions
    {
        //屬性: 當Controller Action收到View來的處理時
        private string m_sShopStr = "";
        private char m_cParseDelimiter = ',';

        //是否刪除Column Cell全為0的Rows
        public string m_sIsDelAllColZeroRows { get; set; }

        public string m_sAccountPeriod { get; set; }

        //20210204 CCL+ 改用年月抓
        public string m_sFiscalYear { get; set; }

        //區間: 開始
        public string m_sStartDate { get; set; }

        //區間: 結束
        public string m_sEndDate { get; set; }

        public string m_sManager { get; set; } //督導


        //20201227 CCL+ 多店鋪顯示 //////////////////////////////
        public string m_sTmpShopNo { get; set; } //
        public char m_cParseDelimi
        {
            get
            {
                return m_cParseDelimiter;
            }
            set
            {
                m_cParseDelimiter = value;
            }
        }
        public string m_sShop
        {
            get
            {
                return m_sShopStr;
            }
            set
            {
                m_sShopStr = value;
                //把值解析
                m_sShopList = ParseShopStrToken(m_sShopStr, m_cParseDelimi);
                m_iShopCount = m_sShopList.Count();
            }
        }

        public int m_iShopCount { get; set; } //店數

        public List<string> m_sShopList { get; set; }

        //Constructor
        public MERP_ProcessExcelOptions()
        {
            m_sShopList = new List<string>();

        }

        public List<string> ParseShopStrToken(String p_sShopStr, char p_cDelimi)
        {
            string[] l_aryShops = null;

            if(!string.IsNullOrEmpty(p_sShopStr))
            {
                l_aryShops = p_sShopStr.Split(p_cDelimi);
                return l_aryShops.ToList();
            }
            return null;
        }
        //20201227 CCL+ /////////////////////////////////////////
    }
}