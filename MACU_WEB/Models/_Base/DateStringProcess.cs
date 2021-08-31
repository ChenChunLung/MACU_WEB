using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public static class DateStringProcess
    {
        public static int m_Year { get; set; }

        public static int m_Month { get; set; }

        public static int m_Day { get; set; }

        //去除DateTime字串中
        public static string Del_MonthDayZero(string p_sYearMonthDay, string p_cDelimiter, string p_cNewDelimiter)
        {
            string l_sTmpStr = "";
            int l_iMonth = 0;
            int l_iDay = 0;
            int l_iYear = 0;

            //去掉後面的" 上午 HH:MM:SS"
            if (p_sYearMonthDay.Contains(' '))
            {
                l_sTmpStr = p_sYearMonthDay.Substring(0, p_sYearMonthDay.IndexOf(' '));
            }
            else
            {
                l_sTmpStr = p_sYearMonthDay;
            }


            l_iYear = Convert.ToInt32(l_sTmpStr.Substring(0, l_sTmpStr.IndexOf(p_cDelimiter)));
            l_sTmpStr = l_sTmpStr.Substring(l_sTmpStr.IndexOf(p_cDelimiter) + 1);
            l_iMonth = Convert.ToInt32(l_sTmpStr.Substring(0, l_sTmpStr.IndexOf(p_cDelimiter)));

            l_sTmpStr = l_sTmpStr.Substring(l_sTmpStr.IndexOf(p_cDelimiter) + 1);
            l_iDay = Convert.ToInt32(l_sTmpStr.Substring(0));

            m_Year = l_iYear;
            m_Month = l_iMonth;
            m_Day = l_iDay;

            if (p_cNewDelimiter == "")
            {
                return l_iYear + p_cDelimiter + l_iMonth + p_cDelimiter + l_iDay;

            }
            else
            {

                return l_iYear + p_cNewDelimiter + l_iMonth + p_cNewDelimiter + l_iDay;
            }


        }

        //20210107 CCL+ For <input type='date' value='xxxx-0x-0x'> Control
        public static string ReStore_MonthDayZero(string p_sYearMonthDay, string p_cDelimiter, string p_cNewDelimiter)
        {
            string l_sTmpStr = "";
            int l_iMonth = 0;
            int l_iDay = 0;
            int l_iYear = 0;
            string l_sYear = "";
            string l_sDay = "";
            string l_sMonth = "";

            //去掉後面的" 上午 HH:MM:SS"
            if (p_sYearMonthDay.Contains(' '))
            {
                l_sTmpStr = p_sYearMonthDay.Substring(0, p_sYearMonthDay.IndexOf(' '));
            }
            else
            {
                l_sTmpStr = p_sYearMonthDay;
            }


            l_iYear = Convert.ToInt32(l_sTmpStr.Substring(0, l_sTmpStr.IndexOf(p_cDelimiter)));
            l_sTmpStr = l_sTmpStr.Substring(l_sTmpStr.IndexOf(p_cDelimiter) + 1);
            l_iMonth = Convert.ToInt32(l_sTmpStr.Substring(0, l_sTmpStr.IndexOf(p_cDelimiter)));

            l_sTmpStr = l_sTmpStr.Substring(l_sTmpStr.IndexOf(p_cDelimiter) + 1);
            l_iDay = Convert.ToInt32(l_sTmpStr.Substring(0));


            m_Year = l_iYear;
            m_Month = l_iMonth;
            m_Day = l_iDay;

            if (l_iYear < 10)
            {
                l_sYear = "0" + l_iYear;
            }
            else
            {
                l_sYear = l_iYear.ToString();
            }

            if (l_iMonth < 10)
            {
                l_sMonth = "0" + l_iMonth;
            }
            else
            {
                l_sMonth = l_iMonth.ToString();
            }

            if (l_iDay < 10)
            {
                l_sDay = "0" + l_iDay;
            }
            else
            {
                l_sDay = l_iDay.ToString();
            }



            if (p_cNewDelimiter == "")
            {
                return l_sYear + p_cDelimiter + l_sMonth + p_cDelimiter + l_sDay;

            }
            else
            {

                return l_sYear + p_cNewDelimiter + l_sMonth + p_cNewDelimiter + l_sDay;
            }


        }

        //20210111 CCL+, 上傳勞健保資料時,手動輸入的年月
        //去除Month字串中的0
        public static string Del_MonthZero(string p_sYearMonth, string p_cDelimiter, string p_cNewDelimiter)
        {
            string l_sTmpStr = "";
            int l_iMonth = 0;           
            int l_iYear = 0;

            //去掉後面的" 上午 HH:MM:SS"
            if (p_sYearMonth.Contains(' '))
            {
                l_sTmpStr = p_sYearMonth.Substring(0, p_sYearMonth.IndexOf(' '));
            }
            else
            {
                l_sTmpStr = p_sYearMonth;
            }


            l_iYear = Convert.ToInt32(l_sTmpStr.Substring(0, l_sTmpStr.IndexOf(p_cDelimiter)));
            l_sTmpStr = l_sTmpStr.Substring(l_sTmpStr.IndexOf(p_cDelimiter) + 1);
            l_iMonth = Convert.ToInt32(l_sTmpStr);
           
            m_Year = l_iYear;
            m_Month = l_iMonth;           

            if (p_cNewDelimiter == "")
            {
                return l_iYear + p_cDelimiter + l_iMonth + p_cDelimiter;

            }
            else
            {

                return l_iYear + p_cNewDelimiter + l_iMonth + p_cNewDelimiter;
            }

        }


    }
}