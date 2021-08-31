using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public class ForPaging
    {
        //當前頁數
        public int m_iNowPage { get; set; }

        //最大頁數
        public int m_iMaxPage { get; set; }

        //一頁筆數
        public int m_iItemNum { get; set; }

        //Constructor
        public ForPaging()
        {
            //預設當前頁數為1
            this.m_iNowPage = 1;
        }

        public ForPaging(int p_iPage)
        {
            this.m_iNowPage = p_iPage;
        }

        //設定正確頁數的方法,以避免輸入不正確的值
        public void SetRightPage()
        {
            //當判斷值小於1
            if(this.m_iNowPage < 1)
            {
                this.m_iNowPage = 1;
            }
            //當判斷當前值大於總頁數
            else if(this.m_iNowPage > this.m_iMaxPage)
            {
                //設定當前值等於總頁數
                this.m_iNowPage = this.m_iMaxPage;
            }

            //當無資料時的當前頁數,重設為1
            if(this.m_iMaxPage.Equals(0))
            {
                this.m_iNowPage = 1;
            }


        }
    }
}