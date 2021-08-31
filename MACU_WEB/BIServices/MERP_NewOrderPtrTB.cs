using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.BIServices
{
    public class MERP_NewOrderPtrTB
    {
        //public List<TmpExcelItem> m_NewOrderItemPtrs { get; set; }
        //Shop Name
        public string m_sShopName { get; set; }
        public string m_sShopNo { get; set; }
        //20210103 CCL+ 加上RowCount
        public int m_TBTolRowCount { get; set; }

        //放其他計算double值
        public List<double> m_NewOrderItemPtrs { get; set; }



        //分群Base項目
        public List<TmpExcelItem> m_oNOAttrItemsGID4 { get; set; }
        public List<TmpExcelItem> m_oNOAttrItemsGID5 { get; set; }
        public List<TmpExcelItem> m_oNOAttrItemsGID6 { get; set; }
        public List<TmpExcelItem> m_oNOAttrItemsGID7 { get; set; }
        public List<TmpExcelItem> m_oNOAttrItemsGID8 { get; set; }
       
        //Beginning of raw materials 原料期初
        public List<TmpExcelItem> m_oNOAttrBeginRawMaterial { get; set; }
        //End of raw material period 減 原料期末
        public List<TmpExcelItem> m_oNOAttrEndRawMaterial { get; set; }

        //20201227 CCL+ 計算比率
        public double Calc_PercentVal(double p_dValue, double p_dOperaIncome)
        {
            //(p_dValue / m_dOperaIncome ) * 100 //營業收入
            double l_dRtnVal = 0;
            const int l_iDotNum = 2;//四捨五入至小數點第2位

            //壁面無限大
            if ((p_dOperaIncome != 0))
            {

                l_dRtnVal = (p_dValue / p_dOperaIncome) * 100;
                l_dRtnVal = Math.Round(l_dRtnVal, l_iDotNum);

                return l_dRtnVal;
            }
            return l_dRtnVal;
        }

        public double Calc_PercentVal(double p_dValue)
        {
            //(p_dValue / m_dOperaIncome ) * 100 //營業收入
            double l_dRtnVal = 0;
            const int l_iDotNum = 2;//四捨五入至小數點第2位

            //壁面無限大
            if ( (m_NewOrderItemPtrs[0] != 0))
            {

                l_dRtnVal = (p_dValue / m_NewOrderItemPtrs[0]) * 100;
                l_dRtnVal = Math.Round(l_dRtnVal, l_iDotNum);

                return l_dRtnVal;
            }
            return l_dRtnVal;
        }

        public MERP_NewOrderPtrTB()
        {
            //m_NewOrderItemPtrs = new List<TmpExcelItem>();
            m_NewOrderItemPtrs = new List<double>();

            //20210103 CCL+ TolRowCount
            m_TBTolRowCount = 0;

            m_oNOAttrItemsGID4 = new List<TmpExcelItem>();
            m_oNOAttrItemsGID5 = new List<TmpExcelItem>();
            m_oNOAttrItemsGID6 = new List<TmpExcelItem>();
            m_oNOAttrItemsGID7 = new List<TmpExcelItem>();
            m_oNOAttrItemsGID8 = new List<TmpExcelItem>();

            m_oNOAttrBeginRawMaterial = new List<TmpExcelItem>();
            m_oNOAttrEndRawMaterial = new List<TmpExcelItem>();
        }
    }
}