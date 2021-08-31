using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.Models._Base
{
    public static class AmountNumProcess
    {
        private static string CAMMA = ",";
        public static int NUM = 3;

        public static string m_sCamma {
            get
            {
                return CAMMA;
            }
            set
            {
                CAMMA = value;
            }
        }

        public static string ShowAmountComma(double p_dAmount)
        {
            string l_sRtnAmount = "";
            string l_sAmount = p_dAmount.ToString();

            if (p_dAmount == 0)
            {
                return "0";
            }

            //處理負數
            l_sRtnAmount = ShowMinusAmountPrint(l_sAmount);

            //如果是負數
            if (l_sRtnAmount.ElementAt(0) == '(')
            {
                //直接回傳
                return l_sRtnAmount;
            }

            //正數依NUM決定幾位數
            if (l_sAmount.Length > NUM)
            {
                l_sRtnAmount = l_sAmount.Insert(l_sAmount.Length - NUM, m_sCamma);
            } else
            {
                //l_sRtnAmount = l_sRtnAmount;
            }



            return l_sRtnAmount;

        }

        public static string ShowAmountComma(string p_sAmount)
        {
            string l_sRtnAmount = "";

            if ((p_sAmount == null) || (p_sAmount == ""))
            {
                return "0";
            }

            //處理負數
            l_sRtnAmount = ShowMinusAmountPrint(p_sAmount);

            //如果是負數
            if (l_sRtnAmount.ElementAt(0) == '(')
            {
                //直接回傳
                return l_sRtnAmount;
            }

            //正數依NUM決定幾位數
            if (p_sAmount.Length > NUM)
            {
              

                l_sRtnAmount = p_sAmount.Insert(p_sAmount.Length - NUM, m_sCamma);
            }
            else
            {
                //l_sRtnAmount = l_sRtnAmount;
            }



            return l_sRtnAmount;

        }


        public static string ShowMinusAmountPrint(string p_sAmount)
        {
            string l_sRtnAmount = "";
            string l_tmpStr = "";

            l_sRtnAmount = p_sAmount;

            //20210104 CCL- if ((p_sAmount != null) && 
            //   (Convert.ToInt32(p_sAmount) < 0))
            if ((p_sAmount != null) &&
               (Convert.ToDouble(p_sAmount) < 0))
            {
                
                //先幫去-後的值加上Camma
                if (p_sAmount.Substring(1).Length > NUM)
                {
                    
                    l_sRtnAmount = p_sAmount.Insert(p_sAmount.Length - NUM, m_sCamma);
                }
                //去掉-,改成()
                l_sRtnAmount = "(" + l_sRtnAmount.Substring(1) + ")";
            } else
            {
                //l_sRtnAmount = p_sAmount;


            }


            return l_sRtnAmount;
        }

        //20201231 CCL+
        public static string ShowMinusAmountPrint(double p_dAmount)
        {
            string l_sRtnAmount = "";
            string l_tmpStr = "";
            string l_sAmount = "";

            l_sRtnAmount = p_dAmount.ToString();
            l_sAmount = p_dAmount.ToString();

            //20210104 CCL- if ((l_sAmount != null) &&
            //   (Convert.ToInt32(l_sAmount) < 0))
            if ((l_sAmount != null) &&
               (p_dAmount < 0))
            {

                //先幫去-後的值加上Camma
                if (l_sAmount.Substring(1).Length > NUM)
                {

                    l_sRtnAmount = l_sAmount.Insert(l_sAmount.Length - NUM, m_sCamma);
                }
                //去掉-,改成()
                l_sRtnAmount = "(" + l_sRtnAmount.Substring(1) + ")";
            }
            else
            {
                //l_sRtnAmount = p_sAmount;


            }


            return l_sRtnAmount;
        }

        //20210103 CCL+ 比率負數顯示
        
        public static string ShowMinusPercent(double p_dAmount)
        {

            string l_sRtnAmount = "";
            string l_tmpStr = "";

            l_sRtnAmount = p_dAmount.ToString();
            

            if ((l_sRtnAmount != null) &&
               (p_dAmount < 0))
            {

               
                //去掉-,改成()
                l_sRtnAmount = "(" + l_sRtnAmount.Substring(1) + ")";
            }
            else
            {
                //l_sRtnAmount = p_sAmount;


            }


            return l_sRtnAmount;

            
        }

        public static string ShowMinusPercent(string p_sAmount)
        {

            string l_sRtnAmount = "";
            string l_tmpStr = "";

            l_sRtnAmount = p_sAmount;


            if ((p_sAmount != null) &&
               (Convert.ToDouble(p_sAmount) < 0))
            {

                //去掉-,改成()
                l_sRtnAmount = "(" + l_sRtnAmount.Substring(1) + ")";
            }
            else
            {
                //l_sRtnAmount = p_sAmount;


            }


            return l_sRtnAmount;
        }

        //20201229 CCL+ 如果為負數(XXX) 顯示紅色字//////////////////////////////////////////////////////
        public static bool ChkMinusNumRedFont(string p_sAmount)
        {
            const int HEAD_ROWS = 4; //新版表頭數目
            //int l_iRowIndex = p_iRowIndex;
            //int l_iPadColIndex = p_iPadColIndex;
            string l_sAmount = "";

            //IXLWorksheet l_oWooksheet = p_oWooksheet;

            //l_sAmount = AmountNumProcess.ShowAmountComma(p_sAmount);
            //l_sAmount = p_sAmount;

            if (!string.IsNullOrEmpty(p_sAmount))
            {
                //if(l_sAmount.Contains("("))
                if (p_sAmount.Contains("-"))
                {
                    //為負數設定Style Red Font
                    //IXLRange l_oRangeRedF = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    //l_oRangeRedF.Style.Font.SetFontColor(XLColor.Red);
                    return true;
                }
            }

            return false;
        }

        public static bool ChkMinusNumRedFont(double p_dAmount)
        {
            const int HEAD_ROWS = 4; //新版表頭數目
            //int l_iRowIndex = p_iRowIndex;
            //int l_iPadColIndex = p_iPadColIndex;
            string l_sAmount = "";

            //IXLWorksheet l_oWooksheet = p_oWooksheet;

            //l_sAmount = AmountNumProcess.ShowAmountComma(p_dAmount);

            //20210104 CCL- if (!string.IsNullOrEmpty(p_dAmount.ToString()))
            if (p_dAmount < 0)
            {
                //if (l_sAmount.Contains("("))
                //if (l_sAmount.Contains("-"))
                //{
                    //為負數設定Style Red Font
                    //IXLRange l_oRangeRedF = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    //l_oRangeRedF.Style.Font.SetFontColor(XLColor.Red);
                    return true;
                //}
            }

            return false;
        }
        // /////////////////////////////////////////////////////////////////////////////////////////////


    }
}