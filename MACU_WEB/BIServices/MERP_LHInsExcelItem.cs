using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using MACU_WEB.Services;


namespace MACU_WEB.BIServices
{
    //紀錄勞健保計算後的值Item
    public class MERP_LHInsExcelItem
    {
        //<0>.基本資料
        public string m_sMemberName { get; set; }

        public string m_sPlusInsCompany { get; set; }

        public string m_sDepartName { get; set; }
        // ////////////////////////////////////////////////////////////


        public string m_OnJobDate { get; set; }

        public string m_ResignDate { get; set; }
        // ///////////////////////////////////////////////////////////
        public decimal m_DependentsNum { get; set; }

        //20210125 CCL+ 加入備註
        public int m_LHInsType { get; set; }

        //原始 勞保 值       
        public decimal m_dLaborInsOrgSalary { get; set; }

        //原始 健保 值
        public decimal m_dHealInsOrgSalary { get; set; }

        //依天數計算出來的薪水
        public decimal m_dLaborInsDatesSalary { get; set; }

        public decimal m_dHealInsDatesSalary { get; set; }
        // ////////////////////////////////////////////////////////////

        //<1>.勞保
        //  (1).單位(不含職災)
        public decimal m_dComLaborInsAmount { get; set; }
        //  (2).職災 金額
        public decimal m_dComLabOccuDisaInsAmount { get; set; }

        //  (3).個人
        public decimal m_dPerLaborInsAmount { get; set; }
        //  (4).單位(含職災) OccuDisaInsRate = (1) + (2)
        public decimal m_dComTolLaborODInsAmount { get; set; }
        //計算勞保-個人+單位 小計
        public decimal m_dTolLabPerComInsAmount { get; set; }

        //<5>.墊償
        //20210129 CCL+, 各公司的墊償和合計
        // (1).墊償
        public decimal m_dComLaborFundAmount { get; set; }
        // (2).墊償 小計 = 勞保-小計 + 墊償
        public decimal m_dTolLabComFundAmount { get; set; }

        //<2>.健保
        //  (1).單位
        public decimal m_dComHealInsAmount { get; set; }
        //  (2).個人
        public decimal m_dPerHealInsAmount { get; set; }
        //計算健保-個人+單位 小計
        public decimal m_dTolHealPerComInsAmount { get; set; }



        //<3>.勞退
        //  (1).單位
        public decimal m_dComRetireInsAmount { get; set; }
        //  (2).個人
        //public decimal m_dPerRetireInsAmount { get; set; }
        /// ///////////////////////////////////////////////////////////////////////////


        
        /// ///////////////////////////////////////////////////////////////////////////
        //<4>.合計
        //  (1).單位
        public decimal m_dTolComLHRInsAmount { get; set; }
        //  (2).個人
        public decimal m_dTolPerLHRInsAmount { get; set; }
        //  所有Cols Sum
        public decimal m_dTolAllLHRInsAmount { get; set; }
        // ////////////////////////////////////////////////////////////

        /* 20210127 CCL- 算法有誤,結果有誤差 
        public decimal Fun_CalcLabInsDateSalary(string p_sStartDate, string p_sEndDate)
        {
            //勞保投保天數:
            //(1)不管大小月,都以30天計算之。
            //(2)投保天數=30-上工日+1

            double l_dTolDayCounts = 0;
            int l_iYearDiff = 0, l_iHeadMonthDiff = 0, l_iTailMonthDiff = 0;
            int l_iTolMonths = 0, l_iYearMonths = 0;
            double l_dHeadDays = 0, l_dTailDays = 0;

            //四個日期會有切成三個區間
            DateTime l_oBeginDate = new DateTime();
            DateTime l_oOverDate = new DateTime();
            string l_sOnJobDate = "", l_sResignDate = "";
            string l_sStartDate = "", l_sEndDate = "";

            //20210122 CCL Mod 修正當到職離職日都是"",直接回傳0
            if (string.IsNullOrEmpty(m_OnJobDate) && string.IsNullOrEmpty(m_ResignDate))           
            {
                //計算總勞保薪資
                m_dLaborInsDatesSalary = 0 * m_dLaborInsDatesSalary;

                return m_dLaborInsDatesSalary;
            }


            if (string.IsNullOrEmpty(m_OnJobDate))
            {
                l_sOnJobDate = "0001/01/01";
            }
            else
            {
                l_sOnJobDate = m_OnJobDate;
            }

            if (string.IsNullOrEmpty(m_ResignDate))
            {
                l_sResignDate = "9999/12/30";
            } else
            {
                l_sResignDate = m_ResignDate;
            }

            if (string.IsNullOrEmpty(p_sStartDate))
            {
                l_sStartDate = "0001/01/01";
            } else
            {
                l_sStartDate = p_sStartDate;
            }

            if (string.IsNullOrEmpty(p_sEndDate))
            {
                l_sEndDate = "9999/12/30";
            } else
            {
                l_sEndDate = p_sEndDate;
            }

            DateTime l_oStartDate = Convert.ToDateTime(l_sStartDate);
            DateTime l_oEndDate = Convert.ToDateTime(l_sEndDate);
            DateTime l_oOnJobDate = Convert.ToDateTime(l_sOnJobDate);
            DateTime l_oResignDate = Convert.ToDateTime(l_sResignDate);

            //決定Begin Date /////////////////////////////////////
            if (l_oStartDate.CompareTo(l_oOnJobDate) < 0)
            {
                //選擇日期早於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) == 0)
            {
                //選擇日期等於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) > 0)
            {
                //就職日早於 選擇日期(選擇日期 晚於 就職日)
                //以選擇日期為主
                l_oBeginDate = l_oStartDate;
            }

            //決定Over Date //////////////////////////////////////
            if (l_oEndDate.CompareTo(l_oResignDate) < 0)
            {
                //選擇日期早於 離職日
                //以選擇日期為主
                l_oOverDate = l_oEndDate;
            }
            else if (l_oEndDate.CompareTo(l_oResignDate) == 0)
            {
                //選擇日期等於 離職日
                //以離職日為主
                l_oOverDate = l_oResignDate;
            }
            else if (l_oEndDate.CompareTo(l_oResignDate) > 0)
            {
                //離職日早於 選擇日期
                //以離職日為主
                l_oOverDate = l_oResignDate;
            }

            //For Debug
            if (m_OnJobDate == "2020/10/8")
            {
                string l_test = m_OnJobDate;
            }
            if (l_sResignDate == "2020/2/18")
            {
                string l_test = l_sResignDate;
            }
            if (l_sEndDate == "2020/10/31")
            {
                string l_test = l_sEndDate;
            }

            //如果End < OnJob Resign < Start 選擇日期完全在OnJob-Resign區間外 不計算天數
            if ((l_oEndDate.CompareTo(l_oOnJobDate) < 0) || (l_oStartDate.CompareTo(l_oResignDate) > 0))
            {
                //選擇日期完全超出區間外 不計算天數
                l_dTolDayCounts = 0;
            }
            else
            {
                //TimeSpan l_oDaysSpan = l_oOverDate.Subtract(l_oBeginDate);
                //計算總天數
                //l_dTolDayCounts = l_oDaysSpan.Days; //計算錯誤.少一天
                //MACU計算方式,如果是31天算30天

                //l_dTolDayCounts = l_oDaysSpan.Days;

                //如果滿一整個月(不管大小月,都以30天計算之),就以上傳的LaborInsOrgSalary勞保薪資下去乘幾個月
                //頭尾不足月的以天數下去算
                l_iYearDiff = l_oOverDate.Year - l_oBeginDate.Year;
                //if(l_iYearDiff )

                //年分間距大於等於1
                if(l_oBeginDate.Month >= 1 && l_iYearDiff >= 1)
                {
                    l_iHeadMonthDiff = (12 - l_oBeginDate.Month);
                } 
                if(l_oOverDate.Month <= 12 && l_iYearDiff >= 1)
                {
                    l_iTailMonthDiff = l_oOverDate.Month - 1;
                }
                
                //同一年只計算月份
                if(l_iYearDiff == 0)
                {
                    if(l_oOverDate.Month == l_oBeginDate.Month)
                    {
                        //又同月
                        l_iHeadMonthDiff = 0;
                    } else
                    {
                        l_iHeadMonthDiff = (l_oOverDate.Month - l_oBeginDate.Month) - 1;
                    }
                    
                    l_iTailMonthDiff = 0;
                    
                }
                //同一年[0年],(隔年)差一年[1年] 不乘以年分X12
                if (l_iYearDiff >= 2)
                {
                    l_iYearMonths = (l_iYearDiff - 1) * 12;
                } else
                {
                    l_iYearMonths = 0;
                }

                l_iTolMonths = l_iHeadMonthDiff + l_iYearMonths +  l_iTailMonthDiff;

                //DateTime l_oNextMonthDay1 = new DateTime(l_oBeginDate.Year, l_oBeginDate.Month + 1, 1);
                //l_dHeadDays = l_oNextMonthDay1.Subtract(l_oBeginDate).Days;
                //到職日直接用30天下去減(不管大小月30,31日,2月28,29日),一律用30天下去減
                //超過30天,以30算  
                if (l_oBeginDate.Day < 30)
                {
                    l_dHeadDays = (30 - l_oBeginDate.Day) + 1;
                }
                else { l_dHeadDays = 30;  }

                if(l_oOverDate.Day > 30 )
                {

                    l_dTailDays = 30;
                } else 
                {

                    l_dTailDays = l_oOverDate.Day;
                }

                //同一年,又同月,改用其他算法
                if (l_iYearDiff == 0)
                {
                    if (l_oOverDate.Month == l_oBeginDate.Month)
                    {
                        //又同月,以離職日算天數
                        l_dTailDays = 0;
                        l_dHeadDays = 0;
                        if(l_oBeginDate.Day > 30)
                        {
                            l_dHeadDays = 30;
                        } else
                        {
                            l_dHeadDays = l_oBeginDate.Day;
                        }

                        if(l_oOverDate.Day > 30)
                        {
                            //用30下去減
                            l_dTailDays = 30;
                        } else
                        {
                            l_dTailDays = l_oOverDate.Day;
                        }

                        //此CASE只加l_dTailDays
                        l_dTailDays = (l_dTailDays - l_dHeadDays) + 1;
                        //清空l_dHeadDays,只以l_dTailDays值為主l_dHeadDays借來運算而已
                        l_dHeadDays = 0;

                    }
                }
                                                           
                l_dTolDayCounts = l_dHeadDays + l_iTolMonths*30 + l_dTailDays;

            }



            //計算基數
            decimal l_dMolecular = m_dLaborInsOrgSalary / 30; //勞保

            //計算總勞保薪資
            m_dLaborInsDatesSalary = (decimal)l_dTolDayCounts * l_dMolecular;
            //四捨五入 ->整數 CCL- 等最後算完再四捨五入
            //m_dLaborInsDatesSalary = Math.Round(m_dLaborInsDatesSalary);


            return m_dLaborInsDatesSalary;
        }
        */

        /// 20210127 CCL+ 修正結果誤差,改算法: 總薪資 = 原匯入投保薪資 X (天數/30)
        public decimal Fun_CalcLabInsDateSalary(string p_sStartDate, string p_sEndDate)
        {
            //勞保投保天數:
            //(1)不管大小月,都以30天計算之。
            //(2)投保天數=30-上工日+1

            double l_dTolDayCounts = 0;
            int l_iYearDiff = 0, l_iHeadMonthDiff = 0, l_iTailMonthDiff = 0;
            int l_iTolMonths = 0, l_iYearMonths = 0;
            double l_dHeadDays = 0, l_dTailDays = 0;

            //四個日期會有切成三個區間
            DateTime l_oBeginDate = new DateTime();
            DateTime l_oOverDate = new DateTime();
            string l_sOnJobDate = "", l_sResignDate = "";
            string l_sStartDate = "", l_sEndDate = "";

            //20210122 CCL Mod 修正當到職離職日都是"",直接回傳0
            if (string.IsNullOrEmpty(m_OnJobDate) && string.IsNullOrEmpty(m_ResignDate))
            {
                //計算總勞保薪資
                m_dLaborInsDatesSalary = 0 * m_dLaborInsDatesSalary;

                return m_dLaborInsDatesSalary;
            }


            if (string.IsNullOrEmpty(m_OnJobDate))
            {
                l_sOnJobDate = "0001/01/01";
            }
            else
            {
                l_sOnJobDate = m_OnJobDate;
            }

            if (string.IsNullOrEmpty(m_ResignDate))
            {
                l_sResignDate = "9999/12/30";
            }
            else
            {
                l_sResignDate = m_ResignDate;
            }

            if (string.IsNullOrEmpty(p_sStartDate))
            {
                l_sStartDate = "0001/01/01";
            }
            else
            {
                l_sStartDate = p_sStartDate;
            }

            if (string.IsNullOrEmpty(p_sEndDate))
            {
                l_sEndDate = "9999/12/30";
            }
            else
            {
                l_sEndDate = p_sEndDate;
            }

            DateTime l_oStartDate = Convert.ToDateTime(l_sStartDate);
            DateTime l_oEndDate = Convert.ToDateTime(l_sEndDate);
            DateTime l_oOnJobDate = Convert.ToDateTime(l_sOnJobDate);
            DateTime l_oResignDate = Convert.ToDateTime(l_sResignDate);

            //決定Begin Date /////////////////////////////////////
            if (l_oStartDate.CompareTo(l_oOnJobDate) < 0)
            {
                //選擇日期早於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) == 0)
            {
                //選擇日期等於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) > 0)
            {
                //就職日早於 選擇日期(選擇日期 晚於 就職日)
                //以選擇日期為主
                l_oBeginDate = l_oStartDate;
            }

            //決定Over Date //////////////////////////////////////
            if (l_oEndDate.CompareTo(l_oResignDate) < 0)
            {
                //選擇日期早於 離職日
                //以選擇日期為主
                l_oOverDate = l_oEndDate;
            }
            else if (l_oEndDate.CompareTo(l_oResignDate) == 0)
            {
                //選擇日期等於 離職日
                //以離職日為主
                l_oOverDate = l_oResignDate;
            }
            else if (l_oEndDate.CompareTo(l_oResignDate) > 0)
            {
                //離職日早於 選擇日期
                //以離職日為主
                l_oOverDate = l_oResignDate;
            }

            //For Debug
            if (m_OnJobDate == "2020/10/8")
            {
                string l_test = m_OnJobDate;
            }
            if (l_sResignDate == "2020/2/18")
            {
                string l_test = l_sResignDate;
            }
            if (l_sEndDate == "2020/10/31")
            {
                string l_test = l_sEndDate;
            }

            //如果End < OnJob Resign < Start 選擇日期完全在OnJob-Resign區間外 不計算天數
            if ((l_oEndDate.CompareTo(l_oOnJobDate) < 0) || (l_oStartDate.CompareTo(l_oResignDate) > 0))
            {
                //選擇日期完全超出區間外 不計算天數
                l_dTolDayCounts = 0;
            }
            else
            {
                //TimeSpan l_oDaysSpan = l_oOverDate.Subtract(l_oBeginDate);
                //計算總天數
                //l_dTolDayCounts = l_oDaysSpan.Days; //計算錯誤.少一天
                //MACU計算方式,如果是31天算30天

                //l_dTolDayCounts = l_oDaysSpan.Days;

                //如果滿一整個月(不管大小月,都以30天計算之),就以上傳的LaborInsOrgSalary勞保薪資下去乘幾個月
                //頭尾不足月的以天數下去算
                l_iYearDiff = l_oOverDate.Year - l_oBeginDate.Year;
                //if(l_iYearDiff )

                //年分間距大於等於1
                if (l_oBeginDate.Month >= 1 && l_iYearDiff >= 1)
                {
                    l_iHeadMonthDiff = (12 - l_oBeginDate.Month);
                }
                if (l_oOverDate.Month <= 12 && l_iYearDiff >= 1)
                {
                    l_iTailMonthDiff = l_oOverDate.Month - 1;
                }

                //同一年只計算月份
                if (l_iYearDiff == 0)
                {
                    if (l_oOverDate.Month == l_oBeginDate.Month)
                    {
                        //又同月
                        l_iHeadMonthDiff = 0;
                    }
                    else
                    {
                        l_iHeadMonthDiff = (l_oOverDate.Month - l_oBeginDate.Month) - 1;
                    }

                    l_iTailMonthDiff = 0;

                }
                //同一年[0年],(隔年)差一年[1年] 不乘以年分X12
                if (l_iYearDiff >= 2)
                {
                    l_iYearMonths = (l_iYearDiff - 1) * 12;
                }
                else
                {
                    l_iYearMonths = 0;
                }

                l_iTolMonths = l_iHeadMonthDiff + l_iYearMonths + l_iTailMonthDiff;

                //DateTime l_oNextMonthDay1 = new DateTime(l_oBeginDate.Year, l_oBeginDate.Month + 1, 1);
                //l_dHeadDays = l_oNextMonthDay1.Subtract(l_oBeginDate).Days;
                //到職日直接用30天下去減(不管大小月30,31日,2月28,29日),一律用30天下去減
                //超過30天,以30算  
                if (l_oBeginDate.Day < 30)
                {
                    l_dHeadDays = (30 - l_oBeginDate.Day) + 1;
                }
                else { l_dHeadDays = 30; }

                if (l_oOverDate.Day > 30)
                {

                    l_dTailDays = 30;
                }
                else
                {

                    l_dTailDays = l_oOverDate.Day;
                }

                //同一年,又同月,改用其他算法
                if (l_iYearDiff == 0)
                {
                    if (l_oOverDate.Month == l_oBeginDate.Month)
                    {
                        //又同月,以離職日算天數
                        l_dTailDays = 0;
                        l_dHeadDays = 0;
                        if (l_oBeginDate.Day > 30)
                        {
                            l_dHeadDays = 30;
                        }
                        else
                        {
                            l_dHeadDays = l_oBeginDate.Day;
                        }

                        if (l_oOverDate.Day > 30)
                        {
                            //用30下去減
                            l_dTailDays = 30;
                        }
                        else
                        {
                            l_dTailDays = l_oOverDate.Day;
                        }

                        //此CASE只加l_dTailDays
                        l_dTailDays = (l_dTailDays - l_dHeadDays) + 1;
                        //清空l_dHeadDays,只以l_dTailDays值為主l_dHeadDays借來運算而已
                        l_dHeadDays = 0;

                    }
                }

                l_dTolDayCounts = l_dHeadDays + l_iTolMonths * 30 + l_dTailDays;

            }



            //計算基數
            //20210127 CCL- decimal l_dMolecular = m_dLaborInsOrgSalary / 30; //勞保
            //總薪資 = 原匯入投保薪資 X(天數 / 30)
            decimal l_dMolecular = (decimal)l_dTolDayCounts / 30; //勞保

            //計算總勞保薪資
            //20210127 CCL- m_dLaborInsDatesSalary = (decimal)l_dTolDayCounts * l_dMolecular;
            m_dLaborInsDatesSalary = m_dLaborInsOrgSalary * l_dMolecular;

            //四捨五入 ->整數 CCL- 等最後算完再四捨五入
            //m_dLaborInsDatesSalary = Math.Round(m_dLaborInsDatesSalary);


            return m_dLaborInsDatesSalary;
        }



        //健保
        public decimal Fun_CalcHealInsDateSalary(string p_sStartDate, string p_sEndDate)
        {
            //健保以月分算
            double l_dTolMonthCounts = 0;
            bool l_bIsResignSubLastOneMonth = false;
            
            DateTime l_oBeginDate = new DateTime();
            DateTime l_oOverDate = new DateTime();
            string l_sOnJobDate = "", l_sResignDate = "";
            string l_sStartDate = "", l_sEndDate = "";

            int l_iYearDiff = 0, l_iHeadMonthDiff = 0, l_iTailMonthDiff = 0;
            int l_iTolMonths = 0, l_iYearMonths = 0;

            //20210122 CCL Mod 修正當到職離職日都是"",直接回傳0
            if (string.IsNullOrEmpty(m_OnJobDate) && string.IsNullOrEmpty(m_ResignDate))
            {
                //計算總健保薪資
                m_dHealInsDatesSalary = 0 * m_dHealInsOrgSalary;
                
                return m_dHealInsDatesSalary;
            }


            if (string.IsNullOrEmpty(m_OnJobDate))
            {
                l_sOnJobDate = "0001/01/01";
            }
            else
            {
                l_sOnJobDate = m_OnJobDate;
            }

            if (string.IsNullOrEmpty(m_ResignDate))
            {
                l_sResignDate = "9999/12/30";
            }
            else
            {
                l_sResignDate = m_ResignDate;
            }

            if (string.IsNullOrEmpty(p_sStartDate))
            {
                l_sStartDate = "0001/01/01";
            }
            else
            {
                l_sStartDate = p_sStartDate;
            }

            if (string.IsNullOrEmpty(p_sEndDate))
            {
                l_sEndDate = "9999/12/30";
            }
            else
            {
                l_sEndDate = p_sEndDate;
            }

            DateTime l_oStartDate = Convert.ToDateTime(l_sStartDate);
            DateTime l_oEndDate = Convert.ToDateTime(l_sEndDate);
            DateTime l_oOnJobDate = Convert.ToDateTime(l_sOnJobDate);
            DateTime l_oResignDate = Convert.ToDateTime(l_sResignDate);

            //決定Begin Date /////////////////////////////////////
            if (l_oStartDate.CompareTo(l_oOnJobDate) < 0)
            {
                //選擇日期早於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) == 0)
            {
                //選擇日期等於 就職日
                //以就職日為主
                l_oBeginDate = l_oOnJobDate;
            }
            else if (l_oStartDate.CompareTo(l_oOnJobDate) > 0)
            {
                //就職日早於 選擇日期
                //以選擇日期為主
                l_oBeginDate = l_oStartDate;
            }

            //決定Over Date //////////////////////////////////////
            if (l_oEndDate.CompareTo(l_oResignDate) < 0)
            {
                //選擇日期早於 離職日
                //以選擇日期為主
                l_oOverDate = l_oEndDate;
                //必須判斷是否選擇日期與離職日位於同一年的同一個月,是的話必須減,而如果是月底最後一天不用減
                if ( (l_oEndDate.Year == l_oResignDate.Year) && 
                    (l_oEndDate.Month == l_oResignDate.Month))
                {
                    l_bIsResignSubLastOneMonth = true; //必須減去離職那最後一個月不算健保
                }

                //20210122 CCL Mod 需判斷當離職日位於選擇日期那一個月,需要再判斷離職日是否是當月月底最後一日,是的話要算那一個月健保
                //Ex: 選擇日期:2020/11/1 ~ 2020/11/30; 離職日: 2020/11/30; 11月共有30日
                if (l_oOverDate.Day == DateTime.DaysInMonth(l_oOverDate.Year, l_oOverDate.Month))
                {
                    l_bIsResignSubLastOneMonth = false; //離職那最後一個月要算健保
                }

            }
            else if (l_oEndDate.CompareTo(l_oResignDate) == 0)
            {
                //選擇日期等於 離職日
                //以離職日為主
                l_oOverDate = l_oResignDate;
                l_bIsResignSubLastOneMonth = true; //必須減去離職那最後一個月不算健保

                //20210122 CCL Mod 需判斷當離職日位於選擇日期那一個月,需要再判斷離職日是否是當月月底最後一日,是的話要算那一個月健保
                //Ex: 選擇日期:2020/11/1 ~ 2020/11/30; 離職日: 2020/11/30; 11月共有30日
                if (l_oOverDate.Day == DateTime.DaysInMonth(l_oOverDate.Year, l_oOverDate.Month))
                {
                    l_bIsResignSubLastOneMonth = false; //離職那最後一個月要算健保
                }
            }
            else if (l_oEndDate.CompareTo(l_oResignDate) > 0)
            {
                //離職日早於 選擇日期
                //以離職日為主
                l_oOverDate = l_oResignDate;
                l_bIsResignSubLastOneMonth = true; //必須減去離職那最後一個月不算健保

                //20210122 CCL Mod 需判斷當離職日位於選擇日期那一個月,需要再判斷離職日是否是當月月底最後一日,是的話要算那一個月健保
                //Ex: 選擇日期:2020/11/1 ~ 2020/11/30; 離職日: 2020/11/30; 11月共有30日
                if (l_oOverDate.Day == DateTime.DaysInMonth(l_oOverDate.Year, l_oOverDate.Month))
                {
                    l_bIsResignSubLastOneMonth = false; //離職那最後一個月要算健保
                }
            }


            //For Debug
            if (l_sResignDate == "2020/11/30")
            {
                string l_test = l_sResignDate;
            }
            if (l_sEndDate == "2020/11/30")
            {
                string l_test = l_sEndDate;
            }


            //如果End < OnJob Resign < Start 選擇日期完全在OnJob-Resign區間外 不計算月數
            if ((l_oEndDate.CompareTo(l_oOnJobDate) < 0) || (l_oStartDate.CompareTo(l_oResignDate) > 0))
            {
                //選擇日期完全超出區間外 不計算月數
                l_dTolMonthCounts = 0;
            }
            else
            {
                l_iYearDiff = l_oOverDate.Year - l_oBeginDate.Year;

                //年分間距大於等於1
                if (l_oBeginDate.Month >= 1 && l_iYearDiff >= 1)
                {
                    l_iHeadMonthDiff = (12 - l_oBeginDate.Month);
                }
                if (l_oOverDate.Month <= 12 && l_iYearDiff >= 1)
                {
                    
                    l_iTailMonthDiff = l_oOverDate.Month - 1;
                }

                //同一年只計算月份
                if (l_iYearDiff == 0)
                {
                    if (l_oOverDate.Month == l_oBeginDate.Month)
                    {
                        //又同月
                        l_iHeadMonthDiff = 0;
                    }
                    else
                    {
                        l_iHeadMonthDiff = (l_oOverDate.Month - l_oBeginDate.Month) - 1;
                    }

                    l_iTailMonthDiff = 0;

                }
                //同一年[0年],(隔年)差一年[1年] 不乘以年分X12
                if (l_iYearDiff >= 2)
                {
                    l_iYearMonths = (l_iYearDiff - 1) * 12;
                }
                else
                {
                    l_iYearMonths = 0;
                }

                l_iTolMonths = l_iHeadMonthDiff + l_iYearMonths + l_iTailMonthDiff;


                
                if (l_bIsResignSubLastOneMonth)
                {
                    //計算總月數  (離職那一個月不算,所以減1個月)
                    //l_dTolMonthCounts = (l_oOverDate.Year * 12 + l_oOverDate.Month - 1) -
                    //                    (l_oBeginDate.Year * 12 + l_oBeginDate.Month);
                    l_dTolMonthCounts = l_iTolMonths - 1;

                    l_dTolMonthCounts += 1; // 相減後要+1
                } else
                {
                    //計算總月數  (選擇日期低於離職日那一年那一個月,不必減1個月)
                    //l_dTolMonthCounts = (l_oOverDate.Year * 12 + l_oOverDate.Month) -
                    //                    (l_oBeginDate.Year * 12 + l_oBeginDate.Month);
                    l_dTolMonthCounts = l_iTolMonths;

                    l_dTolMonthCounts += 1; // 相減後要+1
                }
                


            }

            //計算總健保薪資
            m_dHealInsDatesSalary = (decimal)l_dTolMonthCounts * m_dHealInsOrgSalary;            
            //四捨五入 ->整數  CCL- 等最後算完再四捨五入
            //m_dHealInsDatesSalary = Math.Round(m_dHealInsDatesSalary);

            return m_dHealInsDatesSalary;
        }
        

        /// ////////////////////////////////////////////////////////////////////////////////////////////////////



    }
}