using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using MACU_WEB.Services;
using System.Data;
using System.Diagnostics;

namespace MACU_WEB.BIServices
{
    //勞健保Excel主輸出物件
    public class MERP_LHInsExcelExpore
    {        

        public List<MERP_LHInsExcelItem> m_oLHInsExcelResultItems { get; set; }

        //依公司計算合計
        public List<MERP_LHInsExcelItem> m_oLHInsExcelByComResults { get; set; }

        //依公司合計 計算 墊償 //20210129 CCL+ 放各公司墊償結果
        public List<MERP_LHInsExcelItem> m_oLInsFundByComResults { get; set; }


        //依公司-門市計算合計
        public List<MERP_LHInsExcelItem> m_oLHInsExcelByComDepResults { get; set; }

        //公司數目
        public int m_iCompanyCount { get; set; }
        //公司-門市數目
        public int m_iComDeptCount { get; set; }

        // /////////////////////////////////////////////////////////////////
        //User選擇日期區間
        public string m_StartDate { get; set; }

        public string m_EndDate { get; set; }
        // /////////////////////////////////////////////////////////////////
        //20210129 CCL+, 新增勞保設定對應表
        //public List<FA_LaborInsSetMapComSet> m_oLInsSetMapComSet { get; set; }
        private MERP_FA_LaborInsSetMapComSetDBService m_oLInsSetMapComSetDBService; //DBService
        //20210129 CCL+, 為了取出勞保設定,新增目前投保公司名全域變數
        private string m_sCurrentUsedPlusCom = "";

        //最新勞保費率設定
        public FA_LaborInsSet m_oLaborInsSet { get; set; }

        //最新健保費率設定
        public FA_HealthInsSet m_oHealInsSet { get; set; }

        //總投保薪資 (所有勞保薪資相加)
        public double m_dTolLaborInsVal { get; set; }

        //總投保薪資總額度Quota: 
        public decimal m_dTolLaborInsValQuota { get; set; }

        //總健保薪資 (所有健保薪資相加)
        public double m_dTolHealInsVal { get; set; }

        /// ///////////////////////////////////////////////////////////////////////////
        
        ///////////////////////// 最底下總和那一行 ////////////////////////////////////
        //所有勞保-單位 值相加
        public double m_dComLaborInsTolAmounts { get; set; }

        //所有勞保-個人 值相加
        public double m_dPerLaborInsTolAmounts { get; set; }

        //20210125 CCL+ 所有勞保-個人+單位 小計 值相加  
        public double m_dPerComLaborInsTolAmounts { get; set; }

        //20210129 CCL+ 墊償總計
        //所有墊償 值相加
        public double m_dLaborInsFundTolAmounts { get; set; }

        //所有墊償 小計 值相加  
        public double m_dComTolLInsFundTolAmounts { get; set; }


        //所有健保-單位 值相加
        public double m_dComHealInsTolAmounts { get; set; }

        //所有健保-個人 值相加
        public double m_dPerHealInsTolAmounts { get; set; }

        //20210125 CCL+ 所有健保-個人+單位 小計 值相加  
        public double m_dPerComHealInsTolAmounts { get; set; }


        //所有勞退-單位 值相加
        public double m_dComRetireInsTolAmounts { get; set; }

        //所有LHR合計(總和)-個人 值相加
        public double m_dPerLHRInsTolAmounts { get; set; }

        //所有LHR合計(總和)-單位 值相加
        public double m_dComLHRInsTolAmounts { get; set; }

        //所有LHR合計(總和) 值相加
        public double m_dAllLHRInsTolAmounts { get; set; }

        /// ////////////////////////////////////////////////////////////////////////////////////////////
        /// 
        //20210129 CCL+, 根據PlusCompany公司名找出[勞保設定對應投保公司設定]對應表中所對應的該勞保設定流水號
        //藉此抓出該勞保設定
        public double Fun_GetLInsSetByPlusComMapSet(string p_sPlusInsCompany)
        {
            //如果PlusCompany不依樣,更新勞保設定
            if(m_oLInsSetMapComSetDBService != null )
            {
                if(m_sCurrentUsedPlusCom != p_sPlusInsCompany)
                {
                    //取出下一個公司 勞保設定
                    FA_LaborInsSet l_oTmpLInsSet =
                    m_oLInsSetMapComSetDBService.FA_LaborInsSetMapComSet_GetDataLInsSetByPlusInsCom(p_sPlusInsCompany);
                    //更新目前要使用的勞保設定
                    m_oLaborInsSet = l_oTmpLInsSet;
                    //更新目前使用公司名
                    m_sCurrentUsedPlusCom = p_sPlusInsCompany;
                    //計算 店家墊償
                    MERP_LHInsExcelItem l_oTmpComFundItem = new MERP_LHInsExcelItem();
                    l_oTmpComFundItem.m_sPlusInsCompany = p_sPlusInsCompany;
                    l_oTmpComFundItem.m_dComLaborFundAmount = Fun_CalcTolLaborInsValQuotaV2();
                    m_oLInsFundByComResults.Add(l_oTmpComFundItem);

                }

            }

            return 0;
        }

        //20210129 CCL+ 
        //計算 店家墊償 版本2
        public decimal Fun_CalcTolLaborInsValQuotaV2()
        {
            decimal l_dComLaborFundAmount = 0;

            if (m_oLaborInsSet != null)
            {                

                //公式: 勞保代墊基金 X 勞保代墊基金費率/100
                decimal l_dLaborSubsFundRate = Convert.ToDecimal(m_oLaborInsSet.LaborSubsFundRate) / 100;
                decimal l_dLaborSubsFund = Convert.ToDecimal(m_oLaborInsSet.LaborSubsFund);

                l_dComLaborFundAmount = l_dLaborSubsFund * l_dLaborSubsFundRate;
                //四捨五入 -> 整數
                l_dComLaborFundAmount = Math.Round(l_dComLaborFundAmount, MidpointRounding.AwayFromZero);
            }

            return l_dComLaborFundAmount;
        }


        //計算 所有人勞保總薪資
        public double Fun_CalcAllLaborInsVals(Object p_sLaborInsVal)
        {
            double l_iVal = 0;
            if(p_sLaborInsVal == null)
            {
                l_iVal = 0;
            } else
            {
                l_iVal = Convert.ToDouble(p_sLaborInsVal);
            }

            m_dTolLaborInsVal += l_iVal;

            return m_dTolLaborInsVal;
        }


        //計算 總投保薪資總額度Quota 版本0
        public decimal Fun_CalcTolLaborInsValQuota()
        {
            if (m_oLaborInsSet != null )
            {
                //公式: 總投保薪資 (所有勞保薪資相加) X 勞保代墊基金費率/100
                decimal l_dLaborSubsFundRate = Convert.ToDecimal(m_oLaborInsSet.LaborSubsFundRate) / 100;

                m_dTolLaborInsValQuota = (decimal)m_dTolLaborInsVal * l_dLaborSubsFundRate;
                //四捨五入 -> 整數
                m_dTolLaborInsValQuota = Math.Round(m_dTolLaborInsValQuota, MidpointRounding.AwayFromZero);
            }

            return m_dTolLaborInsValQuota;
        }


        //計算 總投保薪資總額度Quota 版本1
        public decimal Fun_CalcTolLaborInsValQuotaV1()
        {
            if (m_oLaborInsSet != null)
            {
                //公式: 勞保代墊基金 X 勞保代墊基金費率/100
                decimal l_dLaborSubsFundRate = Convert.ToDecimal(m_oLaborInsSet.LaborSubsFundRate) / 100;
                decimal l_dLaborSubsFund = Convert.ToDecimal(m_oLaborInsSet.LaborSubsFund) ;

                m_dTolLaborInsValQuota = l_dLaborSubsFund * l_dLaborSubsFundRate;
                //四捨五入 -> 整數
                m_dTolLaborInsValQuota = Math.Round(m_dTolLaborInsValQuota, MidpointRounding.AwayFromZero);
            }

            return m_dTolLaborInsValQuota;
        }

        //計算All 版本2
        public int Fun_CalcAllLHInsResult(DataSet p_oDTItemData)
        {
            int l_iRowCount = 0;
            decimal l_dLaborIns = 0;
            decimal l_dHealthIns = 0;
            decimal l_dDependentsNum = 0;

            if ((m_oLaborInsSet != null) &&
                (p_oDTItemData != null) && (p_oDTItemData.Tables[0].Rows.Count > 0))
            {

                //計算
                foreach (DataRow row in p_oDTItemData.Tables[0].Rows)
                {
                    l_iRowCount++;

                    MERP_LHInsExcelItem l_oLHInsItem = new MERP_LHInsExcelItem();

                    if (row["MemberName"].ToString() == "丁紹恩")
                    {
                        Trace.WriteLine(l_oLHInsItem.m_sMemberName);
                    }

                    //基本資料
                    l_oLHInsItem.m_sMemberName = row["MemberName"].ToString();
                    l_oLHInsItem.m_sPlusInsCompany = row["PlusInsCompany"].ToString();
                    l_oLHInsItem.m_sDepartName = row["DepartName"].ToString();
                    l_oLHInsItem.m_OnJobDate = row["OnJobDate"].ToString();
                    l_oLHInsItem.m_ResignDate = row["ResignDate"].ToString();
                    l_oLHInsItem.m_LHInsType = Convert.ToInt32(row["LHInsType"].ToString()); //20210125 CCL+ 加上備註
                    if (string.IsNullOrEmpty(row["Dependents"].ToString()))
                    { l_dDependentsNum = 0; }
                    else { l_dDependentsNum = Convert.ToDecimal(row["Dependents"]); }
                    l_oLHInsItem.m_DependentsNum = l_dDependentsNum; //20210121 CCL+
                    if (string.IsNullOrEmpty(row["LaborIns"].ToString()))
                    { l_dLaborIns = 0; }
                    else { l_dLaborIns = Convert.ToDecimal(row["LaborIns"]); }
                    l_oLHInsItem.m_dLaborInsOrgSalary = l_dLaborIns;
                    if (string.IsNullOrEmpty(row["HealthIns"].ToString()))
                    { l_dHealthIns = 0; }
                    else { l_dHealthIns = Convert.ToDecimal(row["HealthIns"]); }
                    l_oLHInsItem.m_dHealInsOrgSalary = l_dHealthIns;
                    l_oLHInsItem.Fun_CalcLabInsDateSalary(m_StartDate, m_EndDate); //計算選擇日期總天數勞保薪資
                    l_oLHInsItem.Fun_CalcHealInsDateSalary(m_StartDate, m_EndDate); //計算選擇日期總天數健保薪資


                    //--------------------- 勞保Str ----------------------------------------
                    //20210129 CCL+ 根據不同公司更新套用不同勞保設定
                    Fun_GetLInsSetByPlusComMapSet(l_oLHInsItem.m_sPlusInsCompany); //20210129 CCL+

                    //計算總投保薪資
                    //Fun_CalcAllLaborInsVals(row["LaborIns"]);
                    Fun_CalcAllLaborInsVals(l_oLHInsItem.m_dLaborInsDatesSalary); //20210119 CCL+


                    //計算勞保-個人 ==> [勞保_個人負擔 Expore]
                    //l_oLHInsItem.m_dPerLaborInsAmount = Fun_CalcPerLaborInsResult(row);
                    l_oLHInsItem.m_dPerLaborInsAmount = Fun_CalcPerLaborInsResult(l_oLHInsItem);
                    //計算勞保-單位(不含職災)
                    //l_oLHInsItem.m_dComLaborInsAmount = Fun_CalcComLaborInsResult(row);
                    l_oLHInsItem.m_dComLaborInsAmount = Fun_CalcComLaborInsResult(l_oLHInsItem);
                    //計算職災-只有單位
                    //l_oLHInsItem.m_dComLabOccuDisaInsAmount = Fun_CalcOccuDisaInsResult(row);
                    l_oLHInsItem.m_dComLabOccuDisaInsAmount = Fun_CalcOccuDisaInsResult(l_oLHInsItem);
                    //計算勞保-單位(含職災) ==> [勞保_單位負擔 Expore]
                    l_oLHInsItem.m_dComTolLaborODInsAmount = l_oLHInsItem.m_dComLaborInsAmount +
                                                             l_oLHInsItem.m_dComLabOccuDisaInsAmount;
                    //20210125 CCL+ 修正必須與職災相加後才能四捨五入,不然會有誤差
                    //20210127 CCL- 改成各自四捨五入後再相加 l_oLHInsItem.m_dComTolLaborODInsAmount = Math.Round(l_oLHInsItem.m_dComTolLaborODInsAmount);

                    //計算勞退-只有單位 ==> [勞退_單位負擔 Expore]
                    //l_oLHInsItem.m_dComRetireInsAmount = Fun_CalcLaborRetireInsResult(row);
                    l_oLHInsItem.m_dComRetireInsAmount = Fun_CalcLaborRetireInsResult(l_oLHInsItem);

                    //計算勞保-個人+單位 小計
                    l_oLHInsItem.m_dTolLabPerComInsAmount = l_oLHInsItem.m_dPerLaborInsAmount +
                                                            l_oLHInsItem.m_dComTolLaborODInsAmount;

                    //--------------------- 勞保End ----------------------------------------

                    //--------------------- 健保Str ----------------------------------------
                    //計算健保-個人
                    //l_oLHInsItem.m_dPerHealInsAmount = Fun_CalcPerHealInsResult(row);
                    l_oLHInsItem.m_dPerHealInsAmount = Fun_CalcPerHealInsResult(l_oLHInsItem);
                    //計算健保-單位
                    //l_oLHInsItem.m_dComHealInsAmount = Fun_CalcComHealInsResult(row);
                    l_oLHInsItem.m_dComHealInsAmount = Fun_CalcComHealInsResult(l_oLHInsItem);

                    //計算健保-個人+單位 小計
                    l_oLHInsItem.m_dTolHealPerComInsAmount = l_oLHInsItem.m_dPerHealInsAmount +
                                                            l_oLHInsItem.m_dComHealInsAmount;
                    //--------------------- 健保End ----------------------------------------


                    //--------------------- 合計Str ----------------------------------------
                    //計算合計 (L Labor + H Heal + R Retire = T TolLHR)
                    //Fun_CalcTolLHRInsResult(l_oLHInsItem); //20210118 CCL- 改成合計也分單位,個人
                    //計算合計-個人 (L PerLabor個人 + H PerHeal個人  = T PerTolLHR個人)
                    l_oLHInsItem.m_dTolPerLHRInsAmount = Fun_CalcTolPerLHRInsResult(l_oLHInsItem);
                    //計算合計-單位 (L ComLabor單位 + H ComHeal單位 + R Retire單位  = T ComTolLHR單位)
                    l_oLHInsItem.m_dTolComLHRInsAmount = Fun_CalcTolComLHRInsResult(l_oLHInsItem);
                    //計算合計-所有 = 合計-單位 + 合計-個人
                    l_oLHInsItem.m_dTolAllLHRInsAmount = Fun_CalcTolAllLHRInsResult(l_oLHInsItem);
                    //--------------------- 合計End ----------------------------------------

                    //--------------------- 總和Str ----------------------------------------
                    //[最底部總和]
                    //所有勞保-單位 值相加
                    m_dComLaborInsTolAmounts += (double)l_oLHInsItem.m_dComTolLaborODInsAmount;
                    //所有勞保-個人 值相加
                    m_dPerLaborInsTolAmounts += (double)l_oLHInsItem.m_dPerLaborInsAmount;
                    //20210125 CCL+ 所有勞保-個人+單位 小計 值相加  
                    m_dPerComLaborInsTolAmounts += (double)l_oLHInsItem.m_dTolLabPerComInsAmount;

                    //所有健保-單位 值相加
                    m_dComHealInsTolAmounts += (double)l_oLHInsItem.m_dComHealInsAmount;
                    //所有健保-個人 值相加
                    m_dPerHealInsTolAmounts += (double)l_oLHInsItem.m_dPerHealInsAmount;
                    //20210125 CCL+ 所有健保-個人+單位 小計 值相加  
                    m_dPerComHealInsTolAmounts += (double)l_oLHInsItem.m_dTolHealPerComInsAmount;

                    //所有勞退-單位 值相加
                    m_dComRetireInsTolAmounts += (double)l_oLHInsItem.m_dComRetireInsAmount;

                    //所有LHR合計-個人 值相加
                    m_dPerLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolPerLHRInsAmount;
                    //所有LHR合計-單位 值相加
                    m_dComLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolComLHRInsAmount;
                    //所有LHR合計 值相加
                    m_dAllLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolAllLHRInsAmount;
                    //--------------------- 總和End ----------------------------------------

                    //加入List
                    m_oLHInsExcelResultItems.Add(l_oLHInsItem);
                }

                //計算 總投保薪資總額度Quota 版本0
                Fun_CalcTolLaborInsValQuota();

                //依公司計算合計
                //m_iCompanyCount = Fun_CalcLHInsByComResults(m_oLHInsExcelResultItems);
                //依公司-門市計算合計
                m_iComDeptCount = Fun_CalcLHInsByComDepResults(m_oLHInsExcelResultItems);
                //依公司-門市合計 計算 公司合計
                m_iCompanyCount = Fun_CalcLHInsByComanyResults(m_oLHInsExcelByComDepResults); //20210120 CCL+
                //依公司合計 計算 墊償 墊償合計
                Fun_CalcLInsFundByCompanyResults(m_oLHInsExcelByComResults); //20210129 CCL+
                //20210225 CCL+ 利用計算出的墊償, 墊償小計 算出各門市 墊償比例
                Fun_CalcLHInsFundPercentByComanyResults(m_oLHInsExcelByComResults);


            }
            return l_iRowCount;
        }

        /*
        //計算All 版本1
        public int Fun_CalcAllLHInsResult(DataSet p_oDTItemData)
        {
            int l_iRowCount = 0;
            decimal l_dLaborIns = 0;
            decimal l_dHealthIns = 0;
            decimal l_dDependentsNum = 0;

            if ((m_oLaborInsSet != null) &&
                (p_oDTItemData != null) && (p_oDTItemData.Tables[0].Rows.Count > 0))
            {

                //計算
                foreach (DataRow row in p_oDTItemData.Tables[0].Rows)
                {
                    l_iRowCount++;

                    MERP_LHInsExcelItem l_oLHInsItem = new MERP_LHInsExcelItem();

                    if (row["MemberName"].ToString() == "陳煜妍")
                    {
                        Trace.WriteLine(l_oLHInsItem.m_sMemberName);
                    }

                    //基本資料
                    l_oLHInsItem.m_sMemberName = row["MemberName"].ToString();
                    l_oLHInsItem.m_sPlusInsCompany = row["PlusInsCompany"].ToString();                    
                    l_oLHInsItem.m_sDepartName = row["DepartName"].ToString();
                    l_oLHInsItem.m_OnJobDate = row["OnJobDate"].ToString();                    
                    l_oLHInsItem.m_ResignDate = row["ResignDate"].ToString();
                    l_oLHInsItem.m_LHInsType = Convert.ToInt32(row["LHInsType"].ToString()); //20210125 CCL+ 加上備註
                    if (string.IsNullOrEmpty(row["Dependents"].ToString()))
                    { l_dDependentsNum = 0; }
                    else { l_dDependentsNum = Convert.ToDecimal(row["Dependents"]); }
                    l_oLHInsItem.m_DependentsNum = l_dDependentsNum; //20210121 CCL+
                    if (string.IsNullOrEmpty(row["LaborIns"].ToString()))
                    { l_dLaborIns = 0; } else { l_dLaborIns = Convert.ToDecimal(row["LaborIns"]); }
                    l_oLHInsItem.m_dLaborInsOrgSalary = l_dLaborIns;
                    if (string.IsNullOrEmpty(row["HealthIns"].ToString()))
                    { l_dHealthIns = 0; } else { l_dHealthIns = Convert.ToDecimal(row["HealthIns"]); }
                    l_oLHInsItem.m_dHealInsOrgSalary = l_dHealthIns;
                    l_oLHInsItem.Fun_CalcLabInsDateSalary(m_StartDate, m_EndDate); //計算選擇日期總天數勞保薪資
                    l_oLHInsItem.Fun_CalcHealInsDateSalary(m_StartDate, m_EndDate); //計算選擇日期總天數健保薪資


                    //--------------------- 勞保Str ----------------------------------------                 
                    //計算總投保薪資
                    //Fun_CalcAllLaborInsVals(row["LaborIns"]);
                    Fun_CalcAllLaborInsVals(l_oLHInsItem.m_dLaborInsDatesSalary); //20210119 CCL+


                    //計算勞保-個人 ==> [勞保_個人負擔 Expore]
                    //l_oLHInsItem.m_dPerLaborInsAmount = Fun_CalcPerLaborInsResult(row);
                    l_oLHInsItem.m_dPerLaborInsAmount = Fun_CalcPerLaborInsResult(l_oLHInsItem);
                    //計算勞保-單位(不含職災)
                    //l_oLHInsItem.m_dComLaborInsAmount = Fun_CalcComLaborInsResult(row);
                    l_oLHInsItem.m_dComLaborInsAmount = Fun_CalcComLaborInsResult(l_oLHInsItem);
                    //計算職災-只有單位
                    //l_oLHInsItem.m_dComLabOccuDisaInsAmount = Fun_CalcOccuDisaInsResult(row);
                    l_oLHInsItem.m_dComLabOccuDisaInsAmount = Fun_CalcOccuDisaInsResult(l_oLHInsItem);
                    //計算勞保-單位(含職災) ==> [勞保_單位負擔 Expore]
                    l_oLHInsItem.m_dComTolLaborODInsAmount = l_oLHInsItem.m_dComLaborInsAmount +
                                                             l_oLHInsItem.m_dComLabOccuDisaInsAmount;
                    //20210125 CCL+ 修正必須與職災相加後才能四捨五入,不然會有誤差
                    //20210127 CCL- 改成各自四捨五入後再相加 l_oLHInsItem.m_dComTolLaborODInsAmount = Math.Round(l_oLHInsItem.m_dComTolLaborODInsAmount);

                    //計算勞退-只有單位 ==> [勞退_單位負擔 Expore]
                    //l_oLHInsItem.m_dComRetireInsAmount = Fun_CalcLaborRetireInsResult(row);
                    l_oLHInsItem.m_dComRetireInsAmount = Fun_CalcLaborRetireInsResult(l_oLHInsItem);

                    //計算勞保-個人+單位 小計
                    l_oLHInsItem.m_dTolLabPerComInsAmount = l_oLHInsItem.m_dPerLaborInsAmount + 
                                                            l_oLHInsItem.m_dComTolLaborODInsAmount;
                   
                    //--------------------- 勞保End ----------------------------------------

                    //--------------------- 健保Str ----------------------------------------
                    //計算健保-個人
                    //l_oLHInsItem.m_dPerHealInsAmount = Fun_CalcPerHealInsResult(row);
                    l_oLHInsItem.m_dPerHealInsAmount = Fun_CalcPerHealInsResult(l_oLHInsItem);
                    //計算健保-單位
                    //l_oLHInsItem.m_dComHealInsAmount = Fun_CalcComHealInsResult(row);
                    l_oLHInsItem.m_dComHealInsAmount = Fun_CalcComHealInsResult(l_oLHInsItem);

                    //計算健保-個人+單位 小計
                    l_oLHInsItem.m_dTolHealPerComInsAmount = l_oLHInsItem.m_dPerHealInsAmount +
                                                            l_oLHInsItem.m_dComHealInsAmount;
                    //--------------------- 健保End ----------------------------------------


                    //--------------------- 合計Str ----------------------------------------
                    //計算合計 (L Labor + H Heal + R Retire = T TolLHR)
                    //Fun_CalcTolLHRInsResult(l_oLHInsItem); //20210118 CCL- 改成合計也分單位,個人
                    //計算合計-個人 (L PerLabor個人 + H PerHeal個人  = T PerTolLHR個人)
                    l_oLHInsItem.m_dTolPerLHRInsAmount = Fun_CalcTolPerLHRInsResult(l_oLHInsItem);
                    //計算合計-單位 (L ComLabor單位 + H ComHeal單位 + R Retire單位  = T ComTolLHR單位)
                    l_oLHInsItem.m_dTolComLHRInsAmount = Fun_CalcTolComLHRInsResult(l_oLHInsItem);
                    //計算合計-所有 = 合計-單位 + 合計-個人
                    l_oLHInsItem.m_dTolAllLHRInsAmount = Fun_CalcTolAllLHRInsResult(l_oLHInsItem);
                    //--------------------- 合計End ----------------------------------------

                    //--------------------- 總和Str ----------------------------------------
                    //[最底部總和]
                    //所有勞保-單位 值相加
                    m_dComLaborInsTolAmounts += (double)l_oLHInsItem.m_dComTolLaborODInsAmount;
                    //所有勞保-個人 值相加
                    m_dPerLaborInsTolAmounts += (double)l_oLHInsItem.m_dPerLaborInsAmount;
                    //20210125 CCL+ 所有勞保-個人+單位 小計 值相加  
                    m_dPerComLaborInsTolAmounts += (double)l_oLHInsItem.m_dTolLabPerComInsAmount;

                    //所有健保-單位 值相加
                    m_dComHealInsTolAmounts += (double)l_oLHInsItem.m_dComHealInsAmount;
                    //所有健保-個人 值相加
                    m_dPerHealInsTolAmounts += (double)l_oLHInsItem.m_dPerHealInsAmount;
                    //20210125 CCL+ 所有健保-個人+單位 小計 值相加  
                    m_dPerComHealInsTolAmounts += (double)l_oLHInsItem.m_dTolHealPerComInsAmount;

                    //所有勞退-單位 值相加
                    m_dComRetireInsTolAmounts += (double)l_oLHInsItem.m_dComRetireInsAmount;

                    //所有LHR合計-個人 值相加
                    m_dPerLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolPerLHRInsAmount;
                    //所有LHR合計-單位 值相加
                    m_dComLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolComLHRInsAmount;
                    //所有LHR合計 值相加
                    m_dAllLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolAllLHRInsAmount;
                    //--------------------- 總和End ----------------------------------------

                    //加入List
                    m_oLHInsExcelResultItems.Add(l_oLHInsItem);
                }

                //計算 總投保薪資總額度Quota 版本0
                Fun_CalcTolLaborInsValQuota();

                //依公司計算合計
                //m_iCompanyCount = Fun_CalcLHInsByComResults(m_oLHInsExcelResultItems);
                //依公司-門市計算合計
                m_iComDeptCount = Fun_CalcLHInsByComDepResults(m_oLHInsExcelResultItems);
                //依公司-門市合計 計算 公司合計
                m_iCompanyCount = Fun_CalcLHInsByComanyResults(m_oLHInsExcelByComDepResults); //20210120 CCL+


            }
            return l_iRowCount;
        }
        */

        /*
        //計算All
        public int Fun_CalcAllLHInsResult(DataSet p_oDTItemData)
        {
            int l_iRowCount = 0;

            if ((m_oLaborInsSet != null) &&
                (p_oDTItemData != null) && (p_oDTItemData.Tables[0].Rows.Count > 0))
            {
               
                //計算
                foreach (DataRow row in p_oDTItemData.Tables[0].Rows)
                {
                    l_iRowCount++;

                    MERP_LHInsExcelItem l_oLHInsItem = new MERP_LHInsExcelItem();

                    //基本資料
                    l_oLHInsItem.m_sMemberName = row["MemberName"].ToString();
                    l_oLHInsItem.m_sPlusInsCompany = row["PlusInsCompany"].ToString();
                    l_oLHInsItem.m_sDepartName = row["DepartName"].ToString();

                    //--------------------- 勞保Str ----------------------------------------
                    //計算總投保薪資
                    Fun_CalcAllLaborInsVals(row["LaborIns"]);

                    //計算勞保-個人 ==> [勞保_個人負擔 Expore]
                    l_oLHInsItem.m_dPerLaborInsAmount = Fun_CalcPerLaborInsResult(row);
                    //計算勞保-單位(不含職災)
                    l_oLHInsItem.m_dComLaborInsAmount = Fun_CalcComLaborInsResult(row);
                    //計算職災-只有單位
                    l_oLHInsItem.m_dComLabOccuDisaInsAmount = Fun_CalcOccuDisaInsResult(row);
                    //計算勞保-單位(含職災) ==> [勞保_單位負擔 Expore]
                    l_oLHInsItem.m_dComTolLaborODInsAmount = l_oLHInsItem.m_dComLaborInsAmount +
                                                             l_oLHInsItem.m_dComLabOccuDisaInsAmount;
                    //計算勞退-只有單位 ==> [勞退_單位負擔 Expore]
                    l_oLHInsItem.m_dComRetireInsAmount = Fun_CalcLaborRetireInsResult(row);

                    //--------------------- 勞保End ----------------------------------------

                    //--------------------- 健保Str ----------------------------------------
                    //計算健保-個人
                    l_oLHInsItem.m_dPerHealInsAmount = Fun_CalcPerHealInsResult(row);
                    //計算健保-單位
                    l_oLHInsItem.m_dComHealInsAmount = Fun_CalcComHealInsResult(row);
                    //--------------------- 健保End ----------------------------------------


                    //--------------------- 合計Str ----------------------------------------
                    //計算合計 (L Labor + H Heal + R Retire = T TolLHR)
                    //Fun_CalcTolLHRInsResult(l_oLHInsItem); //20210118 CCL- 改成合計也分單位,個人
                    //計算合計-個人 (L PerLabor個人 + H PerHeal個人  = T PerTolLHR個人)
                    l_oLHInsItem.m_dTolPerLHRInsAmount = Fun_CalcTolPerLHRInsResult(l_oLHInsItem);
                    //計算合計-單位 (L ComLabor單位 + H ComHeal單位 + R Retire單位  = T ComTolLHR單位)
                    l_oLHInsItem.m_dTolComLHRInsAmount = Fun_CalcTolComLHRInsResult(l_oLHInsItem);
                    //計算合計-所有 = 合計-單位 + 合計-個人
                    l_oLHInsItem.m_dTolAllLHRInsAmount = Fun_CalcTolAllLHRInsResult(l_oLHInsItem);
                    //--------------------- 合計End ----------------------------------------

                    //--------------------- 總和Str ----------------------------------------
                    //所有勞保-單位 值相加
                    m_dComLaborInsTolAmounts += (double)l_oLHInsItem.m_dComTolLaborODInsAmount;
                    //所有勞保-個人 值相加
                    m_dPerLaborInsTolAmounts += (double)l_oLHInsItem.m_dPerLaborInsAmount;
                    //所有健保-單位 值相加
                    m_dComHealInsTolAmounts += (double)l_oLHInsItem.m_dComHealInsAmount;
                    //所有健保-個人 值相加
                    m_dPerHealInsTolAmounts += (double)l_oLHInsItem.m_dPerHealInsAmount;
                    //所有勞退-單位 值相加
                    m_dComRetireInsTolAmounts += (double)l_oLHInsItem.m_dComRetireInsAmount;

                    //所有LHR合計-個人 值相加
                    m_dPerLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolPerLHRInsAmount;
                    //所有LHR合計-單位 值相加
                    m_dComLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolComLHRInsAmount;
                    //所有LHR合計 值相加
                    m_dAllLHRInsTolAmounts += (double)l_oLHInsItem.m_dTolAllLHRInsAmount;
                    //--------------------- 總和End ----------------------------------------

                    //加入List
                    m_oLHInsExcelResultItems.Add(l_oLHInsItem);
                }

                //計算 總投保薪資總額度Quota 版本0
                Fun_CalcTolLaborInsValQuota();

                //依公司計算合計
                //m_iCompanyCount = Fun_CalcLHInsByComResults(m_oLHInsExcelResultItems);
                //依公司-門市計算合計
                m_iComDeptCount = Fun_CalcLHInsByComDepResults(m_oLHInsExcelResultItems);


            }
            return l_iRowCount;
        }
        */

        //改成By 公司-門市 小計
        public int Fun_CalcLHInsByComDepResults(List<MERP_LHInsExcelItem> p_oLHInsResultItems)
        {
            int l_iRowIndex = 0;
            string l_sPrevCompanyName = "";
            string l_sPrevDepartName = "";

            decimal l_dComTolLaborODInsAmount = 0;
            decimal l_dPerLaborInsAmount = 0;
            decimal l_dComHealInsAmount = 0;
            decimal l_dPerHealInsAmount = 0;
            decimal l_dComRetireInsAmount = 0;
            decimal l_dTolComLHRInsAmount = 0;
            decimal l_dTolPerLHRInsAmount = 0;
            decimal l_dTolAllLHRInsAmount = 0;
            //20210125 CCL+ 計算勞健保-個人+單位 小計
            decimal l_dTolLabPerComInsAmount = 0;
            decimal l_dTolHealPerComInsAmount = 0;

            int l_iPrintComDepCount = 0;

            if ((p_oLHInsResultItems != null) &&
                (p_oLHInsResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in p_oLHInsResultItems)
                {
                    ++l_iRowIndex;
                    //找出公司-部門
                    //m_oLHInsExcelByComDepResults
                    if (l_sPrevCompanyName == "" && l_sPrevDepartName == "")
                    {
                        //第一筆
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount = Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount = Item.m_dTolHealPerComInsAmount;

                        //更新公司-部門
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;
                        l_sPrevDepartName = Item.m_sDepartName;
                    }
                    else if ((Item.m_sPlusInsCompany == l_sPrevCompanyName) &&
                             (Item.m_sDepartName == l_sPrevDepartName))
                    {
                        //累加公司-部門
                        l_dComTolLaborODInsAmount += Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount += Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount += Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount += Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount += Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount += Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount += Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount += Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount += Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount += Item.m_dTolHealPerComInsAmount;

                    }
                    else if ( (Item.m_sDepartName != l_sPrevDepartName) || 
                              (Item.m_sPlusInsCompany != l_sPrevCompanyName))
                    {


                        ++l_iPrintComDepCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;
                        l_oNewItem.m_sDepartName = l_sPrevDepartName;

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_oNewItem.m_dTolLabPerComInsAmount = l_dTolLabPerComInsAmount;
                        l_oNewItem.m_dTolHealPerComInsAmount = l_dTolHealPerComInsAmount;

                        m_oLHInsExcelByComDepResults.Add(l_oNewItem);

                        ////////////////////////////////////////////////////////////////
                        //更新公司
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;
                        l_sPrevDepartName = Item.m_sDepartName;
                        //設為這一輪的值
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount = Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount = Item.m_dTolHealPerComInsAmount;

                    }

                    if (l_iRowIndex == p_oLHInsResultItems.Count())
                    {
                        //最後一筆時也要執行這裡
                        ++l_iPrintComDepCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;
                        l_oNewItem.m_sDepartName = l_sPrevDepartName;                        

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_oNewItem.m_dTolLabPerComInsAmount = l_dTolLabPerComInsAmount;
                        l_oNewItem.m_dTolHealPerComInsAmount = l_dTolHealPerComInsAmount;

                        m_oLHInsExcelByComDepResults.Add(l_oNewItem);

                    }

                }
            }


            return l_iPrintComDepCount;
        }

        //利用算完的 公司-門市 小計 算出--> 公司 小計 
        public int Fun_CalcLHInsByComanyResults(List<MERP_LHInsExcelItem> p_oLHInsExcelByComDepItems)
        {
            int l_iRowIndex = 0;
            string l_sPrevCompanyName = "";
            string l_sPrevDepartName = "";

            decimal l_dComTolLaborODInsAmount = 0;
            decimal l_dPerLaborInsAmount = 0;
            decimal l_dComHealInsAmount = 0;
            decimal l_dPerHealInsAmount = 0;
            decimal l_dComRetireInsAmount = 0;
            decimal l_dTolComLHRInsAmount = 0;
            decimal l_dTolPerLHRInsAmount = 0;
            decimal l_dTolAllLHRInsAmount = 0;
            //20210125 CCL+ 計算勞健保-個人+單位 小計
            decimal l_dTolLabPerComInsAmount = 0;
            decimal l_dTolHealPerComInsAmount = 0;
          


            int l_iPrintComCount = 0;

            if ((p_oLHInsExcelByComDepItems != null) &&
                (p_oLHInsExcelByComDepItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in p_oLHInsExcelByComDepItems)
                {
                    ++l_iRowIndex;
                    //找出公司
                    if (l_sPrevCompanyName == "" && l_sPrevDepartName == "")
                    {
                        //第一筆
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount = Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount = Item.m_dTolHealPerComInsAmount;

                        //更新公司-部門
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;
                        l_sPrevDepartName = Item.m_sDepartName;
                    }
                    else if (Item.m_sPlusInsCompany == l_sPrevCompanyName)
                    {
                        //累加公司-部門
                        l_dComTolLaborODInsAmount += Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount += Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount += Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount += Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount += Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount += Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount += Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount += Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount += Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount += Item.m_dTolHealPerComInsAmount;

                    }
                    else if (Item.m_sPlusInsCompany != l_sPrevCompanyName)
                    {


                        ++l_iPrintComCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;
                        l_oNewItem.m_sDepartName = l_sPrevDepartName;

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_oNewItem.m_dTolLabPerComInsAmount = l_dTolLabPerComInsAmount;
                        l_oNewItem.m_dTolHealPerComInsAmount = l_dTolHealPerComInsAmount;

                        m_oLHInsExcelByComResults.Add(l_oNewItem);

                        ////////////////////////////////////////////////////////////////
                        //更新公司
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;
                        l_sPrevDepartName = Item.m_sDepartName;
                        //設為這一輪的值
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_dTolLabPerComInsAmount = Item.m_dTolLabPerComInsAmount;
                        l_dTolHealPerComInsAmount = Item.m_dTolHealPerComInsAmount;


                    }

                    if (l_iRowIndex == p_oLHInsExcelByComDepItems.Count())
                    {
                        //最後一筆時也要執行這裡
                        ++l_iPrintComCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;
                        l_oNewItem.m_sDepartName = l_sPrevDepartName;

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;
                        //20210125 CCL+ 計算勞健保-個人+單位 小計
                        l_oNewItem.m_dTolLabPerComInsAmount = l_dTolLabPerComInsAmount;
                        l_oNewItem.m_dTolHealPerComInsAmount = l_dTolHealPerComInsAmount;

                        m_oLHInsExcelByComResults.Add(l_oNewItem);

                    }


                }

            }


            return l_iPrintComCount;
        }

        //20210129 CCL+  利用算完的 公司 小計 算出--> 勞保 墊償 墊償小計 
        public int Fun_CalcLInsFundByCompanyResults(List<MERP_LHInsExcelItem> p_oLHInsExcelByComItems)
        {

            int l_iRowIndex = 0;
            string l_sPrevCompanyName = "";

            //20210125 CCL+ 計算勞保-個人+單位 小計
            decimal l_dTolLabPerComInsAmount = 0;
            //20210129 CCL+ 計算墊償和小計
            decimal l_dComLaborFundAmount = 0;
            decimal l_dTolLabComFundAmount = 0;

            int l_iPrintComCount = 0;

            if ((p_oLHInsExcelByComItems != null) &&
                (p_oLHInsExcelByComItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in p_oLHInsExcelByComItems)
                {
                    ++l_iRowIndex;

                    //更新原本存放代償的List 到 存放公司總計List                    
                    if (m_oLInsFundByComResults != null && m_oLInsFundByComResults.Count() > 0)
                    {
                        MERP_LHInsExcelItem l_oItem = null;
                        l_oItem = m_oLInsFundByComResults[l_iRowIndex-1];
                        //l_oItem = m_oLInsFundByComResults
                        //            .Where(m => m.m_sPlusInsCompany == Item.m_sPlusInsCompany).First();

                        //20210125 CCL+ 勞保-個人+單位 小計
                        l_dTolLabPerComInsAmount = Item.m_dTolLabPerComInsAmount;
                        //20210129 CCL+ 計算墊償和小計
                        l_dComLaborFundAmount = l_oItem.m_dComLaborFundAmount;//之前存的墊償
                        l_dTolLabComFundAmount = l_dTolLabPerComInsAmount + l_dComLaborFundAmount;
                        l_oItem.m_dTolLabComFundAmount = l_dTolLabComFundAmount;

                        //更新原本存放代償的List 到 存放公司總計List 
                        Item.m_dComLaborFundAmount = l_oItem.m_dComLaborFundAmount;
                        Item.m_dTolLabComFundAmount = l_oItem.m_dTolLabComFundAmount;
                        //所有墊償 值相加
                        m_dLaborInsFundTolAmounts += (double)Item.m_dComLaborFundAmount;
                        //所有墊償 小計 值相加  
                        m_dComTolLInsFundTolAmounts += (double)Item.m_dTolLabComFundAmount;

                        

                    }


                }
            }

            l_iPrintComCount = l_iRowIndex;

            return l_iPrintComCount;

        }


        //20210225 CCL+ 利用計算出的墊償, 墊償小計 算出各門市 墊償比例
        public int Fun_CalcLHInsFundPercentByComanyResults(List<MERP_LHInsExcelItem> p_oLHInsExcelByComItems)
        {
            int l_iRowIndex = 0;
            string l_sPrevCompanyName = "";
            string l_sPrevDepartName = "";

            //各店墊償比例 = 公司墊償 X (公司部門 [單+個]欄位值 /公司 [單+個]欄位值)
            //各店墊償比例小計 = 公司墊償小計 X (公司部門 [單+個]欄位值 /公司 [單+個]欄位值)
            decimal l_dComLaborFundShopPercent = 0; //各店墊償比例
            decimal l_dTolLabComFundShopPercent = 0; // 各店墊償比例小計

            decimal l_dTolLabInsComAmount = 0; //公司 [單+個]欄位值
            decimal l_dTolLabInsComDepAmount = 0; //公司部門 [單+個]欄位值
            decimal l_dTolLabInsCalcPercent = 0; //計算比例
            //公司墊償和小計
            decimal l_dComLaborFundAmount = 0; //公司墊償
            decimal l_dTolLabComFundAmount = 0; //公司墊償小計


            if ((p_oLHInsExcelByComItems != null) &&
                (p_oLHInsExcelByComItems.Count() > 0))
            {
                //各公司小計
                foreach (MERP_LHInsExcelItem Item in p_oLHInsExcelByComItems)
                {
                    ++l_iRowIndex;

                    //取出 公司 [單+個]欄位值
                    l_dTolLabInsComAmount = Item.m_dTolLabPerComInsAmount;
                    //取出 公司墊償
                    l_dComLaborFundAmount = Item.m_dComLaborFundAmount;
                    //取出 公司墊償小計
                    l_dTolLabComFundAmount = Item.m_dTolLabComFundAmount;


                    //找出該公司所屬各店 門市小計Data
                    var l_oFindData =
                        from DataVals in m_oLHInsExcelByComDepResults
                        where DataVals.m_sPlusInsCompany == Item.m_sPlusInsCompany
                        select DataVals;

                    List<MERP_LHInsExcelItem> l_oFindThisComAllShopsDataList = l_oFindData.ToList();
                    if(l_oFindThisComAllShopsDataList != null && l_oFindThisComAllShopsDataList.Count() > 0)
                    {
                        //從找出的門市小計中,找出 [單+個]欄位值
                        foreach(MERP_LHInsExcelItem ShopDataItem in l_oFindThisComAllShopsDataList)
                        {
                            

                            //取出 公司部門 [單+個]欄位值
                            l_dTolLabInsComDepAmount = ShopDataItem.m_dTolLabPerComInsAmount;

                            //計算比例
                            l_dTolLabInsCalcPercent = l_dTolLabInsComDepAmount / l_dTolLabInsComAmount;

                            //計算墊償比例
                            l_dComLaborFundShopPercent = l_dComLaborFundAmount * l_dTolLabInsCalcPercent;
                            //計算墊償小計比例
                            l_dTolLabComFundShopPercent = l_dTolLabComFundAmount * l_dTolLabInsCalcPercent;

                            //計算出比例,更新寫回到m_oLHInsExcelByComDepResults Data
                            //更新寫回 門市小計 墊償 值
                            ShopDataItem.m_dComLaborFundAmount = Math.Round(l_dComLaborFundShopPercent, MidpointRounding.AwayFromZero);
                            //更新寫回 門市小計 墊償小計 值
                            ShopDataItem.m_dTolLabComFundAmount = Math.Round(l_dTolLabComFundShopPercent, MidpointRounding.AwayFromZero);


                        }

                    }
                       

                }

            }




                    return 0;
        }


        /*
        //改成By 公司
        public int Fun_CalcLHInsByComResults(List<MERP_LHInsExcelItem> p_oLHInsResultItems)
        {
            int l_iRowIndex = 0;
            string l_sPrevCompanyName = "";

            decimal l_dComTolLaborODInsAmount = 0;
            decimal l_dPerLaborInsAmount = 0;
            decimal l_dComHealInsAmount = 0;
            decimal l_dPerHealInsAmount = 0;
            decimal l_dComRetireInsAmount = 0;
            decimal l_dTolComLHRInsAmount = 0;
            decimal l_dTolPerLHRInsAmount = 0;
            decimal l_dTolAllLHRInsAmount = 0;
           
            int l_iPrintCompanyCount = 0;

            if ((p_oLHInsResultItems != null) &&
                (p_oLHInsResultItems.Count() > 0) )
            {
                foreach(MERP_LHInsExcelItem Item in p_oLHInsResultItems)
                {
                    ++l_iRowIndex;

                    if (l_sPrevCompanyName == "")
                    {
                        //第一筆
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        //更新公司
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;

                    }
                    else if (Item.m_sPlusInsCompany == l_sPrevCompanyName)
                    {
                        //累加
                        l_dComTolLaborODInsAmount += Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount += Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount += Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount += Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount += Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount += Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount += Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount += Item.m_dTolAllLHRInsAmount;

                        
                    }
                    else if (Item.m_sPlusInsCompany != l_sPrevCompanyName)
                    {
                        

                        ++l_iPrintCompanyCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;

                        m_oLHInsExcelByComResults.Add(l_oNewItem);

                        ////////////////////////////////////////////////////////////////
                        //更新公司
                        l_sPrevCompanyName = Item.m_sPlusInsCompany;
                        //設為這一輪的值
                        l_dComTolLaborODInsAmount = Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount = Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount = Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount = Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount = Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount = Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount = Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount = Item.m_dTolAllLHRInsAmount;
                        

                    }


                    if( l_iRowIndex == p_oLHInsResultItems.Count() )
                    {
                        //最後一筆時也要執行這裡
                        ++l_iPrintCompanyCount;
                        /////////////////////////////////////////////////////////////
                        //儲存上一間店的總計
                        MERP_LHInsExcelItem l_oNewItem = new MERP_LHInsExcelItem();
                        l_oNewItem.m_sPlusInsCompany = l_sPrevCompanyName;

                        l_oNewItem.m_dComTolLaborODInsAmount = l_dComTolLaborODInsAmount;
                        l_oNewItem.m_dPerLaborInsAmount = l_dPerLaborInsAmount;
                        l_oNewItem.m_dComHealInsAmount = l_dComHealInsAmount;
                        l_oNewItem.m_dPerHealInsAmount = l_dPerHealInsAmount;
                        l_oNewItem.m_dComRetireInsAmount = l_dComRetireInsAmount;
                        l_oNewItem.m_dTolComLHRInsAmount = l_dTolComLHRInsAmount;
                        l_oNewItem.m_dTolPerLHRInsAmount = l_dTolPerLHRInsAmount;
                        l_oNewItem.m_dTolAllLHRInsAmount = l_dTolAllLHRInsAmount;

                        m_oLHInsExcelByComResults.Add(l_oNewItem);

                    }

                }
            }

            

            return l_iPrintCompanyCount;
        }
        */

        //計算合計 (L Labor + H Heal + R Retire = T TolLHR)
        /*
        public decimal Fun_CalcTolLHRInsResult(MERP_LHInsExcelItem p_oLHRInsItem)
        {
            //L單位+L個人+H單位+H個人+R單位 = Total
            p_oLHRInsItem.m_dTolLHRInsAmount = p_oLHRInsItem.m_dComTolLaborODInsAmount +
                                                p_oLHRInsItem.m_dPerLaborInsAmount +
                                                p_oLHRInsItem.m_dComHealInsAmount +
                                                p_oLHRInsItem.m_dPerHealInsAmount + 
                                                p_oLHRInsItem.m_dComRetireInsAmount;

           

            return p_oLHRInsItem.m_dTolLHRInsAmount;
        }
        */

        //改成合計也分單位,個人
        //計算合計-個人 (L PerLabor個人 + H PerHeal個人  = T PerTolLHR個人)
        public decimal Fun_CalcTolPerLHRInsResult(MERP_LHInsExcelItem p_oLHRInsItem)
        {
            decimal l_dRtnVal = 0;

            l_dRtnVal = p_oLHRInsItem.m_dPerLaborInsAmount +
                              p_oLHRInsItem.m_dPerHealInsAmount;

            return l_dRtnVal;
        }

        //計算合計-單位 (L ComLabor單位 + H ComHeal單位 + R Retire單位  = T ComTolLHR單位)
        public decimal Fun_CalcTolComLHRInsResult(MERP_LHInsExcelItem p_oLHRInsItem)
        {
            decimal l_dRtnVal = 0;
            //PS:要用含職災的m_dComTolLaborODInsAmount
            l_dRtnVal = p_oLHRInsItem.m_dComTolLaborODInsAmount +
                               p_oLHRInsItem.m_dComHealInsAmount +
                               p_oLHRInsItem.m_dComRetireInsAmount;

            return l_dRtnVal;
        }

        //計算合計-所有 = 合計-單位 + 合計-個人
        public decimal Fun_CalcTolAllLHRInsResult(MERP_LHInsExcelItem p_oLHRInsItem)
        {
            decimal l_dRtnVal = 0;

            l_dRtnVal = p_oLHRInsItem.m_dTolPerLHRInsAmount +
                               p_oLHRInsItem.m_dTolComLHRInsAmount;

            return l_dRtnVal;
            //m_dAllLHRInsTolAmounts = m_dPerLHRInsTolAmounts + m_dComLHRInsTolAmounts;

            //return m_dAllLHRInsTolAmounts;
        }


        //計算勞保-個人 ----------------------------------------------------------------------------
        /* //20210127 CCL- 修正算法:要拆開來各自四捨五入後再相加, 
         * //總薪資 X 普通事故保險費率10% X 負擔比例(20%,70%) ==> 四捨五入 + 
         * //總薪資 X 就業保險費率1% X 負擔比例(20%,70%) ==> 四捨五入 +
         * //總薪資 X 職災保險費率0.17% X 負擔比例(100%) ==> 四捨五入
        public decimal Fun_PerLaborInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 個人保險費率(PersonalInsRate)/100 X 被保險人負擔比例(LaborBurdenRatio)/100 

            //如果備註Type: 0 --> 一般 1; --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10%      

            decimal l_dRtnVal = 0;
          
            if (m_oLaborInsSet != null)
            {
                decimal l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.PersonalInsRate) / 100;
                decimal l_dLaborBurdenRatio = Convert.ToDecimal(m_oLaborInsSet.LaborBurdenRatio) / 100;

                switch(p_iLHInsType)
                {
                    
                    case 2:
                        l_dRtnVal = 0;
                        break;
                    case 3:
                        l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                        l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dLaborBurdenRatio;
                        break;

                    case 1:
                    default:
                        l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dLaborBurdenRatio;
                        break;
                }

                
                //四捨五入 -> 整數
                l_dRtnVal = Math.Round(l_dRtnVal);
            }

            return l_dRtnVal;
        }
        */


        //20210127 CCL Mod 勞保: 個人; (修正算法:要拆開來各自四捨五入後再相加),  
        //總薪資 X 普通事故保險費率10% X 負擔比例(20%,70%) ==> 四捨五入 + 
        //總薪資 X 就業保險費率1% X 負擔比例(20%,70%) ==> 四捨五入 +
        //總薪資 X 職災保險費率0.17% X 負擔比例(100%) ==> 四捨五入
        public decimal Fun_PerLaborInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 個人保險費率(PersonalInsRate)/100 X 被保險人負擔比例(LaborBurdenRatio)/100 

            //如果備註Type: 0 --> 一般 1; --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10%      
            

            decimal l_dRtnVal = 0;
            

            if (m_oLaborInsSet != null)
            {
                //decimal l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.PersonalInsRate) / 100;
                decimal l_dOrdAccidentInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                decimal l_dEmployInsRate = Convert.ToDecimal(m_oLaborInsSet.EmployInsRate) / 100;
                decimal l_dLaborBurdenRatio = Convert.ToDecimal(m_oLaborInsSet.LaborBurdenRatio) / 100;

                decimal l_dLaborInsAccVal = 0;
                decimal l_dLaborInsEmpVal = 0;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;                       
                        break;
                    case 3:
                        //l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                        l_dRtnVal = p_dLaborInsVal * l_dOrdAccidentInsRate * l_dLaborBurdenRatio;
                        //四捨五入 -> 整數
                        l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
                        break;

                    case 1:
                    default:
                        //l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dLaborBurdenRatio;
                        l_dLaborInsAccVal = p_dLaborInsVal * l_dOrdAccidentInsRate * l_dLaborBurdenRatio;
                        l_dRtnVal = Math.Round(l_dLaborInsAccVal, MidpointRounding.AwayFromZero);
                        l_dLaborInsEmpVal = p_dLaborInsVal * l_dEmployInsRate * l_dLaborBurdenRatio;
                        l_dRtnVal += Math.Round(l_dLaborInsEmpVal, MidpointRounding.AwayFromZero);
                        break;
                }

               
            }

            return l_dRtnVal;
        }

        public decimal Fun_CalcPerLaborInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {

            decimal l_dLaborInsOrgVal = 0;
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dLaborInsDatesSalary.ToString()))
            {
                l_dLaborInsOrgVal = p_oLHInsItem.m_dLaborInsDatesSalary; //Index 4
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            //return Fun_PerLaborInsFormula(l_dLaborInsOrgVal);
            return Fun_PerLaborInsFormula(l_dLaborInsOrgVal, l_iLHInsType);

        }

        /* 20210120 CCL-
        public decimal Fun_CalcPerLaborInsResult(DataRow p_oRowItemData)
        {
            
            decimal l_dLaborInsOrgVal = 0;
            //計算
            if (!string.IsNullOrEmpty(p_oRowItemData["LaborIns"].ToString()))
            {
                l_dLaborInsOrgVal = Convert.ToDecimal(p_oRowItemData["LaborIns"]); //Index 4
            }

            return Fun_PerLaborInsFormula(l_dLaborInsOrgVal);

        }
        */
        // --------------------------------------------------------------------------------------


        //計算勞保-單位(不含職災) ---------------------------------------------------------------
        /* 20210127 CCL-
        public decimal Fun_ComLaborInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 個人保險費率(PersonalInsRate)/100 X 投保單位負擔比例(ComBurdenRatio)/100

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10% 

            decimal l_dRtnVal = 0;
           
            if (m_oLaborInsSet != null)
            {
                decimal l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.PersonalInsRate) / 100;
                decimal l_dComBurdenRatio = Convert.ToDecimal(m_oLaborInsSet.ComBurdenRatio) / 100;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;
                    case 3:
                        l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                        l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio;
                        break;

                    case 1:
                    default:
                        l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio;
                        break;
                }

                //l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio ;
                //四捨五入 -> 整數
                //20210125 CCL- 必須與職災相加後才能四捨五入,不然會有誤差 l_dRtnVal = Math.Round(l_dRtnVal);
            }

            return l_dRtnVal;
        }
        */

        //20210127 CCL Mod 勞保: 單位; (修正算法:要拆開來各自四捨五入後再相加),  
        //總薪資 X 普通事故保險費率10% X 負擔比例(20%,70%) ==> 四捨五入 + 
        //總薪資 X 就業保險費率1% X 負擔比例(20%,70%) ==> 四捨五入 +
        //總薪資 X 職災保險費率0.17% X 負擔比例(100%) ==> 四捨五入
        public decimal Fun_ComLaborInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 個人保險費率(PersonalInsRate)/100 X 投保單位負擔比例(ComBurdenRatio)/100

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10% 

            decimal l_dRtnVal = 0;

            if (m_oLaborInsSet != null)
            {
                //decimal l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.PersonalInsRate) / 100;
                decimal l_dOrdAccidentInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                decimal l_dEmployInsRate = Convert.ToDecimal(m_oLaborInsSet.EmployInsRate) / 100;
                decimal l_dComBurdenRatio = Convert.ToDecimal(m_oLaborInsSet.ComBurdenRatio) / 100;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;
                    case 3:
                        //l_dPersonalInsRate = Convert.ToDecimal(m_oLaborInsSet.OrdAccidentInsRate) / 100;
                        //l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio;
                        l_dRtnVal = p_dLaborInsVal * l_dOrdAccidentInsRate * l_dComBurdenRatio;
                        //四捨五入 -> 整數
                        l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
                        break;

                    case 1:
                    default:
                        //l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio;
                        decimal l_dLaborInsAccVal = p_dLaborInsVal * l_dOrdAccidentInsRate * l_dComBurdenRatio;
                        l_dRtnVal = Math.Round(l_dLaborInsAccVal, MidpointRounding.AwayFromZero);
                        //l_dRtnVal = Math.Round(l_dLaborInsAccVal, MidpointRounding.ToEven);
                        decimal l_dLaborInsEmpVal = p_dLaborInsVal * l_dEmployInsRate * l_dComBurdenRatio;
                        l_dRtnVal += Math.Round(l_dLaborInsEmpVal, MidpointRounding.AwayFromZero);
                        break;
                }

                //l_dRtnVal = p_dLaborInsVal * l_dPersonalInsRate * l_dComBurdenRatio ;
                
            }

            return l_dRtnVal;
        }

        public decimal Fun_CalcComLaborInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {

            decimal l_dLaborInsOrgVal = 0;
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dLaborInsDatesSalary.ToString()))
            {
                l_dLaborInsOrgVal = p_oLHInsItem.m_dLaborInsDatesSalary; //Index 4
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            return Fun_ComLaborInsFormula(l_dLaborInsOrgVal, l_iLHInsType);
        }

        /* 20210120 CCL-
        public decimal Fun_CalcComLaborInsResult(DataRow p_oRowItemData)
        {
            
            decimal l_dLaborInsOrgVal = 0;
            //計算
            if (!string.IsNullOrEmpty(p_oRowItemData["LaborIns"].ToString()))
            {
                l_dLaborInsOrgVal = Convert.ToDecimal(p_oRowItemData["LaborIns"]); //Index 4
            }

            return Fun_ComLaborInsFormula(l_dLaborInsOrgVal);
        }
        */
        // -------------------------------------------------------------------------------------

        //(職災只屬單位)
        //計算職災-只有單位 --------------------------------------------------------------------
        /* 20210127 CCl-
        public decimal Fun_OccuDisaInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 職業災害保險費率(OccuDisaInsRate)/100

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10% 

            decimal l_dRtnVal = 0;
                       
            if (m_oLaborInsSet != null)
            {
                decimal l_dOccuDisaInsRate = Convert.ToDecimal(m_oLaborInsSet.OccuDisaInsRate) / 100;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;

                    case 3:                      
                    case 1:
                    default:
                        l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate;
                        break;
                }


                //l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate ;
                //四捨五入 -> 整數
                //20210125 CCL- 必須與勞保-單位相加後才能四捨五入,不然會有誤差 l_dRtnVal = Math.Round(l_dRtnVal);
            }

            return l_dRtnVal;
        }
        */

        //20210127 CCL Mod 勞保: 單位; (修正算法:要拆開來各自四捨五入後再相加),  
        //總薪資 X 普通事故保險費率10% X 負擔比例(20%,70%) ==> 四捨五入 + 
        //總薪資 X 就業保險費率1% X 負擔比例(20%,70%) ==> 四捨五入 +
        //總薪資 X 職災保險費率0.17% X 負擔比例(100%) ==> 四捨五入
        public decimal Fun_OccuDisaInsFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 職業災害保險費率(OccuDisaInsRate)/100

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10% 

            decimal l_dRtnVal = 0;

            if (m_oLaborInsSet != null)
            {
                decimal l_dOccuDisaInsRate = Convert.ToDecimal(m_oLaborInsSet.OccuDisaInsRate) / 100;
                decimal l_dOccuDisComBurdenRatio = Convert.ToDecimal(m_oLaborInsSet.OccuDisComBurdenRatio) / 100;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;

                    case 3:
                    case 1:
                    default:
                        //l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate;
                        l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate * l_dOccuDisComBurdenRatio;
                        //四捨五入 -> 整數
                        l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
                        break;
                }


                //l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate ;
               
            }

            return l_dRtnVal;
        }


        public decimal Fun_CalcOccuDisaInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {

            decimal l_dLaborInsOrgVal = 0;
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dLaborInsDatesSalary.ToString()))
            {
                l_dLaborInsOrgVal = p_oLHInsItem.m_dLaborInsDatesSalary; //Index 4
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            //return Fun_OccuDisaInsFormula(l_dLaborInsOrgVal);
            return Fun_OccuDisaInsFormula(l_dLaborInsOrgVal, l_iLHInsType);
        }

        /* 20210120 CCL-
        public decimal Fun_CalcOccuDisaInsResult(DataRow p_oRowItemData)
        {
            
            decimal l_dLaborInsOrgVal = 0;
            //計算
            if (!string.IsNullOrEmpty(p_oRowItemData["LaborIns"].ToString()))
            {
                l_dLaborInsOrgVal = Convert.ToDecimal(p_oRowItemData["LaborIns"]); //Index 4
            }

            return Fun_OccuDisaInsFormula(l_dLaborInsOrgVal);
        }
        */
        //-------------------------------------------------------------------------------------

        //(勞退只屬單位)
        //計算勞退-只有單位 ----------------------------------------------------------------------------
        public decimal Fun_LaborRetireFormula(decimal p_dLaborInsVal, int p_iLHInsType)
        {
            //公式: 勞保 X 勞退費率(LaborRetireRate)/100

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10% 

            decimal l_dRtnVal = 0;

            if (m_oLaborInsSet != null)
            {
                decimal l_dOccuDisaInsRate = Convert.ToDecimal(m_oLaborInsSet.LaborRetireRate) / 100;

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;
                    case 3:
                        l_dRtnVal = 0;
                        break;

                    case 1:
                    default:
                        l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate;
                        break;
                }

                //l_dRtnVal = p_dLaborInsVal * l_dOccuDisaInsRate;
                //四捨五入 -> 整數
                l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
            }

            return l_dRtnVal;
        }

        public decimal Fun_CalcLaborRetireInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {

            decimal l_dLaborInsOrgVal = 0;
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dLaborInsDatesSalary.ToString()))
            {
                l_dLaborInsOrgVal = p_oLHInsItem.m_dLaborInsDatesSalary; //Index 4
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            //return Fun_LaborRetireFormula(l_dLaborInsOrgVal);
            return Fun_LaborRetireFormula(l_dLaborInsOrgVal, l_iLHInsType);
        }

        /* 20210120 CCL-
       public decimal Fun_CalcLaborRetireInsResult(DataRow p_oRowItemData)
       {

           decimal l_dLaborInsOrgVal = 0;
           //計算
           if (!string.IsNullOrEmpty(p_oRowItemData["LaborIns"].ToString()))
           {
               l_dLaborInsOrgVal = Convert.ToDecimal(p_oRowItemData["LaborIns"]); //Index 4
           }

           return Fun_LaborRetireFormula(l_dLaborInsOrgVal);
       }
       */
        //-------------------------------------------------------------------------------------

        /// <summary>
        /// ////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>

        //計算健保-個人 ----------------------------------------------------------------------------
        public decimal Fun_PerHealFormula(decimal p_dHealInsVal, decimal p_dDependentsNum, int p_iLHInsType)
        {
            //公式: 健保 X 健康保險費率(Heal_Rate)/100 X 個人負擔比例(Heal_LaborInsBurdenRatio)/100 
            //結果 = 結果 + (結果X眷屬)

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10%    

            decimal l_dRtnVal = 0;

            if (m_oLaborInsSet != null)
            {
                decimal l_dHeal_Rate = Convert.ToDecimal(m_oHealInsSet.Heal_Rate) / 100;
                decimal l_dHeal_LInsBurdenRatio = Convert.ToDecimal(m_oHealInsSet.Heal_LaborInsBurdenRatio) / 100;
                //20210128 CCL Mod 必須將整數先轉成decimal, 再除,不然值會直接是 整數 = 整數/整數 = 0(0.5的5被自動去掉) 
                decimal l_dLowPer_Rate = (decimal)50/ (decimal)100; //健保個人低收

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;
                    case 1:
                        l_dRtnVal = p_dHealInsVal * l_dHeal_Rate * l_dHeal_LInsBurdenRatio;
                        l_dRtnVal = l_dRtnVal * l_dLowPer_Rate;
                        break;

                    case 3:
                    default:
                        l_dRtnVal = p_dHealInsVal * l_dHeal_Rate * l_dHeal_LInsBurdenRatio;
                        break;
                }


                //l_dRtnVal = p_dHealInsVal * l_dHeal_Rate * l_dHeal_LInsBurdenRatio;
                //四捨五入 -> 整數
                l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
            }

            if(p_dDependentsNum > 0)
            {
                l_dRtnVal = l_dRtnVal + (l_dRtnVal * p_dDependentsNum);
            }

            return l_dRtnVal;
        }

        public decimal Fun_CalcPerHealInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {
            decimal l_dHealInsOrgVal = 0;
            decimal l_dDependentsNum = 0; //20210121 CCL+ 眷屬
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dHealInsDatesSalary.ToString()))
            {
                l_dHealInsOrgVal = p_oLHInsItem.m_dHealInsDatesSalary; //Index 4
                l_dDependentsNum = p_oLHInsItem.m_DependentsNum;
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            //20210121 CCL* return Fun_PerHealFormula(l_dHealInsOrgVal);
            //return Fun_PerHealFormula(l_dHealInsOrgVal, l_dDependentsNum);
            return Fun_PerHealFormula(l_dHealInsOrgVal, l_dDependentsNum, l_iLHInsType);
        }

        /* 20210120 CCL-
        public decimal Fun_CalcPerHealInsResult(DataRow p_oRowItemData)
        {
            decimal l_dHealInsOrgVal = 0;
            //計算
            if (!string.IsNullOrEmpty(p_oRowItemData["HealthIns"].ToString()))
            {
                l_dHealInsOrgVal = Convert.ToDecimal(p_oRowItemData["HealthIns"]); //Index 4
            }            

            return Fun_PerHealFormula(l_dHealInsOrgVal);
        }
        */
        //-------------------------------------------------------------------------------------


        //計算健保-單位
        public decimal Fun_ComHealFormula(decimal p_dHealInsVal, int p_iLHInsType)
        {
            //20210125 CCL+, 修正健保單位要乘以平均眷口數費率 1.58
            //公式: 健保 X 健康保險費率(Heal_Rate)/100 X 公司負擔比例(Heal_ComInsBurdenRatio)/100 X 平均眷口數費率(Heal_AverhouseholdsNum)

            //如果備註Type: 1 --> 健保個人低收要X0.5(一半); 2 --> 育嬰個人單位勞健保都是0; 
            //3 --> 不適用一般勞工(外國人)改乘以[普通事故保險費率OrdAcciInsRate] 10%    

            decimal l_dRtnVal = 0;
            int l_iLHInsType = 0;

            if (m_oLaborInsSet != null)
            {
                decimal l_dHeal_Rate = Convert.ToDecimal(m_oHealInsSet.Heal_Rate) / 100;
                decimal l_dHeal_ComInsBurdenRatio = Convert.ToDecimal(m_oHealInsSet.Heal_ComInsBurdenRatio) / 100;
                decimal l_dHeal_AverhouseholdsNum = Convert.ToDecimal(m_oHealInsSet.Heal_AverhouseholdsNum); //20210125 CCL+

                switch (p_iLHInsType)
                {

                    case 2:
                        l_dRtnVal = 0;
                        break;

                    case 3:
                    case 1:
                    default:
                        l_dRtnVal = p_dHealInsVal * l_dHeal_Rate * l_dHeal_ComInsBurdenRatio * l_dHeal_AverhouseholdsNum;
                        break;
                }


                //l_dRtnVal = p_dHealInsVal * l_dHeal_Rate * l_dHeal_ComInsBurdenRatio * l_dHeal_AverhouseholdsNum;
                //四捨五入 -> 整數
                l_dRtnVal = Math.Round(l_dRtnVal, MidpointRounding.AwayFromZero);
            }



            return l_dRtnVal;
        }

        public decimal Fun_CalcComHealInsResult(MERP_LHInsExcelItem p_oLHInsItem)
        {

            decimal l_dHealInsOrgVal = 0;
            int l_iLHInsType = 0;

            //計算
            if (!string.IsNullOrEmpty(p_oLHInsItem.m_dHealInsDatesSalary.ToString()))
            {
                l_dHealInsOrgVal = p_oLHInsItem.m_dHealInsDatesSalary; //Index 4
            }

            l_iLHInsType = p_oLHInsItem.m_LHInsType;

            //return Fun_ComHealFormula(l_dHealInsOrgVal);
            return Fun_ComHealFormula(l_dHealInsOrgVal, l_iLHInsType);
        }

        /* 20210120 CCL-
        public decimal Fun_CalcComHealInsResult(DataRow p_oRowItemData)
        {
            
            decimal l_dHealInsOrgVal = 0;
            //計算
            if (!string.IsNullOrEmpty(p_oRowItemData["HealthIns"].ToString()))
            {
                l_dHealInsOrgVal = Convert.ToDecimal(p_oRowItemData["HealthIns"]); //Index 4
            }
            return Fun_ComHealFormula(l_dHealInsOrgVal);
        }
        */


        /// <summary>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>

        //Constructor
        public MERP_LHInsExcelExpore()
        {
            m_oLHInsExcelResultItems = new List<MERP_LHInsExcelItem>();

            m_dTolLaborInsVal = 0;
            m_dTolLaborInsValQuota = 0;

            //所有勞保-單位 值相加
            m_dComLaborInsTolAmounts = 0;

            //所有勞保-個人 值相加
            m_dPerLaborInsTolAmounts = 0;

            //所有健保-單位 值相加
            m_dComHealInsTolAmounts = 0;

            //所有健保-個人 值相加
            m_dPerHealInsTolAmounts = 0;

            //所有勞退-單位 值相加
            m_dComRetireInsTolAmounts = 0;

            //所有LHR合計 值相加
            m_dAllLHRInsTolAmounts = 0;

            //20210119 CCL+
            m_oLHInsExcelByComResults = new List<MERP_LHInsExcelItem>();

            m_oLHInsExcelByComDepResults = new List<MERP_LHInsExcelItem>();

            //20210129 CCL+ //取出勞保設定
            m_oLInsSetMapComSetDBService = new MERP_FA_LaborInsSetMapComSetDBService();

            //20210129 CCL+ 放各公司墊償結果
            m_oLInsFundByComResults = new List<MERP_LHInsExcelItem>();
        }

        //
    }
}
 
 
 
 
 