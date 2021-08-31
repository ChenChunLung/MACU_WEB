using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using MACU_WEB.Services;

namespace MACU_WEB.BIServices
{
    public class MERP_AccInfoExpore
    {
       

        //Base Attr (內放Combine過 AccNo,DtlAccNo,CAmount,DAmount; 其中金額已累加過)
        public List<TmpExcelItem> m_oBaseAttrItems { get; set; }
        //分群Base項目
        public List<TmpExcelItem> m_oAttrItemsGID4 { get; set; }
        public List<TmpExcelItem> m_oAttrItemsGID5 { get; set; }
        public List<TmpExcelItem> m_oAttrItemsGID6 { get; set; }
        public List<TmpExcelItem> m_oAttrItemsGID7 { get; set; }
        public List<TmpExcelItem> m_oAttrItemsGID8 { get; set; }
        //剩下排除的
        public List<TmpExcelItem> m_oAttrItemsOthers { get; set; }
        //Beginning of raw materials 原料期初
        public List<TmpExcelItem> m_oAttrBeginRawMaterial { get; set; }
        //End of raw material period 減 原料期末
        public List<TmpExcelItem> m_oAttrEndRawMaterial { get; set; }

        //找出來的對應的AccountInfo Map項目
        //public List<AccountInfo> m_oMapedAccInfos { get; set; }
        public List<AccountInfo> m_oMapedAccInfosGID4 { get; set; }
        public List<AccountInfo> m_oMapedAccInfosGID5 { get; set; }
        public List<AccountInfo> m_oMapedAccInfosGID6 { get; set; }
        public List<AccountInfo> m_oMapedAccInfosGID7 { get; set; }
        public List<AccountInfo> m_oMapedAccInfosGID8 { get; set; }

        

        //Ext Attr
        //public int m_GroupId { get; set; } //GId
        public string m_sShopId { get; set; } //ShopId
        public string m_sShopName { get; set; } //ShopName

        public int m_ExcludeCount { get; set; } //非在AccInfo Map Table內的項目數

        public int m_iSubp1192Count { get; set; } //原料期初, 原料期末(都叫原料存貨)



        //全域變數
        //[1].
        public double m_dOperaIncome { get; set; } //營業收入 = (GId=4 相加)
        //[2].
        public double m_dTolOperaCosts { get; set; } //營業總成本 = m_dTolCostOfCashExpend + [原料期初]
                                                            // - [減 原料期末]
        //[3].
        public double m_dActualTolCosts { get; set; } //實際用量總成本 = 營業總成本 = (GId=5 相加)
        //[4].
        public double m_dOperaMargin { get; set; } //營業毛利 = 營業收入 - 實際用量總成本
        //[5].
        public double m_dTopOperaExpense { get; set; } //營業費用 = (GId=6 相加)
        //[6].
        public double m_dBtmOperaExpense { get; set; } //營業費用 = (GId=6 相加)
        //[7].
        public double m_dBussInterest { get; set; } //營業利益 = 營業毛利 - 營業費用
        //[8].
        public double m_dNonOperaIncome { get; set; } //非營業收入 = (GId=7 相加)
        //[9].
        public double m_dNonOperaExpense { get; set; } //非營業支出 = (GId=8 相加)
        //[10].
        public double m_dConsuCurrentProfitLoss { get; set; } //實際用量本期損益 = [1]-[3]-[6]+[8]-[9]
        //[11].
        public double m_dTolCostOfCashExpend { get; set; } //現金支出總成本 = (PrintOrder:14~35相加)
        //[12].
        public double m_dCashExpendForCurrPeriod { get; set; } //現金支出本期損益 = [1]-[11]-[6]+[8]-[9]


        //分群 已加總每天的 BaseItem
        public void GroupingBaseItems(MERP_AccountInfoDBService p_oAccInfoDBService)
        {
            if(m_oBaseAttrItems.Count() > 0)
            {
                AccountInfo l_pTmpAccInfoItem = null;
                m_ExcludeCount = 0;

                m_iSubp1192Count = 0;

                //找出該AccNO,DtlAccNo 的AccInfoTB Map Item
                foreach (TmpExcelItem item in m_oBaseAttrItems)
                {
                    l_pTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo(item.m_ComAccountNo,
                                                                           item.m_ComDetailAccNo);

                    //要小心 有很多科目名稱的ID都是重複的1192,比如[原料存貨]也是 1192-"" 
                    /*
                    if (item.m_ComAccountNo == "1192")
                    {
                        if (!( item.m_ComAccName == "原料期初" || item.m_ComAccName == "減 原料期末" ) )
                        {
                            //在AccInfo Table Map內的 重複的1192 不顯示不計算
                            m_oAttrItemsOthers.Add(item);
                            m_ExcludeCount++;
                            continue;
                        }
                       

                    }
                    */

                    if (l_pTmpAccInfoItem != null)
                    {
                        switch(l_pTmpAccInfoItem.GroupID)
                        {
                            case 4:
                                //加入分群 4
                                item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                m_oAttrItemsGID4.Add(item);
                                m_oMapedAccInfosGID4.Add(l_pTmpAccInfoItem);
                                break;
                            case 5:
                                //加入分群 5
                                //利用傳票號碼來區分 原料期初, 原料期末(都叫原料存貨)
                                if (item.m_ComAccountNo == "1192")
                                {
                                    //1192另外存
                                    if (item.m_ComAccName == "原料存貨")
                                    {
                                        //因為匯入的DataBase是按日期排的,所以傳票號碼在前的一定是第一個
                                        //在AccInfo Table Map內的 1192 利用傳票號碼另外存
                                        ++m_iSubp1192Count;
                                        if (m_iSubp1192Count == 1)
                                        {
                                            //l_dSubpNo = Convert.ToDouble(item.m_ComSubpNo);
                                            item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                            m_oAttrBeginRawMaterial.Add( item); //期初

                                        } else
                                        {
                                            item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                            m_oAttrEndRawMaterial.Add(item); //期末
                                        }

                                        //m_oAttrItemsOthers.Add(item);
                                        m_ExcludeCount++;
                                        continue;
                                    }


                                } else
                                {
                                    item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                    m_oAttrItemsGID5.Add(item);
                                    m_oMapedAccInfosGID5.Add(l_pTmpAccInfoItem);
                                }
                               
                                break;
                            case 6:
                                //加入分群 6
                                item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                m_oAttrItemsGID6.Add(item);
                                m_oMapedAccInfosGID6.Add(l_pTmpAccInfoItem);
                                break;
                            case 7:
                                //加入分群 7
                                item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                m_oAttrItemsGID7.Add(item);
                                m_oMapedAccInfosGID7.Add(l_pTmpAccInfoItem);
                                break;
                            case 8:
                                //加入分群 8
                                item.m_PrintOrder = (int)l_pTmpAccInfoItem.PrintOrder; //20201229 CCL+ ReOrder By PrintOrder
                                m_oAttrItemsGID8.Add(item);
                                m_oMapedAccInfosGID8.Add(l_pTmpAccInfoItem);
                                break;
                            default:
                                
                                break;
                        }

                    } else
                    {
                        //不在AccInfo Table Map內的不顯示不計算
                        m_oAttrItemsOthers.Add(item);
                        m_ExcludeCount++;
                    }

                }
            }
        }



        //20201229 CCL+ ReOrder By PrintOrder
        public void ReOrderByPrintOrder()
        {
            if(m_oAttrItemsGID4.Count() > 0)
            {
                //取得新Order的List
                m_oAttrItemsGID4 = m_oAttrItemsGID4.OrderBy(m => m.m_PrintOrder).ToList();
                //20201231 CCL+
                m_oMapedAccInfosGID4 = m_oMapedAccInfosGID4.OrderBy(m => m.PrintOrder).ToList();
            }
            if (m_oAttrItemsGID5.Count() > 0)
            {
                //取得新Order的List
                m_oAttrItemsGID5 = m_oAttrItemsGID5.OrderBy(m => m.m_PrintOrder).ToList();
                //20201231 CCL+
                m_oMapedAccInfosGID5 = m_oMapedAccInfosGID5.OrderBy(m => m.PrintOrder).ToList();
            }
            if (m_oAttrItemsGID6.Count() > 0)
            {
                //取得新Order的List
                m_oAttrItemsGID6 = m_oAttrItemsGID6.OrderBy(m => m.m_PrintOrder).ToList();
                //20201231 CCL+
                m_oMapedAccInfosGID6 = m_oMapedAccInfosGID6.OrderBy(m => m.PrintOrder).ToList();
            }
            if (m_oAttrItemsGID7.Count() > 0)
            {
                //取得新Order的List
                m_oAttrItemsGID7 = m_oAttrItemsGID7.OrderBy(m => m.m_PrintOrder).ToList();
                //20201231 CCL+
                m_oMapedAccInfosGID7 = m_oMapedAccInfosGID7.OrderBy(m => m.PrintOrder).ToList();
            }
            if (m_oAttrItemsGID8.Count() > 0)
            {
                //取得新Order的List
                m_oAttrItemsGID8 = m_oAttrItemsGID8.OrderBy(m => m.m_PrintOrder).ToList();
                //20201231 CCL+
                m_oMapedAccInfosGID8 = m_oMapedAccInfosGID8.OrderBy(m => m.PrintOrder).ToList();
            }
        }

        public double Fun_CommonExpression(string p_sCountFlag, Double p_dOperaVal, TmpExcelItem p_oItem)
        {
            //20201228 CCL+ ComDAmount,ComCAmount計算完的值放ComAmount內
            p_oItem.m_ComAmount = "0"; //Init
            int l_iDAmount = 0, l_iCAmount = 0;

            Double l_RtnVal = p_dOperaVal;
            if(p_oItem.m_ComAccountNo.Substring(0,1) == "5")
            {
                int l_ixx = 0; //Debug
            }

            if (p_sCountFlag == "D")
            {
                //+借方 D
                if (p_oItem.m_ComDAmount == null)
                { p_oItem.m_ComDAmount = "0"; }
                //20201228 CCL- l_RtnVal += Convert.ToInt32(p_oItem.m_ComDAmount);
                l_iDAmount = Convert.ToInt32(p_oItem.m_ComDAmount);
                l_RtnVal += l_iDAmount;

                //-貸方 C
                if (p_oItem.m_ComCAmount == null)
                { p_oItem.m_ComCAmount = "0"; }
                //20201228 CCL- l_RtnVal -= Convert.ToInt32(p_oItem.m_ComCAmount);
                l_iCAmount = Convert.ToInt32(p_oItem.m_ComCAmount);
                l_RtnVal -= l_iCAmount;

                //20201228 CCL+ ComDAmount,ComCAmount計算完的值放ComAmount內
                p_oItem.m_ComAmount = (l_iDAmount - l_iCAmount).ToString();

            }
            else if (p_sCountFlag == "C")
            {
                //+貸方 C
                if (p_oItem.m_ComCAmount == null)
                { p_oItem.m_ComCAmount = "0"; }
                //20201228 CCL- l_RtnVal += Convert.ToInt32(p_oItem.m_ComCAmount);
                l_iCAmount = Convert.ToInt32(p_oItem.m_ComCAmount);
                l_RtnVal += l_iCAmount;

                //-借方 D
                if (p_oItem.m_ComDAmount == null)
                { p_oItem.m_ComDAmount = "0"; }
                //20201228 CCL- l_RtnVal -= Convert.ToInt32(p_oItem.m_ComDAmount);
                l_iDAmount = Convert.ToInt32(p_oItem.m_ComDAmount);
                l_RtnVal -= l_iDAmount;

                //20201228 CCL+ ComDAmount,ComCAmount計算完的值放ComAmount內
                p_oItem.m_ComAmount = (l_iCAmount - l_iDAmount).ToString();
            }

           

            return l_RtnVal;
        }

        //m_oBaseAttrItems全部塞好後開始下列計算和刪除計算過多餘的AttrItems
        public double Calc_GID4_OperaIncome()
        {
            //計算 營業收入 = (GId=4 相加)
            m_dOperaIncome = 0; //Init
           

            if ((m_oAttrItemsGID4.Count() > 0) && (m_oAttrItemsGID4.Count() == m_oMapedAccInfosGID4.Count()))
            {
               

                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID4)
                {
                    //GID4 是貸方為主 C
                    l_oFindAccInfo = m_oMapedAccInfosGID4[l_iIndex];

                    m_dOperaIncome = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                                                          m_dOperaIncome,
                                                          item);
                    /*
                    if (l_oFindAccInfo.CountFlag == "D")
                    {
                        //+借方 D
                        if (item.m_ComDAmount == null)
                        { item.m_ComDAmount = "0"; }
                        m_dOperaIncome += Convert.ToInt32(item.m_ComDAmount);

                        //-貸方 C
                        if (item.m_ComCAmount == null)
                        { item.m_ComCAmount = "0"; }
                        m_dOperaIncome -= Convert.ToInt32(item.m_ComCAmount);

                    }
                    else if (l_oFindAccInfo.CountFlag == "C")
                    {
                        //+貸方 C
                        if (item.m_ComCAmount == null)
                        { item.m_ComCAmount = "0"; }
                        m_dOperaIncome += Convert.ToInt32(item.m_ComCAmount);

                        //-借方 D
                        if (item.m_ComDAmount == null)
                        { item.m_ComDAmount = "0"; }
                        m_dOperaIncome -= Convert.ToInt32(item.m_ComDAmount);
                    }
                    */

                    l_iIndex++;
                }
            }

            return m_dOperaIncome;
        }

        //計算 TolCostOfCashExpend PS:要在Calc_GID5_OperaCosts() 之前執行
        public double Calc_GID5_TolCostOfCashExpend()
        {


            //現金支出總成本 = (PrintOrder:14~35相加)
            m_dTolCostOfCashExpend = 0; //Init


            if ((m_oAttrItemsGID5.Count() > 0) && (m_oAttrItemsGID5.Count() == m_oMapedAccInfosGID5.Count()))
            {


                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID5)
                {
                    //[　原料進料 鮮奶] ~ [　原料進料 包材]
                    //找出Group5中 AccountNo == 5200 且 DetailAccNo 不為""的 相加
                    if ((item.m_ComAccountNo == "5200") && (item.m_ComDetailAccNo != null))
                    {
                        //GID5 是借方為主 D
                        l_oFindAccInfo = m_oMapedAccInfosGID5[l_iIndex];

                        m_dTolCostOfCashExpend = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                                                              m_dTolCostOfCashExpend,
                                                              item);



                    }

                    l_iIndex++;
                }
            }

            return m_dTolCostOfCashExpend;

        }

        public double Calc_GID5_OperaCosts()
        {

            //營業總成本 = (GId=5 相加)
            m_dTolOperaCosts = 0; //Init


            if ((m_oAttrItemsGID5.Count() > 0) && (m_oAttrItemsGID5.Count() == m_oMapedAccInfosGID5.Count()))
            {

                //m_dTolOperaCosts = Fun_CommonExpression("D",
                //                                        m_dTolOperaCosts,
                //                                        m_oAttrBeginRawMaterial.First());

                /* 20201226 CCL-
                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID5)
                {

                    //只加 1192[原料期初] ~ 　5200-104[原料進料 包材]
                    //需判斷是否是1192 && Name是[原料期初]
                    if((item.m_ComAccountNo == "1192") &&
                       (item.m_ComAccName == "原料期初"))
                    {
                        m_oAttrBeginRawMaterial.Add(item);

                    }
                    //需判斷是否是1192 && Name是[減 原料期末]  不加
                    if((item.m_ComAccountNo == "1192") &&
                       (item.m_ComAccName == "原料期末") )
                    {
                        m_oAttrEndRawMaterial.Add(item);

                    }


                    //GID5 是借方為主 D
                    //l_oFindAccInfo = m_oMapedAccInfosGID5[l_iIndex];

                    //m_dTolOperaCosts = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                    //                                      m_dTolOperaCosts,
                    //                                      item);


                    l_iIndex++;

                }
                */
            }

            ///原料期初: 1,8,15,21 每月四個點做盤點; 第一次的叫做期初; 第二次的叫做期末
            //如果m_oAttrEndRawMaterial超過1筆以上,
            //代表中間有盤點N次,我們只取最後一次的盤點當做原料期末

            //營業總成本 = 1192[原料期初](都叫原料存貨) + (5200-001[原料進料 鮮奶] ~ 5200-104[原料進料 包材]) -
            //             1192[減 原料期末](都叫原料存貨)
            m_dTolOperaCosts = m_dTolCostOfCashExpend;

            //if (m_oAttrBeginRawMaterial != null)
            //不管C,D只要BeginRawMaterial的值就是相加,只要EndRawMaterial的值就是相減
            if (m_oAttrBeginRawMaterial.Count() > 0)
            {
                //DAmount,CAmount其中一個一定是有值
                if (m_oAttrBeginRawMaterial.First().m_ComDAmount == null)
                { m_oAttrBeginRawMaterial.First().m_ComDAmount = "0"; }
                if (m_oAttrBeginRawMaterial.First().m_ComCAmount == null)
                { m_oAttrBeginRawMaterial.First().m_ComCAmount = "0"; }

                //看哪個有值,加哪一個
                if(m_oAttrBeginRawMaterial.First().m_ComDAmount == "0")
                {
                    m_dTolOperaCosts += Convert.ToInt32(m_oAttrBeginRawMaterial.First().m_ComCAmount);
                } else if(m_oAttrBeginRawMaterial.First().m_ComCAmount == "0")
                {
                    m_dTolOperaCosts += Convert.ToInt32(m_oAttrBeginRawMaterial.First().m_ComDAmount);
                }
                
            }
            //if (m_oAttrEndRawMaterial != null)
            if (m_oAttrEndRawMaterial.Count() > 0)
            {
                //如果m_oAttrEndRawMaterial超過1筆以上,
                //代表中間有盤點N次,我們只取最後一次的盤點當做原料期末; 所以用Last()

                //DAmount,CAmount其中一個一定是有值
                if (m_oAttrEndRawMaterial.Last().m_ComDAmount == null)
                { m_oAttrEndRawMaterial.Last().m_ComDAmount = "0"; }
                if (m_oAttrEndRawMaterial.Last().m_ComCAmount == null)
                { m_oAttrEndRawMaterial.Last().m_ComCAmount = "0"; }

                

                //看哪個有值,減哪一個
                //if (m_oAttrEndRawMaterial.First().m_ComDAmount == "0")
                if (m_oAttrEndRawMaterial.Last().m_ComDAmount == "0")
                {
                    m_dTolOperaCosts -= Convert.ToInt32(m_oAttrEndRawMaterial.Last().m_ComCAmount);
                }
                    //else if (m_oAttrEndRawMaterial.First().m_ComCAmount == "0")
                else if (m_oAttrEndRawMaterial.Last().m_ComCAmount == "0")
                {
                    m_dTolOperaCosts -= Convert.ToInt32(m_oAttrEndRawMaterial.Last().m_ComDAmount);
                }
                
            }
                          

            return m_dTolOperaCosts;
        }

  

        public double Calc_GID6_OperaExpense()
        {
            //營業費用 = (GId=6 相加)
            m_dTopOperaExpense = 0;

            if ((m_oAttrItemsGID6.Count() > 0) && (m_oAttrItemsGID6.Count() == m_oMapedAccInfosGID6.Count()))
            {


                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID6)
                {
                    //GID6 是借方為主 D
                    l_oFindAccInfo = m_oMapedAccInfosGID6[l_iIndex];

                    m_dTopOperaExpense = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                                                          m_dTopOperaExpense,
                                                          item);


                    l_iIndex++;
                }
            }

            return m_dTopOperaExpense;
        }

        public double Calc_GID7_NonOperaIncome()
        {
            //非營業收入 = (GId=7 相加)
            m_dNonOperaIncome = 0;

            if ((m_oAttrItemsGID7.Count() > 0) && (m_oAttrItemsGID7.Count() == m_oMapedAccInfosGID7.Count()))
            {


                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID7)
                {
                    //GID7 是貸方為主 C
                    l_oFindAccInfo = m_oMapedAccInfosGID7[l_iIndex];

                    m_dNonOperaIncome = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                                                          m_dNonOperaIncome,
                                                          item);


                    l_iIndex++;
                }
            }

            return m_dNonOperaIncome;
        }

        public double Calc_GID8_NonOperaExpense()
        {
            //非營業支出 = (GId=8 相加)
            m_dNonOperaExpense = 0;

            if ((m_oAttrItemsGID8.Count() > 0) && (m_oAttrItemsGID8.Count() == m_oMapedAccInfosGID8.Count()))
            {


                int l_iIndex = 0;
                AccountInfo l_oFindAccInfo = null;
                foreach (TmpExcelItem item in m_oAttrItemsGID8)
                {
                    //GID7 是貸方為主 C
                    l_oFindAccInfo = m_oMapedAccInfosGID8[l_iIndex];

                    m_dNonOperaExpense = Fun_CommonExpression(l_oFindAccInfo.CountFlag,
                                                          m_dNonOperaExpense,
                                                          item);


                    l_iIndex++;
                }
            }

            return m_dNonOperaExpense;
        }

        //20201227 CCL+ 計算比率
        public double Calc_PercentVal(double p_dValue)
        {
            //(p_dValue / m_dOperaIncome ) * 100 //營業收入
            double l_dRtnVal = 0;
            const int l_iDotNum = 2;//四捨五入至小數點第2位
            
            //壁面無限大
            if ( (m_dOperaIncome != 0))
            {

                l_dRtnVal = (p_dValue / m_dOperaIncome) * 100;
                l_dRtnVal = Math.Round(l_dRtnVal, l_iDotNum);

                return l_dRtnVal;
            }
            return l_dRtnVal;
        }

        public double Calc_RestOthersVal()
        {
            m_dBtmOperaExpense = m_dTopOperaExpense; //營業費用 = (GId=6 相加)

            //[3].
            m_dActualTolCosts = m_dTolOperaCosts; //實際用量總成本 = 營業總成本 = (GId=5 相加)
            //[4].
            m_dOperaMargin = m_dOperaIncome - m_dTolOperaCosts; //營業毛利 = 營業收入 - 實際用量總成本

            //[7].
            m_dBussInterest = m_dOperaMargin - m_dBtmOperaExpense; //營業利益 = 營業毛利 - 營業費用

            //[10].
            m_dConsuCurrentProfitLoss = m_dOperaIncome - m_dActualTolCosts - m_dBtmOperaExpense
                            + m_dNonOperaIncome - m_dNonOperaExpense; //實際用量本期損益 = [1]-[3]-[6]+[8]-[9]

           
            //[12].
            m_dCashExpendForCurrPeriod = m_dOperaIncome - m_dTolCostOfCashExpend - m_dBtmOperaExpense +
                               m_dNonOperaIncome - m_dNonOperaExpense; //現金支出本期損益 = [1]-[11]-[6]+[8]-[9]

            return 0;
        }



        //Constructor
        public MERP_AccInfoExpore()
        {
            //if (m_oBaseAttrItems == null)
            m_oBaseAttrItems = new List<TmpExcelItem>();

            //if (m_oMapedAccInfos == null)
            //m_oMapedAccInfos = new List<AccountInfo>();

            //分群Base項目
            m_oAttrItemsGID4 = new List<TmpExcelItem>();
            m_oAttrItemsGID5 = new List<TmpExcelItem>();
            m_oAttrItemsGID6 = new List<TmpExcelItem>();
            m_oAttrItemsGID7 = new List<TmpExcelItem>();
            m_oAttrItemsGID8 = new List<TmpExcelItem>();

            m_oAttrItemsOthers = new List<TmpExcelItem>(); //存放排除的
            m_oAttrBeginRawMaterial = new List<TmpExcelItem>();
            m_oAttrEndRawMaterial = new List<TmpExcelItem>();

            //找出來的對應的AccountInfo Map項目
            //public List<AccountInfo> m_oMapedAccInfos { get; set; }
            m_oMapedAccInfosGID4 = new List<AccountInfo>();
            m_oMapedAccInfosGID5 = new List<AccountInfo>();
            m_oMapedAccInfosGID6 = new List<AccountInfo>();
            m_oMapedAccInfosGID7 = new List<AccountInfo>();
            m_oMapedAccInfosGID8 = new List<AccountInfo>();


            m_dOperaIncome = 0;
            m_dTolOperaCosts = 0;
            m_dActualTolCosts = 0;
            m_dOperaMargin = 0;
            m_dTopOperaExpense = 0;
            m_dBtmOperaExpense = 0;
            m_dBussInterest = 0;
            m_dNonOperaIncome = 0;
            m_dNonOperaExpense = 0;
            m_dConsuCurrentProfitLoss = 0;
            m_dTolCostOfCashExpend = 0;
            m_dCashExpendForCurrPeriod = 0;

        }

    }
}