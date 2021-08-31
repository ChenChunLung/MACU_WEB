using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MACU_WEB.Models;
using MACU_WEB.Services;

namespace MACU_WEB.BIServices
{
   

    public class MERP_AccInfoTolExcel
    {
        //放各店家調整過的合併Table (Combine One Full Table)

        public List<MERP_NewOrderPtrTB> m_ShopNWOPtrTBs { get; set; }


        //放各店家處理過的Data (Origin Processed Data)
        public List<MERP_AccInfoExpore> m_oRtnExpExcelShops { get; set; }

        public TmpExcelItem m_ZeroValExcelItem;

        //20201229 CCL+ 依據總表和所有的店家的AccInfo Group
        public List<AccountInfo> m_CombineAllAccInfo { get; set; }

        //所有店家的Group 合併
        public List<AccountInfo> m_oCombineAccInfosGID4 { get; set; }
        public List<AccountInfo> m_oCombineAccInfosGID5 { get; set; }
        public List<AccountInfo> m_oCombineAccInfosGID6 { get; set; }
        public List<AccountInfo> m_oCombineAccInfosGID7 { get; set; }
        public List<AccountInfo> m_oCombineAccInfosGID8 { get; set; }

        //存放各Row的所有Col Item合計
        public List<double> m_RowItemCombineAmount; //數目同m_CombineAllAccInfo.Count()
        //public MERP_NewOrderPtrTB m_RowItemComAmount { get; set; }

        public List<AccountInfo> m_BaseAccInfo; //基本固定一定要顯示的AccInfo 先從Map Table取出

        // public List<MERP_NewOrderPtrTB> m_ShopNewOrderPtrTBs;
        //存放每家店的New Order Pointer Table


        //20210103 CCL+ 計算每列的各Column的合計 ////////////////////////////////////////
        public double Calc_RowItemAllColAmount()
        {
            int l_iShopIndex = 0,l_iRowIndex = 0, l_ColIndex = 0;
            int l_iAllRowCount = 0;
           

            foreach (MERP_NewOrderPtrTB Shop in m_ShopNWOPtrTBs)
            {
                l_iRowIndex = 0;

                //如果合計List Count為0,初始化
                if (m_RowItemCombineAmount.Count() == 0)
                {
                    for (int i = 0; i < Shop.m_TBTolRowCount; i++)
                    {
                        m_RowItemCombineAmount.Add(0);
                    }
                }
            
                //1.累計 營業收入
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[0]; //0
                l_iRowIndex++;
                //2.累計 GID4                
                foreach (TmpExcelItem Item in Shop.m_oNOAttrItemsGID4)
                {
                    m_RowItemCombineAmount[l_iRowIndex] += Convert.ToInt32(Item.m_ComAmount); 
                    l_iRowIndex++;
                }
                //3.累計 營業總成本
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[1]; //1
                l_iRowIndex++;
                //4.累計 原料期初
                m_RowItemCombineAmount[l_iRowIndex] += GetBeginEndRawAmount(Shop.m_oNOAttrBeginRawMaterial[0]);
                //Convert.ToInt32(Shop.m_oNOAttrBeginRawMaterial[0].m_ComAmount); //Begin
                l_iRowIndex++;
                //5.累計 GID5
                foreach (TmpExcelItem Item in Shop.m_oNOAttrItemsGID5)
                {
                    m_RowItemCombineAmount[l_iRowIndex] += Convert.ToInt32(Item.m_ComAmount);
                    l_iRowIndex++;
                }
                //6.累計 減 原料期末
                m_RowItemCombineAmount[l_iRowIndex] += GetBeginEndRawAmount(Shop.m_oNOAttrEndRawMaterial[0]);
                                   //Convert.ToInt32(Shop.m_oNOAttrEndRawMaterial[0].m_ComAmount); //End
                l_iRowIndex++;
                //7.累計 實際用量總成本
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[2]; //2
                l_iRowIndex++;
                //8.累計 營業毛利
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[3]; //3
                l_iRowIndex++;
                //9.累計 營業費用
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[4]; //4
                l_iRowIndex++;
                //10.累計 GID6
                foreach (TmpExcelItem Item in Shop.m_oNOAttrItemsGID6)
                {
                    m_RowItemCombineAmount[l_iRowIndex] += Convert.ToInt32(Item.m_ComAmount);
                    l_iRowIndex++;
                }
                //11.累計 Btm營業費用
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[5]; //5
                l_iRowIndex++;
                //12.累計 營業利益
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[6]; //6
                l_iRowIndex++;
                //13.累計 非營業收入
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[7]; //7
                l_iRowIndex++;
                //14.累計 GID7
                foreach (TmpExcelItem Item in Shop.m_oNOAttrItemsGID7)
                {
                    m_RowItemCombineAmount[l_iRowIndex] += Convert.ToInt32(Item.m_ComAmount);
                    l_iRowIndex++;
                }
                //15.累計 非營業支出
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[8]; //8
                l_iRowIndex++;
                //16.累計 GID8
                foreach (TmpExcelItem Item in Shop.m_oNOAttrItemsGID8)
                {
                    m_RowItemCombineAmount[l_iRowIndex] += Convert.ToInt32(Item.m_ComAmount);
                    l_iRowIndex++;
                }
                //17.累計 實際用量本期損益
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[9]; //9
                l_iRowIndex++;
                //18. 空白跳過                
                //19.累計 現金支出總成本
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[10]; //10
                l_iRowIndex++;
                //19.累計 現金支出本期損益
                m_RowItemCombineAmount[l_iRowIndex] += Shop.m_NewOrderItemPtrs[11]; //11                


                l_iShopIndex++;
            }

            return 0;
        }

        public double GetBeginEndRawAmount(TmpExcelItem p_oBERawMat)
        {
            TmpExcelItem l_oBeginEndRawMat = p_oBERawMat;
            double tmpVal = 0;

            if (((l_oBeginEndRawMat.m_ComDAmount != null) && (l_oBeginEndRawMat.m_ComDAmount != "0")) &&
            ((l_oBeginEndRawMat.m_ComCAmount != null) && (l_oBeginEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                tmpVal = Convert.ToInt32(l_oBeginEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginEndRawMat.m_ComCAmount);
                                             
            }
            else
            if ((l_oBeginEndRawMat.m_ComDAmount != null) && (l_oBeginEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                tmpVal = Convert.ToDouble(l_oBeginEndRawMat.m_ComDAmount);
               
            }
            else if ((l_oBeginEndRawMat.m_ComCAmount != null) && (l_oBeginEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                tmpVal = Convert.ToDouble(l_oBeginEndRawMat.m_ComCAmount);
               
            }

            return tmpVal;
        }
        //20210103 CCL+ /////////////////////////////////////////////////////////////////

        //儲存ShopNo,ShopNa到New Order Table Object內
        public int SetShopNoNaToNewOrderTB(MERP_AccInfoExpore p_oShopOrgData, int p_iIndex)
        {
            m_ShopNWOPtrTBs[p_iIndex].m_sShopNo = p_oShopOrgData.m_sShopId;
            m_ShopNWOPtrTBs[p_iIndex].m_sShopName = p_oShopOrgData.m_sShopName;

            return 0;
        }



        //處理各家店的AccInfoGID4,AccInfoGID5,AccInfoGID6,AccInfoGID7,AccInfoGID8
        //20201230 CCL+ 把所有Shop的各AccInfoGID Combine在一起
        public int CombineAllShopGIDAccInfo(MERP_AccInfoExpore p_oShopInfo)
        {
            int l_iRowIndex = 0;

            //Combine All Shop GID4
            if(m_oCombineAccInfosGID4 != null)
            {
                l_iRowIndex = 0;
                if (m_oCombineAccInfosGID4.Count() == 0)
                {
                    //全加入
                    foreach(AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID4)
                    {
                        m_oCombineAccInfosGID4.Add(Info);
                    }
                    
                } else if(m_oCombineAccInfosGID4.Count() > 0)
                {
                    int l_iInsertIndex = 0;
                    m_oCombineAccInfosGID4 = m_oCombineAccInfosGID4.Union(p_oShopInfo.m_oMapedAccInfosGID4).ToList();
                    //比對不在的才加入
                    //foreach (AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID4)
                    //{
                    //    l_iRowIndex++;
                        //if (m_oCombineAccInfosGID4.Exists(m => (m.AccountNo == Info.AccountNo)
                        //                                    && (m.DetailAccNo == Info.DetailAccNo)
                        //                                    && (m.AccountName == Info.AccountName)
                        //                                    && (m.GroupID == Info.GroupID)
                        //                                    && (m.PrintOrder == Info.PrintOrder)
                        //                                    && (m.IsValid == Info.IsValid)) == false)
                    //    if(!m_oCombineAccInfosGID4.Contains(Info))
                    //    {
                            
                    //    }
                    //}
                }

            }

            //Combine All Shop GID5
            if (m_oCombineAccInfosGID5 != null)
            {
                l_iRowIndex = 0;
                if (m_oCombineAccInfosGID5.Count() == 0)
                {
                    //全加入
                    foreach (AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID5)
                    {
                        m_oCombineAccInfosGID5.Add(Info);
                    }

                }
                else if (m_oCombineAccInfosGID5.Count() > 0)
                {
                    int l_iInsertIndex = 0;
                    m_oCombineAccInfosGID5 = m_oCombineAccInfosGID5.Union(p_oShopInfo.m_oMapedAccInfosGID5).ToList();
                  
                }

            }

            //Combine All Shop GID6
            if (m_oCombineAccInfosGID6 != null)
            {
                l_iRowIndex = 0;
                if (m_oCombineAccInfosGID6.Count() == 0)
                {
                    //全加入
                    foreach (AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID6)
                    {
                        m_oCombineAccInfosGID6.Add(Info);
                    }

                }
                else if (m_oCombineAccInfosGID6.Count() > 0)
                {
                    int l_iInsertIndex = 0;
                    m_oCombineAccInfosGID6 = m_oCombineAccInfosGID6.Union(p_oShopInfo.m_oMapedAccInfosGID6).ToList();
                    //比對不在的才加入
                   
                }

            }

            //Combine All Shop GID7
            if (m_oCombineAccInfosGID7 != null)
            {
                l_iRowIndex = 0;
                if (m_oCombineAccInfosGID7.Count() == 0)
                {
                    //全加入
                    foreach (AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID7)
                    {
                        m_oCombineAccInfosGID7.Add(Info);
                    }

                }
                else if (m_oCombineAccInfosGID7.Count() > 0)
                {
                    int l_iInsertIndex = 0;
                    m_oCombineAccInfosGID7 = m_oCombineAccInfosGID7.Union(p_oShopInfo.m_oMapedAccInfosGID7).ToList();
                    //比對不在的才加入

                }

            }

            //Combine All Shop GID8
            if (m_oCombineAccInfosGID8 != null)
            {
                l_iRowIndex = 0;
                if (m_oCombineAccInfosGID8.Count() == 0)
                {
                    //全加入
                    foreach (AccountInfo Info in p_oShopInfo.m_oMapedAccInfosGID8)
                    {
                        m_oCombineAccInfosGID8.Add(Info);
                    }

                }
                else if (m_oCombineAccInfosGID8.Count() > 0)
                {
                    int l_iInsertIndex = 0;
                    m_oCombineAccInfosGID8 = m_oCombineAccInfosGID8.Union(p_oShopInfo.m_oMapedAccInfosGID8).ToList();
                    //比對不在的才加入

                }

            }

            //最後ReOrder m_CombineAllAccInfo的PrintOrder



            return 0;
        }


        public int ReOrderAllShopGIDAccInfo()
        {
            //ReOrder All Shop GID4 By PrintOrder
            if ((m_oCombineAccInfosGID4 != null) && (m_oCombineAccInfosGID4.Count() > 0))
            {
                m_oCombineAccInfosGID4 = m_oCombineAccInfosGID4.OrderBy(m => m.PrintOrder).ToList();
            }

            //ReOrder All Shop GID5 By PrintOrder
            if ((m_oCombineAccInfosGID5 != null) && (m_oCombineAccInfosGID5.Count() > 0))
            {
                m_oCombineAccInfosGID5 = m_oCombineAccInfosGID5.OrderBy(m => m.PrintOrder).ToList();
            }

            //ReOrder All Shop GID6 By PrintOrder
            if ((m_oCombineAccInfosGID6 != null) && (m_oCombineAccInfosGID6.Count() > 0))
            {
                m_oCombineAccInfosGID6 = m_oCombineAccInfosGID6.OrderBy(m => m.PrintOrder).ToList();
            }

            //ReOrder All Shop GID7 By PrintOrder
            if ((m_oCombineAccInfosGID7 != null) && (m_oCombineAccInfosGID7.Count() > 0))
            {
                m_oCombineAccInfosGID7 = m_oCombineAccInfosGID7.OrderBy(m => m.PrintOrder).ToList();
            }

            //ReOrder All Shop GID8 By PrintOrder
            if ((m_oCombineAccInfosGID8 != null) && (m_oCombineAccInfosGID8.Count() > 0))
            {
                m_oCombineAccInfosGID8 = m_oCombineAccInfosGID8.OrderBy(m => m.PrintOrder).ToList();
            }


            return 0;
        }

        //最後執行一次而已
        //處理過後各MERP_AccInfoExpore 店家的TmpExcelItem內會有新Order的順序值
        //處理塞入固定Item,所有Shop的AccInfo Combine在一起成為一個TB
        public int CombineAllShopAccInfoTB()
        {
            //將各Shop的AccInfo List排進全域static AccInfo List成員 //並把新的PrintOrder存進TmpExcelItem內
            //1.先利用丟進來的AccInfo 排出一個所有的AccInfo
            if (m_CombineAllAccInfo.Count() == 0)
            {
                //先塞入固定顯示的Row //20201230 CCL+ 依序放 4營業收入,5營業總成本,實際用量總成本,...                    
                if (m_BaseAccInfo.Count() > 0)
                {
                    //1. Add 4營業收入
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[0]);
                    //2. Group4
                    foreach (AccountInfo item in m_oCombineAccInfosGID4)
                    {
                        m_CombineAllAccInfo.Add(item);
                    }
                    //3. 5營業總成本
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[1]);
                    //4. 1192 原料期初
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[2]);
                    //5. Group5
                    foreach (AccountInfo item in m_oCombineAccInfosGID5)
                    {
                        m_CombineAllAccInfo.Add(item);
                    }
                    //6. 1192 減 原料期末
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[3]);
                    //7. S5實際用量總成本
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[4]);
                    //8. S4M5營業毛利
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[5]);
                    //9. 6營業費用
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[6]);
                    //10. Group6
                    foreach (AccountInfo item in m_oCombineAccInfosGID6)
                    {
                        m_CombineAllAccInfo.Add(item);
                    }
                    //11. S6營業費用
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[7]);
                    //12. S4M5M6營業利益
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[8]);
                    //13. 7非營業收入
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[9]);
                    //14. Group7
                    foreach (AccountInfo item in m_oCombineAccInfosGID7)
                    {
                        m_CombineAllAccInfo.Add(item);
                    }
                    //15. 8非營業支出
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[10]);
                    //14. Group8
                    foreach (AccountInfo item in m_oCombineAccInfosGID8)
                    {
                        m_CombineAllAccInfo.Add(item);
                    }
                    //16. S9實際用量本期損益
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[11]);
                    //17. S10現金支出總成本
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[12]);
                    //18. S11現金支出本期損益
                    m_CombineAllAccInfo.Add(m_BaseAccInfo[13]);

                }
            }

            //最後整合完的Row數,產生[共計List]
            //for(int i=0;  i< m_CombineAllAccInfo.Count(); i++ )
            //{
            //    int l_iVal = 0;
            //    m_RowItemCombineAmount.Add(l_iVal);
            //}

            return m_CombineAllAccInfo.Count();
        }

        //顯示內容為0的固定TmpExcelItem
        //public static void ShowZeroTmpExcelItem()
        //{

        //}

        //從原來資料整出一個按m_CombineAllAccInfo順序排序的Table,空的Row塞ZeroTmpExcelItem
        //public int CombineShopExcelItem(MERP_AccInfoExpore p_oOneShopData)
        //{
        //    if((m_CombineAllAccInfo != null) && (m_CombineAllAccInfo.Count() > 0))
        //    {
        //

        //    }
        //}

        //必須在CombineAllShopGIDAccInfo 之後執行
        public int CompareNOAccInfo(MERP_AccInfoExpore p_oOneShopData)
        {
            //一家店,一個MERP_NewOrderPtrTB 物件
            MERP_NewOrderPtrTB l_oNWOTBItem = new MERP_NewOrderPtrTB();
            m_ShopNWOPtrTBs.Add(l_oNWOTBItem);

            //1.加入 4營業收入
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dOperaIncome); //0
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+

            //找Old Data中是否有
            //2. 找GID4加入            
            foreach (AccountInfo Info in m_oCombineAccInfosGID4)
            {
              
                if (p_oOneShopData.m_oMapedAccInfosGID4.Contains(Info))
                {
                    int l_iIndex = 0;
                    //TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID4.Find(m => (m.m_ComAccountNo == Info.AccountNo) &&
                    //                                                                    (m.m_ComDetailAccNo == Info.DetailAccNo) &&
                    //                                                                    (m.m_PrintOrder == Info.PrintOrder));
                    l_iIndex = p_oOneShopData.m_oMapedAccInfosGID4.IndexOf(Info);
                    TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID4[l_iIndex];

                    l_oNWOTBItem.m_oNOAttrItemsGID4.Add(l_oTmpItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                } else
                {
                    //加入空ZeroItem
                    l_oNWOTBItem.m_oNOAttrItemsGID4.Add(m_ZeroValExcelItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
            }
            //3. 5營業總成本
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dTolOperaCosts); //1
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //4. 1192 原料期初 //m_oAttrBeginRawMaterial
            if (p_oOneShopData.m_oAttrBeginRawMaterial.Count() > 0)
            {
                l_oNWOTBItem.m_oNOAttrBeginRawMaterial.Add(p_oOneShopData.m_oAttrBeginRawMaterial.First());
                l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            } else
            {
                l_oNWOTBItem.m_oNOAttrBeginRawMaterial.Add(m_ZeroValExcelItem);
                l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            }
                
            //5. Group5
            foreach (AccountInfo Info in m_oCombineAccInfosGID5)
            {

                if (p_oOneShopData.m_oMapedAccInfosGID5.Contains(Info))
                {
                    int l_iIndex = 0;
                    l_iIndex = p_oOneShopData.m_oMapedAccInfosGID5.IndexOf(Info);
                    TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID5[l_iIndex];

                    l_oNWOTBItem.m_oNOAttrItemsGID5.Add(l_oTmpItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
                else
                {
                    //加入空ZeroItem
                    l_oNWOTBItem.m_oNOAttrItemsGID5.Add(m_ZeroValExcelItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
            }
            //6. 1192 減 原料期末
            if (p_oOneShopData.m_oAttrEndRawMaterial.Count() > 0)
            {
                l_oNWOTBItem.m_oNOAttrEndRawMaterial.Add(p_oOneShopData.m_oAttrEndRawMaterial.Last());
                l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            } else
            {
                l_oNWOTBItem.m_oNOAttrEndRawMaterial.Add(m_ZeroValExcelItem);
                l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            }
                
            //7. S5實際用量總成本
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dActualTolCosts); //2
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //8. S4M5營業毛利
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dOperaMargin); //3
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //9. 6營業費用
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dTopOperaExpense); //4
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //10. Group6
            foreach (AccountInfo Info in m_oCombineAccInfosGID6)
            {

                if (p_oOneShopData.m_oMapedAccInfosGID6.Contains(Info))
                {
                    int l_iIndex = 0;
                    l_iIndex = p_oOneShopData.m_oMapedAccInfosGID6.IndexOf(Info);
                    TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID6[l_iIndex];

                    l_oNWOTBItem.m_oNOAttrItemsGID6.Add(l_oTmpItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
                else
                {
                    //加入空ZeroItem
                    l_oNWOTBItem.m_oNOAttrItemsGID6.Add(m_ZeroValExcelItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
            }
            //11. Btm營業費用
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dBtmOperaExpense); //5
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //12. 營業利益
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dBussInterest); //6
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //13. 7非營業收入
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dNonOperaIncome); //7
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //14. Group7
            foreach (AccountInfo Info in m_oCombineAccInfosGID7)
            {

                if (p_oOneShopData.m_oMapedAccInfosGID7.Contains(Info))
                {
                    int l_iIndex = 0;
                    l_iIndex = p_oOneShopData.m_oMapedAccInfosGID7.IndexOf(Info);
                    TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID7[l_iIndex];

                    l_oNWOTBItem.m_oNOAttrItemsGID7.Add(l_oTmpItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
                else
                {
                    //加入空ZeroItem
                    l_oNWOTBItem.m_oNOAttrItemsGID7.Add(m_ZeroValExcelItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
            }
            //15. 8非營業支出
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dNonOperaExpense); //8
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //14. Group8
            foreach (AccountInfo Info in m_oCombineAccInfosGID8)
            {

                if (p_oOneShopData.m_oMapedAccInfosGID8.Contains(Info))
                {
                    int l_iIndex = 0;
                    l_iIndex = p_oOneShopData.m_oMapedAccInfosGID8.IndexOf(Info);
                    TmpExcelItem l_oTmpItem = p_oOneShopData.m_oAttrItemsGID8[l_iIndex];

                    l_oNWOTBItem.m_oNOAttrItemsGID8.Add(l_oTmpItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
                else
                {
                    //加入空ZeroItem
                    l_oNWOTBItem.m_oNOAttrItemsGID8.Add(m_ZeroValExcelItem);
                    l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
                }
            }
            //16. S9實際用量本期損益
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dConsuCurrentProfitLoss); //9
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //17. S10現金支出總成本
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dTolCostOfCashExpend); //10
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+
            //18. S11現金支出本期損益
            l_oNWOTBItem.m_NewOrderItemPtrs.Add(p_oOneShopData.m_dCashExpendForCurrPeriod); //11
            l_oNWOTBItem.m_TBTolRowCount++; //20210103 CCL+


            return l_oNWOTBItem.m_TBTolRowCount;
            //return l_oNWOTBItem.m_NewOrderItemPtrs.Count();
        }


        public MERP_AccInfoTolExcel()
        {
            //m_CombineAllAccInfo = new List<AccountInfo>();
            m_ShopNWOPtrTBs = new List<MERP_NewOrderPtrTB>();

            m_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();

            m_CombineAllAccInfo = new List<AccountInfo>();

            m_RowItemCombineAmount = new List<double>();
            // = new MERP_NewOrderPtrTB();

            m_BaseAccInfo = new List<AccountInfo>();

            m_oCombineAccInfosGID4 = new List<AccountInfo>();
            m_oCombineAccInfosGID5 = new List<AccountInfo>();
            m_oCombineAccInfosGID6 = new List<AccountInfo>();
            m_oCombineAccInfosGID7 = new List<AccountInfo>();
            m_oCombineAccInfosGID8 = new List<AccountInfo>();

            //最後一項為內容為0的空TmpExcelItem
            m_ZeroValExcelItem = new TmpExcelItem();
            m_ZeroValExcelItem.m_ComAmount = "0";
            m_ZeroValExcelItem.m_ComCAmount = "0";
            m_ZeroValExcelItem.m_ComDAmount = "0";
            m_ZeroValExcelItem.m_PrintOrder = 0;

        }

        //從Map Table取出基本固定一定要顯示的AccInfo  
        public int GetBaseItemsFromAccInfos(MERP_AccountInfoDBService p_oAccInfoDBService)
        {
            TmpExcelItem item = null;
            AccountInfo l_oTmpAccInfoItem = null;
            
            //從Map Table取出基本固定一定要顯示的AccInfo
            //4營業收入
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("4",
                                                                   "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //5營業總成本
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("5",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //1192 原料期初
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("1192",
                                                                  "", "原料期初");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //1192 減 原料期末
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("1192",
                                                                  "", "減 原料期末");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S5實際用量總成本
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S5",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S4M5營業毛利
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S4M5",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //6營業費用
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("6",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S6營業費用
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S6",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S4M5M6營業利益
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S4M5M6",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //7非營業收入
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("7",
                                                                  "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //8非營業支出
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("8",
                                                                 "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S9實際用量本期損益
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S9",
                                                               "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S10現金支出總成本
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S10",
                                                               "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);
            //S11現金支出本期損益
            l_oTmpAccInfoItem = p_oAccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo("S11",
                                                             "");
            m_BaseAccInfo.Add(l_oTmpAccInfoItem);

            return m_BaseAccInfo.Count();
        }
    }
}