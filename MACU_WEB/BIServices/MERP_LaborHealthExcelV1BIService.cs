using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using System.Data;
using System.IO;
using System.Web.Mvc;
using MACU_WEB.Services;
using MACU_WEB.Models;
using MACU_WEB.Models._Base;
using System.Diagnostics;
using MACU_WEB.Areas.MERP_TCF000.ViewModels;


namespace MACU_WEB.BIServices
{
    public class MERP_LaborHealthExcelV1BIService
    {
        //Origin
        private static XLWorkbook m_oWorkbook = null;
        private static IXLTable m_oImpTable = null;
        //Will be to Modify
        private static XLWorkbook m_oModWorkbook = null;
        private static IXLTable m_oModTable = null;

        private static DataTable m_oExpTable = null;

        //File Content
        public static MERP_FileContentDBService m_FileDBService = new MERP_FileContentDBService();
        public static MERP_FA_LaborHealthInsV1DBService m_LHInsV1DBService = new MERP_FA_LaborHealthInsV1DBService();
        //勞保費率設定
        public static MERP_FA_LaborInsSetDBService m_LaborInsSetDBService = new MERP_FA_LaborInsSetDBService();
        //健保保費率設定
        public static MERP_FA_HealthInsSetDBService m_HealInsSetDBService = new MERP_FA_HealthInsSetDBService();


        public static Boolean ImportExcel(String p_sFullFilePath)
        {

            //20201204 CCL+ 利用ClosedXML匯入Excel檔
            //判斷上傳的Excel是否存在
            if (p_sFullFilePath != null)
            {
                //Origin
                XLWorkbook m_oWorkbook = new XLWorkbook(p_sFullFilePath);

                //複製一份要修改的副本WorkSheet到新的ModWookbook
                m_oModWorkbook = new XLWorkbook();
                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

                //讀取第一個Sheet (多部門損益表)
                IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheet(1);
                //讀取Modify Sheet
                IXLWorksheet l_oWooksheetMod = m_oModWorkbook.Worksheet(1);

                //定義資料使用到的Range範圍 起始/結束Cell
                var l_oFirstCell = l_oWooksheet.FirstCellUsed();
                var l_oEndCell = l_oWooksheet.LastCellUsed();
                //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
                var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);

                //定義資料使用到的Range範圍 起始/結束Cell
                var l_oFirstCellMod = l_oWooksheetMod.FirstCellUsed();
                var l_oEndCellMod = l_oWooksheetMod.LastCellUsed();
                //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
                var l_oDataMod = l_oWooksheetMod.Range(l_oFirstCellMod, l_oEndCellMod);

                //將資料範圍轉型
                m_oImpTable = l_oData.AsTable();
                //IXLTable指向要修改的副本WookSheet的Table
                m_oModTable = l_oDataMod.AsTable();


                //讀取資料 For Debug
                string l_sExcel = "";
                l_sExcel = m_oImpTable.Cell(8, 11).Value.ToString();
                l_sExcel = m_oModTable.Cell(8, 11).Value.ToString();

                //寫入資料
                //m_oImpTable.Cell(2, 1).Value = "test";

                //資料顯示
                //Response.Write("<script language=javascript>alert(" + l_sExcel + ");</script>");

                return true;
                //
            }
            else
            {

                return false;
            }


        }

        //
        public static DataTable ImportExcelToDataTable(IXLTable p_oImpTable)
        {
            //Create a new DataTable.
            DataTable l_oDt = new DataTable();

            //Loop through the Worksheet rows.
            bool firstRow = true;
            foreach (IXLRangeRow row in p_oImpTable.Rows())
            {
                //Use the first row to add columns to DataTable.
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        l_oDt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;

                }
                else
                {
                    //Add rows to DataTable.
                    l_oDt.Rows.Add();

                    int i = 0;

                    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                    {
                        l_oDt.Rows[l_oDt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }

                }
            }

            return l_oDt;
        }


        public static Boolean SaveAsExcel(String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            if ((m_oWorkbook != null) && (m_oModTable != null))
            {
                string type = ".xlsx";//獲取副檔名
                string name = Guid.NewGuid().ToString();//創建文件名稱

                string l_sFilePath = p_oServer.MapPath("~/Dn_Data/" + p_sPROG_ID + "/" + System.DateTime.Now.ToString("yyyyMMdd"));

                //文件夾不存在創建文件夾
                if (!Directory.Exists(l_sFilePath))
                {
                    Directory.CreateDirectory(l_sFilePath);
                }


                string l_sUrl = l_sFilePath + "\\" + name + type; //保存文件


                //利用DBService將檔案資訊存入資料庫
                m_FileDBService.FileContent_DBCreate(name, l_sUrl, 0, "xlsx", "Dn", p_sPROG_ID);

                //l_oFiles.SaveAs(l_sUrl);//保存文件
                m_oModWorkbook.SaveAs(l_sUrl);

            }

            return true;
        }

        //20201204 CCL+ 利用ClosedXML匯出Excel檔
        public static Boolean ExportExcel(String p_sFullFilePath)
        {
            //判斷下載的Excel檔名是否存在
            if (System.IO.File.Exists(p_sFullFilePath))
            {
                //
                XLWorkbook l_oWorkbook = new XLWorkbook();

                l_oWorkbook.Worksheets.Add(m_oExpTable);
                l_oWorkbook.SaveAs(p_sFullFilePath);

                return true;
            }

            return false;
        }


        //20201215 CCL+ 載入上傳的Excel
        public static Boolean ImportExcelCommon(String p_sFullFilePath)
        {
            //20201204 CCL+ 利用ClosedXML匯入Excel檔
            //判斷上傳的Excel是否存在
            if (p_sFullFilePath != null)
            {
                //Origin
                XLWorkbook m_oWorkbook = new XLWorkbook(p_sFullFilePath);

                //讀取第一個Sheet 
                IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheet(1);


                //定義資料使用到的Range範圍 起始/結束Cell
                var l_oFirstCell = l_oWooksheet.FirstCellUsed();
                var l_oEndCell = l_oWooksheet.LastCellUsed();
                //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
                var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);

                //將資料範圍轉型
                m_oImpTable = l_oData.AsTable();

                //讀取資料 For Debug
                string l_sExcel = "";
                l_sExcel = m_oImpTable.Cell(8, 11).Value.ToString();



                return true;
                //
            }
            else
            {

                return false;
            }


        }


        public static string ImportExcelTo_FA_LaborHealthV1(FileContent p_oSearchFile)
        {
            //載入的資料月份
            string l_sLHInsMonth = "";

            //載入上傳的Excel
            ImportExcelCommon(p_oSearchFile.Url);
            //判斷這個月會計期別是否已在DB有資料,有的話,刪除舊的
            if (m_oImpTable != null)
            {
                //從上傳Excel中抓出LHInsPeriod
                //根據傳入的年,月
                int l_iYear = 0, l_iMonth = 0;

                l_iYear = Convert.ToInt32(p_oSearchFile.DataYear);
                l_iMonth = Convert.ToInt32(p_oSearchFile.DataMonth);
                //找出資料庫是否有本年月
                Boolean l_bIsExistData = m_LHInsV1DBService.FA_LaborHealthInsV1_ChkDataByYearMon(l_iYear, l_iMonth);
                if (l_bIsExistData)
                {
                    //刪除舊的
                    m_LHInsV1DBService.FA_LaborHealthInsV1_DBDeleteByYearMon(l_iYear, l_iMonth);

                }
                //,並且匯入DataBase , 傳入手動Key In 資料年月                
                m_LHInsV1DBService.FA_LaborHealthInsV1_SqlDBCreate(m_oImpTable, l_iYear, l_iMonth); //改用ADO.NET提升速度
            }
            //return true;
            return l_sLHInsMonth;


        }

        public static List<FA_LaborHealthInsV1> GetImportExcelInDB_YearMonthData(string p_sYear, string p_sMonth)
        {
            int l_iYear = Convert.ToInt32(p_sYear);
            int l_iMonth = Convert.ToInt32(p_sMonth);
            return m_LHInsV1DBService.FA_LaborHealthInsV1_GetDataByYearMon(l_iYear, l_iMonth).ToList();
        }

        public static List<FA_LaborHealthInsV1> GetImportExcelInDB_YearMonthDataPage(string p_sYear,
                                                                                string p_sMonth, int p_iPageing)
        {

            return m_LHInsV1DBService.FA_LaborHealthInsV1_GetDataByYearMonthPage(p_sYear, p_sMonth, p_iPageing);
        }

        public static List<FA_LaborHealthInsV1> TransDataTableToList(DataSet p_oDataSet)
        {
                        
            return m_LHInsV1DBService.FA_LaborHealthInsV1_DataTableTo_FALHInsV1List(p_oDataSet);
            //FA_FaJournal_DataTableTo_FaJournalsList
        }

        //從DB中處理匯入的Excel
        public static List<FA_LaborHealthInsV1> ProcessImportExcelFromDB(MERP_TCF004_JournalsOptions p_oOption)
        {
            DataSet l_oDataSet = null;
            List<FA_LaborHealthInsV1> l_RtnList = null;

            if (p_oOption != null)
            {
                //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
                l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
                l_RtnList = TransDataTableToList(l_oDataSet);
                return l_RtnList;
            }


            return null;

        }


        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptions(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {            
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";            
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
       
            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address, 
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //健保                  
            l_oWooksheet.Cell(3, TOPCOLS + 3).Value = "健保";
            //勞退                     
            l_oWooksheet.Cell(3, TOPCOLS + 5).Value = "勞退";
            //合計                     
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 健保 勞退 合計 Title  Merge 2 Cols                                 
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 2 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3 + TOPCOLS).Address, l_oWooksheet.Cell(3, 4 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 5 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 8 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保,健保,勞退 三組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;              

               
                l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                if(i == 2)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                } else
                {                   
                    
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }
                
            }

            //勞退的兩行 Merge
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART-2)) + 1).Address, 
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART-2)) + 2).Address);
            l_oRangeRetireTitle.Merge();



            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
                l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            //}

          
            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                ExporeOnePartExcelList(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                    p_sPROG_ID, p_oServer);
                

                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////

   

            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;

            
        }


        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        public static bool ExporeOnePartExcelList(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
                        
            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;


            if((l_oRtnExporeExcel != null) && 
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach(MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    ++l_iRowIndex;
                    //1.顯示 公司
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS,  1).Value = Item.m_sPlusInsCompany;
                    //2.顯示 門市
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS,  2).Value = Item.m_sDepartName;
                    //3.顯示 姓名
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS,  3).Value = Item.m_sMemberName;
                    //4.顯示 勞保-單位(含職災)
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //5.顯示 勞保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //6.顯示 健保-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //7.顯示 健保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //8.顯示 勞退-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //9.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //10.顯示 合計-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //11.顯示 合計-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //12.顯示 總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;

                    //勞退 的值的Cell要合併
                    IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART-2)) + 1).Address, 
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    l_oRetireInsCol.Merge();

                }

            }
         
   
            //18.顯示 空白
            ++l_iRowIndex;
            for(int i=1; i<=8; i++ )
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS,  i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色
            //IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Address);
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) +1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oTolRetireInsCol.Merge();


            //Styling          
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;



            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);

            return true;
        }

        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        /// 20210125 CCL+ 新增勞健保小計
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV1_1(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            //const int PERCOM_COL = 2; //勞健保小計 2行
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            //l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //健保                  
            //l_oWooksheet.Cell(3, TOPCOLS + 3).Value = "健保";
            l_oWooksheet.Cell(3, TOPCOLS + 4).Value = "健保";
            //勞退                     
            //l_oWooksheet.Cell(3, TOPCOLS + 5).Value = "勞退";
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "勞退";
            //合計                     
            //l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "合計";
            l_oWooksheet.Cell(3, TOPCOLS + 10).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 健保 勞退 合計 Title  Merge 2 Cols                                 
            //IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 2 + TOPCOLS).Address);
            //l_oRange8.Merge();
            //IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3 + TOPCOLS).Address, l_oWooksheet.Cell(3, 4 + TOPCOLS).Address);
            //l_oRange9.Merge();
            //IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 5 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            //l_oRange10.Merge();
            //IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 8 + TOPCOLS).Address);
            //l_oRange11.Merge();
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 3 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 4 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 9 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 10 + TOPCOLS).Address, l_oWooksheet.Cell(3, 12 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保,健保,勞退 三組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                if (i == 2)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }
                //20210125 CCL+ 只有勞保 健保 有小計
                if(i < 2)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "小計";
                }

            }

            //勞退的三行 Merge
            //IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                                            l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRangeRetireTitle.Merge();
            //合計的後兩行 Merge
            IXLRange l_oRangeTolTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oRangeTolTitle.Merge();


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                //l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                l_iTolRowCount = ExporeOnePartExcelListV3_1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }


        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        //20210129 CCL+, 多勞保基金墊償 版本
        /// 20210125 CCL+ 新增勞健保小計
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV1_2(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            //const int PERCOM_COL = 2; //勞健保小計 2行
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet; //用最新的當作預設設定
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            //l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //墊償                   
            //l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            l_oWooksheet.Cell(3, TOPCOLS + 4).Value = "墊償";
            //健保                  
            //l_oWooksheet.Cell(3, TOPCOLS + 3).Value = "健保";
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "健保";
            //勞退                     
            //l_oWooksheet.Cell(3, TOPCOLS + 5).Value = "勞退";
            l_oWooksheet.Cell(3, TOPCOLS + 10).Value = "勞退";
            //合計                     
            //l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "合計";
            l_oWooksheet.Cell(3, TOPCOLS + 13).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 墊償 健保 勞退 合計 Title  Merge 3 Cols                                            
            //勞保 6行
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange8.Merge();           
            //健保 3行
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 9 + TOPCOLS).Address);
            l_oRange9.Merge();
            //勞退 3行
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 10 + TOPCOLS).Address, l_oWooksheet.Cell(3, 12 + TOPCOLS).Address);
            l_oRange10.Merge();
            //合計  3行
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 13 + TOPCOLS).Address, l_oWooksheet.Cell(3, 15 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保0,墊償1,健保2,勞退3,合計4 五組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                if(i != 1)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                }
                
                if (i == 3)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else if(i != 1)
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }
                //20210125 CCL+ 只有墊償 健保 有小計
                //墊償
                if (i == 1)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 1).Value = "墊償";
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "小計";
                }
                if (i==2)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "小計";
                }
                //勞保第3行改 "單+個"
                if (i == 0)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "單+個";
                }

              

            }

            //勞退的三行 Merge
            //IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                                            l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRangeRetireTitle.Merge();
            //合計的後兩行 Merge
            IXLRange l_oRangeTolTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oRangeTolTitle.Merge();
            //墊償的後兩行 Merge           
            IXLRange l_oRangeFundTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oRangeFundTitle.Merge();


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                //l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                //l_iTolRowCount = ExporeOnePartExcelListV3_1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細(含墊償)
                l_iTolRowCount = ExporeOnePartExcelListV3_2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }


        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        /// 20210201 CCL+
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV2_2(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            //const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;  //用最新的當作預設設定
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //墊償                               
            l_oWooksheet.Cell(3, TOPCOLS + 4).Value = "墊償";
            //健保                  
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "健保";
            //勞退                     
            l_oWooksheet.Cell(3, TOPCOLS + 10).Value = "勞退";
            //合計                     
            l_oWooksheet.Cell(3, TOPCOLS + 13).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 墊償 健保 勞退 合計 Title  Merge 3 Cols                                 
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 9 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 10 + TOPCOLS).Address, l_oWooksheet.Cell(3, 12 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 13 + TOPCOLS).Address, l_oWooksheet.Cell(3, 15 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保0,墊償1,健保2,勞退3,合計4 五組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                if (i != 1)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                }

                if (i == 3)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else if (i != 1)
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }
                //20210125 CCL+ 只有墊償 健保 有小計
                //墊償
                if (i == 1)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 1).Value = "墊償";
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "小計";
                }
                if (i == 2)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "小計";
                }
                //勞保第3行改 "單+個"
                if (i == 0)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "單+個";
                }



            }

            //勞退的三行 Merge
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRangeRetireTitle.Merge();
            //合計的後兩行 Merge
            IXLRange l_oRangeTolTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oRangeTolTitle.Merge();
            //墊償的後兩行 Merge           
            IXLRange l_oRangeFundTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oRangeFundTitle.Merge();


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                ///l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                ///                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4_1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示總和 含墊償
                l_iTolRowCount = ExporeOnePartExcelListV4_2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }


        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        /// 20210126 CCL+
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV2_1(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //健保                  
            l_oWooksheet.Cell(3, TOPCOLS + 4).Value = "健保";
            //勞退                     
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "勞退";
            //合計                     
            l_oWooksheet.Cell(3, TOPCOLS + 10).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 健保 勞退 合計 Title  Merge 2 Cols                                 
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 3 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 4 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 9 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 10 + TOPCOLS).Address, l_oWooksheet.Cell(3, 12 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保,健保,勞退 三組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                if (i == 2)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }
                //20210125 CCL+ 只有勞保 健保 有小計
                if (i < 2)
                {
                    l_oWooksheet.Cell(4, iCol1Index + 3).Value = "小計";
                }

            }

            //勞退的三行 Merge
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRangeRetireTitle.Merge();
            //合計的後兩行 Merge
            IXLRange l_oRangeTolTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                                       l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oRangeTolTitle.Merge();


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                ///l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                ///                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示總和
                l_iTolRowCount = ExporeOnePartExcelListV4_1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }
       


        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        /// 20210118 CCL+
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV1(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate ;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //健保                  
            l_oWooksheet.Cell(3, TOPCOLS + 3).Value = "健保";
            //勞退                     
            l_oWooksheet.Cell(3, TOPCOLS + 5).Value = "勞退";
            //合計                     
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 健保 勞退 合計 Title  Merge 2 Cols                                 
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 2 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3 + TOPCOLS).Address, l_oWooksheet.Cell(3, 4 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 5 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 8 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保,健保,勞退 三組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                if (i == 2)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }

            }

            //勞退的兩行 Merge
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oRangeRetireTitle.Merge();



            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);
                //顯示總和
                //l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount , TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }


        /// /////////////////////////////////////////////////////////////////////////////////////////////////// 
        /// 20210118 CCL+
        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptionsV2(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            //int l_iPartCount = 3;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();

            //取出最新勞保費率設定
            FA_LaborInsSet l_oLaborInsSet = m_LaborInsSetDBService.FA_LaborInsSet_GetDataByNewestBeginDate();
            //取出最新健保費率設定
            FA_HealthInsSet l_oHealInsSet = m_HealInsSetDBService.FA_HealthInsSet_GetDataByNewestBeginDate();

            //產生最終Expore要輸出的Excel的打包總物件
            MERP_LHInsExcelExpore l_oRtnTolExcelLHIns = new MERP_LHInsExcelExpore();
            l_oRtnTolExcelLHIns.m_oLaborInsSet = l_oLaborInsSet;
            l_oRtnTolExcelLHIns.m_oHealInsSet = l_oHealInsSet;
            //20210119 CCL+ 選擇日期
            l_oRtnTolExcelLHIns.m_StartDate = p_oOption.m_sOnJobDate;
            l_oRtnTolExcelLHIns.m_EndDate = p_oOption.m_sResignDate;

            //要比對的DataSet            

            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;
            int l_iPadColIndex = PRINT_COLS * PRINT_PART;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            //
            //202210103 CCL- IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //20210105 CCL- l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontName = "標楷體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202210103 CCL- IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange6.Merge();
            //202210103 CCL- IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange7.Merge();

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, TOPCOLS + 4 + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + 4 + 2).Address);
            //l_oRetireInsCol.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //勞健保表           
            l_oWooksheet.Cell(1, 1).Value = "勞健保表";

            //資料年月份：              
            //l_oWooksheet.Cell(2, 1).Value = "資料年月份: " + p_oOption.m_sDataYear + "年  " + p_oOption.m_sDataMonth + " 月";
            l_oWooksheet.Cell(2, 1).Value = "選擇日期: " + p_oOption.m_sOnJobDate + "  ~  " + p_oOption.m_sResignDate;

            //公司
            l_oWooksheet.Cell(3, 1).Value = "公司"; //
            //門市
            l_oWooksheet.Cell(3, 2).Value = "門市";
            //姓名
            l_oWooksheet.Cell(3, 3).Value = "姓名";

            //公司, 門市, 姓名, 總合計 2 Rows合併且Vertical置中
            IXLRange l_oRangeComTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address); ; //
            l_oRangeComTitle.Merge();
            l_oRangeComTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeDepartTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2).Address, l_oWooksheet.Cell(4, 2).Address); ; //
            l_oRangeDepartTitle.Merge();
            l_oRangeDepartTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeNameTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3).Address, l_oWooksheet.Cell(4, 3).Address); ; //
            l_oRangeNameTitle.Merge();
            //總合計 2 Rows合併
            l_oRangeNameTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            IXLRange l_oRangeTolAmountTitle = l_oWooksheet.Range(l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address,
                                                                l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Address); //
            l_oRangeTolAmountTitle.Merge();
            l_oRangeTolAmountTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //勞保                     
            l_oWooksheet.Cell(3, TOPCOLS + 1).Value = "勞保";
            //健保                  
            l_oWooksheet.Cell(3, TOPCOLS + 3).Value = "健保";
            //勞退                     
            l_oWooksheet.Cell(3, TOPCOLS + 5).Value = "勞退";
            //合計                     
            l_oWooksheet.Cell(3, TOPCOLS + 7).Value = "合計";
            //總合計                    
            l_oWooksheet.Cell(3, TOPCOLS + (PRINT_COLS * PRINT_PART) + 1).Value = "總合計";

            //20201227 CCL+ 勞保 健保 勞退 合計 Title  Merge 2 Cols                                 
            IXLRange l_oRange8 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + TOPCOLS).Address, l_oWooksheet.Cell(3, 2 + TOPCOLS).Address);
            l_oRange8.Merge();
            IXLRange l_oRange9 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 3 + TOPCOLS).Address, l_oWooksheet.Cell(3, 4 + TOPCOLS).Address);
            l_oRange9.Merge();
            IXLRange l_oRange10 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 5 + TOPCOLS).Address, l_oWooksheet.Cell(3, 6 + TOPCOLS).Address);
            l_oRange10.Merge();
            IXLRange l_oRange11 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 7 + TOPCOLS).Address, l_oWooksheet.Cell(3, 8 + TOPCOLS).Address);
            l_oRange11.Merge();

            //勞保,健保,勞退 三組
            for (int i = 0; i < PRINT_PART; i++)
            {
                //單位,個人
                int iCol1Index = (PRINT_COLS * i) + TOPCOLS;


                l_oWooksheet.Cell(4, iCol1Index + 1).Value = "單位";
                if (i == 2)
                {
                    //第三組勞退只有單位
                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "";
                }
                else
                {

                    l_oWooksheet.Cell(4, iCol1Index + 2).Value = "個人";
                }

            }

            //勞退的兩行 Merge
            IXLRange l_oRangeRetireTitle = l_oWooksheet.Range(l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                        l_oWooksheet.Cell(4, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oRangeRetireTitle.Merge();



            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            //20201231 CCL+ 科目名稱改成唯一一行
            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            //l_oRange2.Merge();
            //l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //儲存本年月資料
            List<DataSet> l_oDataSetList = new List<DataSet>();
            //if (p_oOption.m_iShopCount > 0)
            //{
            //    foreach (string shopNo in p_oOption.m_sShopList)
            //    {
            //        p_oOption.m_sTmpShopNo = shopNo;
            //        l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            //        l_oDataSetList.Add(l_oDataSet);
            //    }
            //}
            //else
            //{
            //20210120 CCL- l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptions(p_oOption);
            l_oDataSet = m_LHInsV1DBService.FA_LaborHealthInsV1_SqlGetDataListByOptionsV1(p_oOption);
            l_oDataSetList.Add(l_oDataSet);
            //}


            int l_iDataIndex = 0;
            //最終Expore要輸出的Data打包物件

            //l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init
            int l_iTolRowCount = 0;

            //取得Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                //計算勞健保
                //20210118 CCL-計數改由顯示回傳 l_iTolRowCount = l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                l_oRtnTolExcelLHIns.Fun_CalcAllLHInsResult(l_oDTItemData);
                //顯示輸出
                //l_iTolRowCount =  ExporeOnePartExcelListV1(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                    p_sPROG_ID, p_oServer);
                //l_iTolRowCount = ExporeOnePartExcelListV2(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                //                   p_sPROG_ID, p_oServer);
                //顯示明細
                ///l_iTolRowCount = ExporeOnePartExcelListV3(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                ///                   p_sPROG_ID, p_oServer);
                //顯示總和
                l_iTolRowCount = ExporeOnePartExcelListV4(l_iDataIndex, l_oRtnTolExcelLHIns, l_oWooksheet,
                                   p_sPROG_ID, p_oServer);


                l_iDataIndex++;
            }


            //20201230 CCL+ 重新排序
            ///l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();


            //20210103 CCL+ 計算所有店的合計
            ///l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ///ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ///ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
            ///                        l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////

            ///l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            //20210118 CCL- IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            //+2 ==> 1行空白Row,1行總和Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
            IXLStyle l_oStyle = l_oRangeAll.Style;
            l_oStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.BottomBorderColor = XLColor.Black;
            l_oStyle.Border.TopBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.TopBorderColor = XLColor.Black;
            l_oStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.LeftBorderColor = XLColor.Black;
            l_oStyle.Border.RightBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.RightBorderColor = XLColor.Black;
            l_oStyle.Border.InsideBorder = XLBorderStyleValues.Thin;
            l_oStyle.Border.InsideBorderColor = XLColor.Black;
            //l_oStyle.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
            //l_oStyle.Border.OutsideBorderColor = XLColor.Black;

            ////////////////////////////////////////////////////////////////////////



            //20210105 CCL+ 科目名稱改為標楷體
            //IXLRange l_oRangeAccNaCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iTolRowCount + 1, 1).Address);
            //IXLStyle l_oStyleAccNa = l_oRangeAccNaCol.Style;
            //l_oStyleAccNa.Font.FontName = "標楷體";


            //20210104 CCL+, 設定Cell寬度
            l_oWooksheet.Columns().AdjustToContents();
            //字會壓縮 l_oWooksheet.Rows().AdjustToContents();


            //20201227 CCL+ 儲存改放來這裡////////////////////////////////////////////
            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            try
            {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }


            SaveAsExcel(p_sPROG_ID, p_oServer);//
            //20201227 /////////////////////////////////////////////////////////////


            return false;


        }


        ////////////////////////////////////// 顯示公司及門市小計 /////////////////////////////////////////////

        public static int DisplayDepartTolAmount(int p_iRowIndex, int p_iComDeptIndex,
                                                 MERP_LHInsExcelItem Item,
                                                 string p_sPrevDeprtName,
                                                 string p_sPrevCompanyName,
                                                 MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                                 string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iComDeptIndex = p_iComDeptIndex;

            //如果公司,門市與前一筆不同,顯示[門市小計],[公司小計]
            //if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
            //{
                //l_oRtnExporeExcel.m_oLHInsExcelByComDepResults.Where(m => m.m_sPlusInsCompany == )
                MERP_LHInsExcelItem l_TmpTolItem =
                                    l_oRtnExporeExcel.m_oLHInsExcelByComDepResults[l_iComDeptIndex];
                l_iComDeptIndex++;


                //1.顯示 公司
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

                //2.顯示 門市
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = l_TmpTolItem.m_sDepartName;

                //3.顯示 姓名
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[門市小計]";

                //1.顯示 勞保-單位(含職災)
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
                //2.顯示 勞保-個人
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
                //3.顯示 健保-單位
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dComHealInsAmount;
                //4.顯示 健保-個人
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dPerHealInsAmount;
                //5.顯示 勞退-單位
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dComRetireInsAmount;
                //6.顯示 勞退
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                //7.顯示 合計-單位
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
                //8.顯示 合計-個人
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
                //9.顯示 總合計
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

                //勞退 的值的Cell要合併
                IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                l_oRetireInsTolCol.Merge();

                if(p_sDisplayType == "Detail")
                {
                    //門市小計 Item欄位填滿底色            
                    IXLRange l_oRangeDepartTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
                    l_oRangeDepartTolAmount.Style.Fill.SetBackgroundColor(XLColor.Yellow);

                }



                //++l_iRowIndex;
            //}


            return l_iComDeptIndex;
        }

        public static int DisplayCompanyTolAmount(int p_iRowIndex, int p_iCompanyIndex,
                                         MERP_LHInsExcelItem Item,                                         
                                         string p_sPrevCompanyName,
                                         MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                         string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            //string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iCompanyIndex = p_iCompanyIndex;

            MERP_LHInsExcelItem l_TmpTolItem =
                   l_oRtnExporeExcel.m_oLHInsExcelByComResults[l_iCompanyIndex];

            l_iCompanyIndex++;

            //1.顯示 公司
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

            //2.顯示 門市
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = "";

            //3.顯示 姓名
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[公司小計]";

            //1.顯示 勞保-單位(含職災)
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
            //2.顯示 勞保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
            //3.顯示 健保-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dComHealInsAmount;
            //4.顯示 健保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dPerHealInsAmount;
            //5.顯示 勞退-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dComRetireInsAmount;
            //6.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            //7.顯示 合計-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
            //8.顯示 合計-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
            //9.顯示 總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

            //勞退 的值的Cell要合併
            IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oRetireInsTolCol.Merge();

            //門市小計 Item欄位填滿底色            
            IXLRange l_oRangeCompanyTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeCompanyTolAmount.Style.Fill.SetBackgroundColor(XLColor.LightBlue);


            return l_iCompanyIndex;
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////////////

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加只顯示投保公司-門市總和 版本4       
        public static int ExporeOnePartExcelListV4(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Total";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    
                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmount(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        
                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmount(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        

                    }

                    //1.顯示 公司
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;
                    //2.顯示 門市
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;
                    //3.顯示 姓名
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;
                    //1.顯示 勞保-單位(含職災)
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 健保-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //4.顯示 健保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //5.顯示 勞退-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //6.顯示 勞退
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //7.顯示 合計-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //8.顯示 合計-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //9.顯示 總合計
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;
                    //勞退 的值的Cell要合併
                    //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                    //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    //l_oRetireInsCol.Merge();

                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmount(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmount(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        

                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }


        ////////////////////////////////////// 顯示公司及門市小計 /////////////////////////////////////////////

        public static int DisplayDepartTolAmountV3_1(int p_iRowIndex, int p_iComDeptIndex,
                                                 MERP_LHInsExcelItem Item,
                                                 string p_sPrevDeprtName,
                                                 string p_sPrevCompanyName,
                                                 MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                                 string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iComDeptIndex = p_iComDeptIndex;

            //如果公司,門市與前一筆不同,顯示[門市小計],[公司小計]
            //if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
            //{
            //l_oRtnExporeExcel.m_oLHInsExcelByComDepResults.Where(m => m.m_sPlusInsCompany == )
            MERP_LHInsExcelItem l_TmpTolItem =
                                l_oRtnExporeExcel.m_oLHInsExcelByComDepResults[l_iComDeptIndex];
            l_iComDeptIndex++;

            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolInsTolCol.Merge();

            //1.顯示 公司
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

            //2.顯示 門市
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = l_TmpTolItem.m_sDepartName;

            //3.顯示 姓名
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[門市小計]";

            //1.顯示 勞保-單位(含職災)
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
            //2.顯示 勞保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
            //3.顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dTolLabPerComInsAmount;
            //4.顯示 健保-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dComHealInsAmount;
            //5.顯示 健保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dPerHealInsAmount;
            //6.顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = l_TmpTolItem.m_dTolHealPerComInsAmount;
            //7.顯示 勞退-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dComRetireInsAmount;
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = "";
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = "";
            //9.顯示 合計-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
            //10.顯示 合計-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
            //10.顯示 合計-小計=總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            //11.顯示 總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                      l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRetireInsTolCol.Merge();

           

            if (p_sDisplayType == "Detail")
            {
                //門市小計 Item欄位填滿底色            
                IXLRange l_oRangeDepartTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
                l_oRangeDepartTolAmount.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            }



            //++l_iRowIndex;
            //}


            return l_iComDeptIndex;
        }

        public static int DisplayCompanyTolAmountV3_1(int p_iRowIndex, int p_iCompanyIndex,
                                         MERP_LHInsExcelItem Item,
                                         string p_sPrevCompanyName,
                                         MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                         string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            //string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iCompanyIndex = p_iCompanyIndex;

            MERP_LHInsExcelItem l_TmpTolItem =
                   l_oRtnExporeExcel.m_oLHInsExcelByComResults[l_iCompanyIndex];

            l_iCompanyIndex++;

            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolInsTolCol.Merge();

            //1.顯示 公司
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

            //2.顯示 門市
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = "";

            //3.顯示 姓名
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[公司小計]";

            //1.顯示 勞保-單位(含職災)
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
            //2.顯示 勞保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
            //3.顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dTolLabPerComInsAmount;
            //4.顯示 健保-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dComHealInsAmount;
            //5.顯示 健保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dPerHealInsAmount;
            //6.顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = l_TmpTolItem.m_dTolHealPerComInsAmount;
            //7.顯示 勞退-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dComRetireInsAmount;
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = "";
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = "";
            //9.顯示 合計-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
            //10.顯示 合計-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
            //10.顯示 合計-小計=總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            //11.顯示 總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRetireInsTolCol.Merge();

         

            //門市小計 Item欄位填滿底色            
            IXLRange l_oRangeCompanyTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeCompanyTolAmount.Style.Fill.SetBackgroundColor(XLColor.LightBlue);


            return l_iCompanyIndex;
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////////////



        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加顯示投保公司總和明細,增加勞健保小計行 版本3_1       
        public static int ExporeOnePartExcelListV3_1(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Detail";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    ++l_iRowIndex;
                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        l_iComDeptIndex = DisplayDepartTolAmountV3_1(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;
                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_1(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;

                    }

                    //合計 的後兩行值的Cell要合併
                    IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
                    l_oTolInsTolCol.Merge();

                    //1.顯示 公司
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;

                    //2.顯示 門市
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;

                    //3.顯示 姓名
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;

                    //1.顯示 勞保-單位(含職災)
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 勞保-個人+單位 小計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dTolLabPerComInsAmount;
                    //4.顯示 健保-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dComHealInsAmount;
                    //5.顯示 健保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dPerHealInsAmount;
                    //6.顯示 健保-個人+單位 小計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = Item.m_dTolHealPerComInsAmount;
                    //7.顯示 勞退-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dComRetireInsAmount;
                    //8.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = "";
                    //8.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = "";
                    //9.顯示 合計-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = Item.m_dTolComLHRInsAmount;
                    //10.顯示 合計-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = Item.m_dTolPerLHRInsAmount;
                    //10.顯示 合計-小計=總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
                    //11.顯示 總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = Item.m_dTolAllLHRInsAmount;

                    //勞退 的值的Cell要合併
                    //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                    //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
                    l_oRetireInsCol.Merge();

                    

                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_1(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_1(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 勞保-小計, 健保-單位, 健保-個人, 健保-小計, 勞退-單位 總合
            ++l_iRowIndex;
            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolPerInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolPerInsTolCol.Merge();

            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dPerComLaborInsTolAmounts;//顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = l_oRtnExporeExcel.m_dPerComHealInsTolAmounts;//顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = ""; //顯示 合計-小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            //IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }

        //////////////////////////// 20210129 CCL+ 新增墊償 ///////////////////////////////////////////////////
        ////////////////////////////////////// 顯示公司及門市小計 /////////////////////////////////////////////
        public static int DisplayDepartTolAmountV3_2(int p_iRowIndex, int p_iComDeptIndex,
                                                 MERP_LHInsExcelItem Item,
                                                 string p_sPrevDeprtName,
                                                 string p_sPrevCompanyName,
                                                 MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                                 string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            //const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iComDeptIndex = p_iComDeptIndex;

            //如果公司,門市與前一筆不同,顯示[門市小計],[公司小計]
            //if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
            //{
            //l_oRtnExporeExcel.m_oLHInsExcelByComDepResults.Where(m => m.m_sPlusInsCompany == )
            MERP_LHInsExcelItem l_TmpTolItem =
                                l_oRtnExporeExcel.m_oLHInsExcelByComDepResults[l_iComDeptIndex];
            l_iComDeptIndex++;

            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolInsTolCol.Merge();

            //墊償 的後兩行值的Cell要合併
            IXLRange l_oTolFundTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oTolFundTolCol.Merge();


            //1.顯示 公司
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

            //2.顯示 門市
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = l_TmpTolItem.m_sDepartName;

            //3.顯示 姓名
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[門市小計]";

            //1.顯示 勞保-單位(含職災)
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
            //2.顯示 勞保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
            //3.顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dTolLabPerComInsAmount;
            //4.顯示 墊償
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dComLaborFundAmount; //20210225 CCL*
            //5.顯示 墊償
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dTolLabComFundAmount; //20210225 CCL*
            //6.顯示 墊償 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            //4.顯示 健保-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dComHealInsAmount;
            //5.顯示 健保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_TmpTolItem.m_dPerHealInsAmount;
            //6.顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_TmpTolItem.m_dTolHealPerComInsAmount;
            //7.顯示 勞退-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_TmpTolItem.m_dComRetireInsAmount;
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = "";
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            //9.顯示 合計-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
            //10.顯示 合計-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 14).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
            //10.顯示 合計-小計=總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 15).Value = "";
            //11.顯示 總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 16).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                      l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRetireInsTolCol.Merge();



            if (p_sDisplayType == "Detail")
            {
                //門市小計 Item欄位填滿底色            
                IXLRange l_oRangeDepartTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
                l_oRangeDepartTolAmount.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            }



            //++l_iRowIndex;
            //}


            return l_iComDeptIndex;
        }

        public static int DisplayCompanyTolAmountV3_2(int p_iRowIndex, int p_iCompanyIndex,
                                         MERP_LHInsExcelItem Item,
                                         string p_sPrevCompanyName,
                                         MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                         string p_sDisplayType)
        {
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行                       
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            //const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;

            //string l_sPrevDeprtName = p_sPrevDeprtName;
            string l_sPrevCompanyName = p_sPrevCompanyName;
            int l_iRowIndex = p_iRowIndex;
            int l_iCompanyIndex = p_iCompanyIndex;

            MERP_LHInsExcelItem l_TmpTolItem =
                   l_oRtnExporeExcel.m_oLHInsExcelByComResults[l_iCompanyIndex];

            l_iCompanyIndex++;

            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolInsTolCol.Merge();

            //墊償 的後兩行值的Cell要合併
            IXLRange l_oTolFundTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oTolFundTolCol.Merge();


            //1.顯示 公司
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = l_TmpTolItem.m_sPlusInsCompany;

            //2.顯示 門市
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = "";

            //3.顯示 姓名
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "[公司小計]";

            //1.顯示 勞保-單位(含職災)
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_TmpTolItem.m_dComTolLaborODInsAmount;
            //2.顯示 勞保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_TmpTolItem.m_dPerLaborInsAmount;
            //3.顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_TmpTolItem.m_dTolLabPerComInsAmount;
            //4.顯示 墊償
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_TmpTolItem.m_dComLaborFundAmount;
            //5.顯示 墊償
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_TmpTolItem.m_dTolLabComFundAmount;
            //6.顯示 墊償 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            //4.顯示 健保-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_TmpTolItem.m_dComHealInsAmount;
            //5.顯示 健保-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_TmpTolItem.m_dPerHealInsAmount;
            //6.顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_TmpTolItem.m_dTolHealPerComInsAmount;
            //7.顯示 勞退-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_TmpTolItem.m_dComRetireInsAmount;
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = "";
            //8.顯示 勞退
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            //9.顯示 合計-單位
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_TmpTolItem.m_dTolComLHRInsAmount;
            //10.顯示 合計-個人
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 14).Value = l_TmpTolItem.m_dTolPerLHRInsAmount;
            //10.顯示 合計-小計=總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 15).Value = "";
            //11.顯示 總合計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 16).Value = l_TmpTolItem.m_dTolAllLHRInsAmount;

            //勞退 的值的Cell要合併
            //IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
            //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oRetireInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oRetireInsTolCol.Merge();



            //門市小計 Item欄位填滿底色            
            IXLRange l_oRangeCompanyTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeCompanyTolAmount.Style.Fill.SetBackgroundColor(XLColor.LightBlue);


            return l_iCompanyIndex;
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////////////



        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加顯示投保公司總和明細,增加勞健保小計行,墊償 版本3_2       
        public static int ExporeOnePartExcelListV3_2(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            //const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Detail";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    ++l_iRowIndex;
                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        l_iComDeptIndex = DisplayDepartTolAmountV3_2(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;
                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_2(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;

                    }

                    //合計 的後兩行值的Cell要合併
                    IXLRange l_oTolInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
                    l_oTolInsTolCol.Merge();

                    //墊償 的後兩行值的Cell要合併
                    IXLRange l_oTolFundTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
                    l_oTolFundTolCol.Merge();

                    //1.顯示 公司
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;

                    //2.顯示 門市
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;

                    //3.顯示 姓名
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;

                    //1.顯示 勞保-單位(含職災)
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 勞保-個人+單位 小計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dTolLabPerComInsAmount;
                    //4.顯示 墊償
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = "";
                    //5.顯示 墊償
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = "";
                    //6.顯示 墊償 小計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //4.顯示 健保-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dComHealInsAmount;
                    //5.顯示 健保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dPerHealInsAmount;
                    //6.顯示 健保-個人+單位 小計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolHealPerComInsAmount;
                    //7.顯示 勞退-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = Item.m_dComRetireInsAmount;
                    //8.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = "";
                    //8.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
                    //9.顯示 合計-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = Item.m_dTolComLHRInsAmount;
                    //10.顯示 合計-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 14).Value = Item.m_dTolPerLHRInsAmount;
                    //10.顯示 合計-小計=總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 15).Value = "";
                    //11.顯示 總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 16).Value = Item.m_dTolAllLHRInsAmount;

                    //勞退 的值的Cell要合併
                    //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                    //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
                    l_oRetireInsCol.Merge();



                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_2(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_2(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 勞保-小計, 健保-單位, 健保-個人, 健保-小計, 勞退-單位 總合
            ++l_iRowIndex;
            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolPerInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolPerInsTolCol.Merge();

            //墊償 的後兩行值的Cell要合併
            IXLRange l_oTolFundInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oTolFundInsTolCol.Merge();

            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dPerComLaborInsTolAmounts;//顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dLaborInsFundTolAmounts; //所有墊償 值相加
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComTolLInsFundTolAmounts; //所有墊償 小計 值相加 
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dPerComHealInsTolAmounts;//顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 14).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 15).Value = ""; //顯示 合計-小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 16).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            //IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //20210201 CCL+ 增加小計, 增加墊償
        //增加只顯示投保公司-門市總和 版本4_2       
        public static int ExporeOnePartExcelListV4_2(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            //const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int PRINT_PART = 5; //勞保,墊償,健保,勞退,合計 五組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Total";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {

                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_2(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);


                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_2(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);



                    }

                    //1.顯示 公司
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;
                    //2.顯示 門市
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;
                    //3.顯示 姓名
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;
                    //1.顯示 勞保-單位(含職災)
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 健保-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //4.顯示 健保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //5.顯示 勞退-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //6.顯示 勞退
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //7.顯示 合計-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //8.顯示 合計-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //9.顯示 總合計
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;
                    //勞退 的值的Cell要合併
                    //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                    //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    //l_oRetireInsCol.Merge();

                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_2(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_2(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);


                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolPerInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolPerInsTolCol.Merge();
            //墊償 的後兩行值的Cell要合併
            IXLRange l_oTolFundTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 4)) + 3).Address);
            l_oTolFundTolCol.Merge();



            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dPerComLaborInsTolAmounts;//顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dLaborInsFundTolAmounts; //所有墊償 值相加
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComTolLInsFundTolAmounts; //所有墊償 小計 值相加 
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dPerComHealInsTolAmounts;//顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 14).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 15).Value = ""; //顯示 合計-小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 16).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            //IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //20210126 CCL+ 增加小計
        //增加只顯示投保公司-門市總和 版本4_1       
        public static int ExporeOnePartExcelListV4_1(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            //const int PRINT_COLS = 2; //單位,個人
            const int PRINT_COLS = 3; //單位,個人,小計
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Total";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {

                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_1(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);


                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_1(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);



                    }

                    //1.顯示 公司
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;
                    //2.顯示 門市
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;
                    //3.顯示 姓名
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;
                    //1.顯示 勞保-單位(含職災)
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 健保-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //4.顯示 健保-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //5.顯示 勞退-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //6.顯示 勞退
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //7.顯示 合計-單位
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //8.顯示 合計-個人
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //9.顯示 總合計
                    //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;
                    //勞退 的值的Cell要合併
                    //IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                    //                           l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    //l_oRetireInsCol.Merge();

                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmountV3_1(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmountV3_1(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);


                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            //合計 的後兩行值的Cell要合併
            IXLRange l_oTolPerInsTolCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 2).Address,
                                       l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 1)) + 3).Address);
            l_oTolPerInsTolCol.Merge();

            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dPerComLaborInsTolAmounts;//顯示 勞保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = l_oRtnExporeExcel.m_dPerComHealInsTolAmounts;//顯示 健保-個人+單位 小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 10).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 11).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 12).Value = ""; //顯示 合計-小計
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 13).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            //IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 3).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }



        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加顯示投保公司總和明細 版本3       
        public static int ExporeOnePartExcelListV3(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";
            string l_sPrevDeprtName = "";

            string l_sDisplayType = "Detail";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iComDeptIndex = 0;
            int l_iCompanyIndex = 0;
            int l_iOrgRowCount = 0;

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    ++l_iRowIndex;
                    l_iOrgRowCount++;

                    //如果公司,門市與前一筆不同,顯示[門市小計]
                    if ((Item.m_sDepartName != l_sPrevDeprtName) && (l_sPrevDeprtName != ""))
                    {
                        l_iComDeptIndex = DisplayDepartTolAmount(l_iRowIndex, l_iComDeptIndex, Item,
                                           l_sPrevDeprtName, l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;
                    }


                    //20210120 CCL+ 如果公司與前一筆不同,顯示[公司小計]
                    if ((Item.m_sPlusInsCompany != l_sPrevCompanyName) && (l_sPrevCompanyName != ""))
                    {

                        l_iCompanyIndex = DisplayCompanyTolAmount(l_iRowIndex, l_iCompanyIndex, Item,
                                           l_sPrevCompanyName,
                                           l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);

                        ++l_iRowIndex;

                    }

                    //1.顯示 公司
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;

                    //2.顯示 門市
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = Item.m_sDepartName;

                    //3.顯示 姓名
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = Item.m_sMemberName;

                    //1.顯示 勞保-單位(含職災)
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 健保-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //4.顯示 健保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //5.顯示 勞退-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //6.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //7.顯示 合計-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //8.顯示 合計-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //9.顯示 總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;

                    //勞退 的值的Cell要合併
                    IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    l_oRetireInsCol.Merge();


                    //更新公司-門市
                    l_sPrevCompanyName = Item.m_sPlusInsCompany;
                    l_sPrevDeprtName = Item.m_sDepartName;

                    ///////////////////// 如果最後一筆 也要顯示 //////////////////////////////////////////////////////////
                    if (l_iOrgRowCount == l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count())
                    {
                        //最後一筆人要先Index累加
                        ++l_iRowIndex;

                        l_iComDeptIndex = DisplayDepartTolAmount(l_iRowIndex, l_iComDeptIndex, Item,
                                          l_sPrevDeprtName, l_sPrevCompanyName,
                                          l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                        l_iCompanyIndex = DisplayCompanyTolAmount(l_iRowIndex, l_iCompanyIndex, Item,
                                        l_sPrevCompanyName,
                                        l_oRtnExporeExcel, l_oWooksheet, l_sDisplayType);
                        ++l_iRowIndex;

                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////
                   

                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);


            return l_iRowIndex;
        }


        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加只顯示投保公司總和 版本2       
        public static int ExporeOnePartExcelListV2(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
          

            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelByComResults.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelByComResults)
                {
                    ++l_iRowIndex;

                    //1.顯示 公司
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;                    

                    //2.顯示 門市
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2).Value = "";

                    //3.顯示 姓名
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3).Value = "";

                    //1.顯示 勞保-單位(含職災)
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = Item.m_dComTolLaborODInsAmount;
                    //2.顯示 勞保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = Item.m_dPerLaborInsAmount;
                    //3.顯示 健保-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = Item.m_dComHealInsAmount;
                    //4.顯示 健保-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = Item.m_dPerHealInsAmount;
                    //5.顯示 勞退-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = Item.m_dComRetireInsAmount;
                    //6.顯示 勞退
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
                    //7.顯示 合計-單位
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = Item.m_dTolComLHRInsAmount;
                    //8.顯示 合計-個人
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = Item.m_dTolPerLHRInsAmount;
                    //9.顯示 總合計
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = Item.m_dTolAllLHRInsAmount;

                    //勞退 的值的Cell要合併
                    IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                               l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                    l_oRetireInsCol.Merge();


                }

            }


            //18.顯示 空白
            ++l_iRowIndex;          
            for (int i = 1; i <= 8; i++)
            {
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";               
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色            
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iRowIndex, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oTolRetireInsCol.Merge();


            //Styling                     
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;


            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);

            
            return l_iRowIndex;
        }


        /*
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////
        //增加只顯示投保公司總和 版本       
        public static int ExporeOnePartExcelListV1(int p_iPartIndex,
                           MERP_LHInsExcelExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {

            int l_iPartCount = 1;
            const int TOPCOLS = 3; //顯示 公司 門市 姓名 3行           
            const string TABLENAME = "勞健保表";
            const int PRINT_COLS = 2; //單位,個人
            const int PRINT_PART = 4; //勞保,健保,勞退,合計 四組
            const int TOLCOUNT_COL = 1; //員工勞健退保合計 1行
            const int HEAD_ROWS = 4; //新版表頭數目

            string l_sPrevCompanyName = "";

            //一組Part用到顯示幾行Column, l_iPartCount以0為起始        
            //int l_iPadColIndex = PRINT_COLS * l_iPartCount;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_LHInsExcelExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            decimal l_dComTolLaborODInsAmount = 0;
            decimal l_dPerLaborInsAmount = 0;
            decimal l_dComHealInsAmount = 0;
            decimal l_dPerHealInsAmount = 0;
            decimal l_dComRetireInsAmount = 0;
            decimal l_dTolComLHRInsAmount = 0;
            decimal l_dTolPerLHRInsAmount = 0;
            decimal l_dTolAllLHRInsAmount = 0;
            bool l_bIsPrintOut = false;
            int l_iPrintOutCount = 0;


            if ((l_oRtnExporeExcel != null) &&
                (l_oRtnExporeExcel.m_oLHInsExcelResultItems.Count() > 0))
            {
                foreach (MERP_LHInsExcelItem Item in l_oRtnExporeExcel.m_oLHInsExcelResultItems)
                {
                    ++l_iRowIndex;
        
                    if(l_sPrevCompanyName == "")
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
                    else if(Item.m_sPlusInsCompany == l_sPrevCompanyName )
                    {
                        //累加不顯示
                        l_dComTolLaborODInsAmount += Item.m_dComTolLaborODInsAmount;
                        l_dPerLaborInsAmount += Item.m_dPerLaborInsAmount;
                        l_dComHealInsAmount += Item.m_dComHealInsAmount;
                        l_dPerHealInsAmount += Item.m_dPerHealInsAmount;
                        l_dComRetireInsAmount += Item.m_dComRetireInsAmount;
                        l_dTolComLHRInsAmount += Item.m_dTolComLHRInsAmount;
                        l_dTolPerLHRInsAmount += Item.m_dTolPerLHRInsAmount;
                        l_dTolAllLHRInsAmount += Item.m_dTolAllLHRInsAmount;

                        l_bIsPrintOut = false;

                    } else if(Item.m_sPlusInsCompany != l_sPrevCompanyName)
                    {

                        l_bIsPrintOut = true;
           
                    }

                    if(l_bIsPrintOut == true)
                    {
                        ++l_iPrintOutCount;

                        //1.顯示 公司
                        //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Item.m_sPlusInsCompany;

                        //顯示上一間公司總計的資料
                        //1.顯示 公司
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, 1).Value = l_sPrevCompanyName;

                        //2.顯示 門市
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, 2).Value = "";

                        //3.顯示 姓名
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, 3).Value = "";

                        //1.顯示 勞保-單位(含職災)
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 1).Value = l_dComTolLaborODInsAmount;
                        //2.顯示 勞保-個人
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 2).Value = l_dPerLaborInsAmount;
                        //3.顯示 健保-單位
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 3).Value = l_dComHealInsAmount;
                        //4.顯示 健保-個人
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 4).Value = l_dPerHealInsAmount;
                        //5.顯示 勞退-單位
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 5).Value = l_dComRetireInsAmount;
                        //6.顯示 勞退
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 6).Value = "";
                        //7.顯示 合計-單位
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 7).Value = l_dTolComLHRInsAmount;
                        //8.顯示 合計-個人
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 8).Value = l_dTolPerLHRInsAmount;
                        //9.顯示 總合計
                        l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 9).Value = l_dTolAllLHRInsAmount;

                        //勞退 的值的Cell要合併
                        IXLRange l_oRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address,
                                                   l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
                        l_oRetireInsCol.Merge();

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
                        l_bIsPrintOut = false;

                    }


                    

                }

            }


            //18.顯示 空白
            ++l_iRowIndex;
            ++l_iPrintOutCount;
            for (int i = 1; i <= 8; i++)
            {
                //l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, i).Value = "";
                l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, i).Value = "";
            }
            //空白 的值的Cell要合併
            IXLRange l_oSpaceCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oSpaceCol.Merge();

            //19.顯示 勞保-單位, 勞保-個人, 健保-單位, 健保-個人, 勞退-單位 總合
            ++l_iRowIndex;
            ++l_iPrintOutCount;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 1).Value = l_oRtnExporeExcel.m_dComLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 2).Value = l_oRtnExporeExcel.m_dPerLaborInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 3).Value = l_oRtnExporeExcel.m_dComHealInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 4).Value = l_oRtnExporeExcel.m_dPerHealInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 5).Value = l_oRtnExporeExcel.m_dComRetireInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 6).Value = "";
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 7).Value = l_oRtnExporeExcel.m_dComLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 8).Value = l_oRtnExporeExcel.m_dPerLHRInsTolAmounts;
            l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + 9).Value = l_oRtnExporeExcel.m_dAllLHRInsTolAmounts;
            //現金總和 Item欄位填滿底色
            //IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Address);
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //勞退 現金總和 的值的Cell要合併
            IXLRange l_oTolRetireInsCol = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 1).Address, l_oWooksheet.Cell(HEAD_ROWS + l_iPrintOutCount, TOPCOLS + (PRINT_COLS * (PRINT_PART - 2)) + 2).Address);
            l_oTolRetireInsCol.Merge();


            //Styling          
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 6).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iPrintOutCount + HEAD_ROWS, TOPCOLS + (PRINT_COLS * PRINT_PART) + TOLCOUNT_COL).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;



            //定義資料使用到的Range範圍 起始/結束Cell
            var l_oFirstCell = l_oWooksheet.FirstCellUsed();
            var l_oEndCell = l_oWooksheet.LastCellUsed();
            //使用資料起始/結束 Cell, 來定義一個資料範圍 (利用Cell Address來定位)
            var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);


            //將資料範圍轉型
            try
            {
                m_oModTable = l_oData.AsTable();
            }
            catch (Exception ex)
            {
                string message = ex.Message.ToString();
            }



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_dTolLaborInsVal);

            //return true;
            return l_iPrintOutCount;
        }
        */

    }
}