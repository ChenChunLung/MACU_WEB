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

namespace MACU_WEB.BIServices
{
    public class MERP_ExcelV1BIService
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
        public static MERP_FA_FaJournalV1DBService m_FaJournalV1DBService = new MERP_FA_FaJournalV1DBService();
        public static MERP_AccountInfoDBService m_AccInfoDBService = new MERP_AccountInfoDBService();

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

        public static Boolean Process_DelAllZeroRowsExcel()
        {
            if (m_oImpTable != null)
            {
                int l_iAllCellVal_TolCount = 0;
                int l_iRowIndex = 0;
                int l_iCellIndex = 0;
                int l_ModRowCount = 0;
                bool l_bAllCellIsZero = false;
                int l_iAllCellVal_RowCount = 0;
                string l_sToDelIndexStr = "";

                int l_iRowCount = 0;
                l_iRowCount = m_oImpTable.RowCount();
                //for(int i=1; i< p_oImpTable.RowCount(); i++)
                //{

                //}

                //Loop through the Worksheet rows.
                foreach (IXLRangeRow row in m_oImpTable.Rows())
                {
                    //從第9列開始算,1~8列是其他訊息
                    l_iAllCellVal_TolCount = 0;
                    l_bAllCellIsZero = false;
                    ++l_iRowIndex; //Row從第9列算


                    if (l_iRowIndex >= 9)
                    {
                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            ++l_iCellIndex; //Cell從第3行算
                            if (cell.Value.ToString() != null)
                            {
                                if (l_iCellIndex > 2)
                                {
                                    //if(cell.Value.ToString() == "0")
                                    //{
                                    //    l_iAllCellVal_TolCount++;
                                    //}

                                    //測試後,必須用ToInt32轉成Int相加才准,用上面判斷"0"字串的方式會不準,有的沒刪除到
                                    l_iAllCellVal_TolCount += Convert.ToInt32(cell.Value.ToString());

                                }
                            }

                        }

                        /*
                        if((l_iAllCellVal_TolCount + 2) == row.CellCount())
                        {
                            l_bAllCellIsZero = true;
                            l_iAllCellVal_RowCount++;
                        } else
                        {
                            l_bAllCellIsZero = false;
                        }
                        */

                        l_iCellIndex = 0;

                        //if (l_bAllCellIsZero)
                        if (l_iAllCellVal_TolCount == 0)
                        {
                            //組合出要刪除的Row Index字串 Ex: "4:6" "3:5,7:8" "12" "5,7,9,16"
                            l_sToDelIndexStr += l_iRowIndex.ToString() + ",";

                            //m_oModTable.Rows().Delete();
                        }


                    }
                }

                if (l_sToDelIndexStr != "")
                {
                    l_sToDelIndexStr = l_sToDelIndexStr.Remove(l_sToDelIndexStr.Length - 1, 1); //去除最後的","
                    m_oModTable.Rows(l_sToDelIndexStr).Delete();
                }

                l_ModRowCount = m_oModTable.RowCount();
                //l_iAllCellVal_TolCount = l_iAllCellVal_RowCount; //for debug
                //l_ModRowCount = m_oModWorkbook.Worksheet(1).RowsUsed().Count();
                return true;
            }
            return false;
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


                //l_oFiles.SaveAs(l_sFilePath + "/" + name + type);//保存文件

                string l_sUrl = l_sFilePath + "\\" + name + type;


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



        //20201223 CCL+ V1版本Excel//////////////////////////////////////////////////////////////////
        /*
        public static string ImportExcelToFA_DayContentDB_V1(String p_sFullFilePath)
        {
            //載入的會計期別
            string l_sAccountPeriod = "";

            //載入上傳的Excel
            ImportExcelCommon(p_sFullFilePath);
            //判斷這個月會計期別是否已在DB有資料,有的話,刪除舊的
            if (m_oImpTable != null)
            {
                //從上傳Excel中抓出AccountPeriod
                //根據第一列的傳票日期,取出年,月
                int l_iYear = 0, l_iMonth = 0;
                string l_sTmpStr = DateStringProcess.Del_MonthDayZero(m_oImpTable.Cell(2, 1).Value.ToString(), "/", "");
                l_iYear = DateStringProcess.m_Year;
                l_iMonth = DateStringProcess.m_Month;
                l_sAccountPeriod = l_iMonth.ToString();
                //找出資料庫是否有本月
                Boolean l_bIsExistData = m_FaJournalV1DBService.FA_JournalV1_FindDataByMonthVal(l_sAccountPeriod.Trim());
                if (l_bIsExistData)
                {
                    //刪除舊的
                    //m_FaJournalDBService.FA_FaJournal_DBDeleteByPeriod(l_sAccountPeriod.Trim());
                    m_FaJournalV1DBService.FA_JournalV1_SqlDBDeleteByPeriod(l_sAccountPeriod.Trim());
                }
                //,並且匯入DataBase
                //20201217 CCL- m_FaJournalDBService.FA_FaJournal_DBCreate(m_oImpTable);
                //20201229 CCL-貸借金額順序改變 m_FaJournalV1DBService.FA_JournalV1_SqlDBCreate(m_oImpTable); //改用ADO.NET提升速度
                m_FaJournalV1DBService.FA_JournalV1_SqlDBCreateV1(m_oImpTable); //改用ADO.NET提升速度
            }
            //return true;
            return l_sAccountPeriod;
        }

        
        public static List<FA_JournalV1> GetImportExcelInDB_PeriodData(string p_sVal)
        {
            
            return m_FaJournalV1DBService.FA_JournalV1_GetDataByMonthVal(p_sVal);
        }

        public static List<FA_JournalV1> GetImportExcelInDB_PeriodDataPage(string p_sVal, int p_iPageing)
        {
            
            return m_FaJournalV1DBService.FA_JournalV1_GetDataByMonthValPage(p_sVal, p_iPageing);
        }
        */
        // ///////////////////////////////////////////////////////////////////////////////////////////

        //20210204 CCl+, 修正以抓出年月為主 //////////////////////////////////////////////////////////
        public static string[] ImportExcelToFA_DayContentDB_V1(String p_sFullFilePath)
        {
            //載入的會計期別
            string l_sAccountPeriod = "";
            //載入的會計年分
            string l_sFiscalYear = "";

            //載入上傳的Excel
            ImportExcelCommon(p_sFullFilePath);
            //判斷這個月會計期別是否已在DB有資料,有的話,刪除舊的
            if (m_oImpTable != null)
            {
                //從上傳Excel中抓出AccountPeriod
                //根據第一列的傳票日期,取出年,月
                int l_iYear = 0, l_iMonth = 0;
                string l_sTmpStr = DateStringProcess.Del_MonthDayZero(m_oImpTable.Cell(2, 1).Value.ToString(), "/", "");
                l_iYear = DateStringProcess.m_Year;
                l_iMonth = DateStringProcess.m_Month;
                l_sAccountPeriod = l_iMonth.ToString();
                l_sFiscalYear = l_iYear.ToString();
                //找出資料庫是否有本年月
                Boolean l_bIsExistData = m_FaJournalV1DBService.FA_JournalV1_FindDataByYearMonthVal(l_sFiscalYear.Trim(),
                                                                                            l_sAccountPeriod.Trim());
                if (l_bIsExistData)
                {
                    //刪除舊的                    
                    m_FaJournalV1DBService.FA_JournalV1_SqlDBDeleteByYearPeriod(l_sFiscalYear.Trim(), 
                                                                                l_sAccountPeriod.Trim());
                }
                //,並且匯入DataBase
                //20201217 CCL- m_FaJournalDBService.FA_FaJournal_DBCreate(m_oImpTable);
                //20201229 CCL-貸借金額順序改變 m_FaJournalV1DBService.FA_JournalV1_SqlDBCreate(m_oImpTable); //改用ADO.NET提升速度
                m_FaJournalV1DBService.FA_JournalV1_SqlDBCreateV1(m_oImpTable); //改用ADO.NET提升速度
            }
            
            //return l_sAccountPeriod;
            return new string[] { l_sFiscalYear, l_sAccountPeriod} ;
        }

        public static List<FA_JournalV1> GetImportExcelInDB_YearPeriodData(string p_Year, string p_sMonth)
        {

            return m_FaJournalV1DBService.FA_JournalV1_GetDataByYearMonthVal(p_Year, p_sMonth);
        }

        public static List<FA_JournalV1> GetImportExcelInDB_YearPeriodDataPage(string p_Year, string p_sMonth, int p_iPageing)
        {
            return m_FaJournalV1DBService.FA_JournalV1_GetDataByYearMonthValPage(p_Year, p_sMonth, p_iPageing);
            
        }
        /// //////////////////////////////////////////////////////////////////////////////////////////

        //20201218 CCL+ For Processing 區間日期Excel 商業logical ////////////////////////////////////
        public static List<FA_JournalV1> TransDataTableToList(DataSet p_oDataSet)
        {
            return m_FaJournalV1DBService.FA_JournalV1_DataTableTo_FaJournalsList(p_oDataSet);
            //FA_FaJournal_DataTableTo_FaJournalsList
        }

        public static List<FA_JournalV1> ProcessImportExcelFromDB(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oDataSet = null;
            List<FA_JournalV1> l_RtnList = null;

            if (p_oOption != null)
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions(p_oOption);
                l_RtnList = m_FaJournalV1DBService.FA_JournalV1_DataTableTo_FaJournalsList(l_oDataSet);
                return l_RtnList;
            }


            return null;

        }

        //20201224 CCL+
        /*
        //從AccountSubjects中分離出AccountNo, DetailAccountNo, AccountName
        public static TmpExcelItem ProcessAccountSubjects(string p_sAccSubjects)
        {
            TmpExcelItem l_oTmpItem = new TmpExcelItem();
            string l_sTmpStr = "";
            //string l_sAccountNo = "", l_sDetailAccountNo = "", l_sAccountName = "";

            //如果空格前面只有4碼
            l_sTmpStr = p_sAccSubjects.Substring(0, p_sAccSubjects.IndexOf(" "));
            //4碼
            l_oTmpItem.m_ComAccountNo = p_sAccSubjects.Substring(0, 4);
            if (l_sTmpStr.Length > 4)
            {
                //如果空格前面大於4碼,便要取出AccountNo
                l_oTmpItem.m_ComDetailAccNo = p_sAccSubjects.Substring(4, l_sTmpStr.Length - 4);
            }
            l_oTmpItem.m_ComAccName = p_sAccSubjects.Substring(p_sAccSubjects.IndexOf(" ") + 1).ToString();


            return l_oTmpItem;
        }
        */

        //全部項目印出
        //20201221 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions(MERP_ProcessExcelOptions p_oOption,
                                                String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //把Mod的砍掉Row只留下1,2,3,4,5,6,7,8; Col只留下1[科目代碼],2[科目名稱],3[店名,實績]
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 2;

            DataSet l_oDataSet = new DataSet();
            m_oWorkbook = new XLWorkbook();
            l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);
            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add("損益表");


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, 1).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            l_oRange6.Style.Font.FontName = "微軟正黑體";
            l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;

            //202006月帳
            l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            l_oWooksheet.Cell(2, 1).Value = "多部門損益表";


            //起迄日期：2020-07-01 ~ 2020-07-31
            l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            //科目代號
            l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            l_oWooksheet.Cell(8, 3).Value = "實績";

            IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            l_oRange2.Merge();
            IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            l_oRange3.Merge();

            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;
                //IXLCells cells = l_oWooksheet.Rows(l_iRowIndex.ToString()).Cells();
                //foreach(IXLCell cell in cells)
                //{
                //cell.Value = 
                //}

                //科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo") + "-" +
                        l_sDetailAccountNo;
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo");
                }

                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.Field<string>("SubjectName") + " " +
                    row.Field<string>("DetailSubjectName");


                if (l_iRowIndex == 1)
                {

                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }


                //借方金額
                if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("DebitAmount");
                }
                else
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("CreditAmount");
                }




            }

            //Styling
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
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

            //複製一份要修改的副本WorkSheet到新的ModWookbook 才能Save
            m_oModWorkbook = new XLWorkbook();
            m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);


            /*
            if (m_oModTable != null)
            {
                
                int l_iRowIndex = 0, l_iColIndex = 0;
                int l_iCellIndex = 0;
                int l_ModRowCount = 0;
                             
                string l_sToDelIndexStr = "";

                int l_iRowCount = 0;
                l_iRowCount = m_oImpTable.RowCount();

                //Loop through the Worksheet cols.
                //foreach(IXLRangeColumn col in m_oModTable.Columns())
                //{
                //如果超過三col,刪除
                //    ++l_iColIndex;
                //    if (l_iColIndex <= 3) continue;

                //}

                //col4以後全刪
                l_sToDelIndexStr = "4:" + m_oModTable.Columns().Count();
                m_oModTable.Columns(l_sToDelIndexStr).Delete();
                //第7列,第3行的店名改掉
                m_oModTable.Cell(7,3).Value = "麻古大安信義";
                string l_sStartEndDate = string.Format("起迄日期：{0} ~ {1}",
                                                       p_oOption.m_sStartDate,
                                                       p_oOption.m_sEndDate);
                m_oModTable.Cell(4, 1).Value = l_sStartEndDate;


                //Loop through the Worksheet rows.
                foreach (IXLRangeRow row in m_oImpTable.Rows())
                {
                    //Row4是起訖日期,修改成我們要的區間
                    //Row4是部門改成我們要的店名

                    //從第9列開始清為0,1~8列是其他訊息

                    ++l_iRowIndex; //Row從第9列算


                    if (l_iRowIndex >= 9)
                    {
                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            ++l_iCellIndex; //Cell從第3行算
                            if (!string.IsNullOrEmpty( cell.Value.ToString() ))
                            {
                                if (l_iCellIndex > 2)
                                {
                                    cell.Value = "0";
                                }
                            }

                        }

                     
                        l_iCellIndex = 0;

                        //組合出要刪除的Row Index字串 Ex: "4:6" "3:5,7:8" "12" "5,7,9,16"
                        //l_sToDelIndexStr += l_iRowIndex.ToString() + ",";
                          

                    }
                }

                //if (l_sToDelIndexStr != "")
                //{
                //    l_sToDelIndexStr = l_sToDelIndexStr.Remove(l_sToDelIndexStr.Length - 1, 1); //去除最後的","
                //    m_oModTable.Rows(l_sToDelIndexStr).Delete();
                //}

                l_ModRowCount = m_oModTable.RowCount();

                SaveAsExcel(p_sPROG_ID, p_oServer);//
                return true;
            }
            */

            SaveAsExcel(p_sPROG_ID, p_oServer);//

            return false;
        }

        /////////////////////////////////////////////////////////////////////////////////////////////

        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201224 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions5(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //把Mod的砍掉Row只留下1,2,3,4,5,6,7,8; Col只留下1[科目代碼],2[科目名稱],3[店名,實績]
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 2;
            const string TABLENAME = "損益表";

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();


            m_oWorkbook = new XLWorkbook();
            //找出該店且區間的Data
            l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);
            //抓出AccInfo Map Table全部Data
            

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, 1).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            l_oRange6.Style.Font.FontName = "微軟正黑體";
            l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;

            //202006月帳
            l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            l_oWooksheet.Cell(2, 1).Value = "多部門損益表";


            //起迄日期：2020-07-01 ~ 2020-07-31
            l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            //科目代號
            l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            l_oWooksheet.Cell(8, 3).Value = "實績";

            IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            l_oRange2.Merge();
            IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            l_oRange3.Merge();

            //最終Expore要輸出的打包物件
            MERP_AccInfoExpore l_oRtnExporeExcel = new MERP_AccInfoExpore();
            l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID

            TmpExcelItem l_oTmpItem = null;

            /*
            //Save To Expore Object
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;
                //Tmp 
                l_oTmpItem = new TmpExcelItem();

                //如果該row的AccNo,DtlAccNo不在AccInfo Map內,不存入 輸出打包物件
                //找出並儲存AccInfo 資訊Item
                l_oTmpAccInfoItem = m_AccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo(row.Field<string>("AccountNo"),
                                                                                          row.Field<string>("DetailAccountNo"));

                if(l_oTmpAccInfoItem != null)
                {
                    //科目代號
                    l_oTmpItem.m_ComAccountNo = row.Field<string>("AccountNo");
                    //明細科目代號
                    l_oTmpItem.m_ComDetailAccNo = row.Field<string>("DetailAccountNo");
                    //代碼全名
                    l_oTmpItem.m_FullNo = l_oTmpItem.m_ComAccountNo + "-" + l_oTmpItem.m_ComDetailAccNo;

                    //科目名稱
                    l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName");
                    //明細科目名稱
                    l_oTmpItem.m_ComDetailAccName = row.Field<string>("DetailSubjectName");
                    //名稱全名
                    l_oTmpItem.m_FullName = l_oTmpItem.m_ComAccName + " " + l_oTmpItem.m_ComDetailAccName;

                    //有在AccInfo Map內 存入 輸出打包物件
                    l_oRtnExporeExcel.m_oMapedAccInfos.Add(l_oTmpAccInfoItem);


                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpAmount = row.Field<string>("DebitAmount");
                        l_oTmpItem.m_ComDAmount = tmpAmount;

                    }
                    else
                    {
                        //貸方金額
                        tmpAmount = row.Field<string>("CreditAmount");
                        l_oTmpItem.m_ComCAmount = tmpAmount;
                    }

                    //儲存 -> Expore Object
                    l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                }

            }
            */

            
            //Compare 把同FullNo 不同天Day的DAmount值,CAmount值相加或新增Item到List
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                string ComAccountName = "";

                //TmpExcelItem l_oProcessedItem = null;
                l_oTmpItem = null;
                //科目代號
                string l_sAccountNo = row.Field<string>("AccountNo");
                //明細科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                //科目名稱
                string l_sAccountName = row.Field<string>("SubjectName");
                //明細科目名稱
                string l_sDetailAccountName = row.Field<string>("DetailSubjectName");
                //傳票號碼 利用傳票號碼來區分 原料期初,原料期末(都叫原料存貨)
                string l_sSubpNo = row.Field<string>("SubpNo");

                //每次找出AccNo,DtlAccNo 就抓出該Map的Info


                //把AccNo-DtlAccNo相同的(不同天日期)相加
                ///if (!string.IsNullOrEmpty(l_oProcessedItem.m_ComDetailAccNo))
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {
                    //string ComAccountNo = row.Field<string>("AccountNo");

                    /// ComAccountNo = l_oProcessedItem.m_ComAccountNo + "-" +
                    ///     l_oProcessedItem.m_ComDetailAccNo;
                    ComAccountNo = l_sAccountNo + "-" + l_sDetailAccountNo;
                    ComAccountName = l_sAccountName + " " + l_sDetailAccountName;
                }
                else
                {
                    ///ComAccountNo = l_oProcessedItem.m_ComAccountNo;
                    ComAccountNo = l_sAccountNo;
                    ComAccountName = l_sAccountName;
                }

                if (l_iRowIndex == 1)
                {

                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }

                //如果有此[科目代碼+明細科目代碼]
                //if ((l_oComDataTable.Rows.Count == 0) || (l_oComDataTable.Rows.Find(l_iRowIndex-1) == null))
                //if ( (l_oComAccountNo.Count() == 0) || (l_oComAccountNo.Contains(ComAccountNo) == false) )
                ///if ((l_oComAccNoAmount.Count() == 0) || (l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo) == null))
                if ((l_oRtnExporeExcel.m_oBaseAttrItems.Count() == 0) || (l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo) == null))
                {
                    //DataRow tmprow = new DataRow();                       
                    //

                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.NewRow();
                    //tmprow["AccountNo"] = ComAccountNo;
                    //建立新的Item
                    l_oTmpItem = new TmpExcelItem();
                    ///l_oTmpItem.m_ComAccountNo = ComAccountNo;
                    ///l_oComAccNoAmount.Add(l_oTmpItem);
                    l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                    l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                    l_oTmpItem.m_FullNo = ComAccountNo;

                    l_oTmpItem.m_ComAccName = l_sAccountName;
                    l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                    l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                    l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpAmount = row.Field<string>("DebitAmount");
                        l_oTmpItem.m_ComDAmount = tmpAmount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpAmount = row.Field<string>("CreditAmount");
                        l_oTmpItem.m_ComCAmount = tmpAmount;
                    }

                    //tmprow["Amount"] = tmpAmount;
                    //l_oComDataTable.Rows.Add(tmprow);
                    //l_oComAmount.Add(tmpAmount);
                    ///l_oTmpItem.m_ComAmount = tmpAmount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    ///l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                    l_oTmpItem.m_FullName = ComAccountName;

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";
                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.Rows.Find(ComAccountNo);
                    //tmpamount = tmprow[1].ToString();
                    //tmpamount = l_oComDataTable.Rows.Find(ComAccountNo).Field<string>(1);
                    ///l_oTmpItem = l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo);
                    ///tmpamount = l_oTmpItem.m_ComAmount;
                    l_oTmpItem = l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo );

                    //20201226 CCL+ 另外特別處理AccountNo == 1192 /////////////////////////
                    if((row.Field<string>("AccountNo") == "1192") && 
                        (row.Field<string>("SubjectName") == "原料存貨"))
                    {
                        //改成不累加金額,而是新增一個新項目,之後用SubpNo傳票編號來判斷哪個是期初;哪個是期末
                        //建立新的Item
                        l_oTmpItem = new TmpExcelItem();                       
                        l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                        l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                        l_oTmpItem.m_FullNo = ComAccountNo;

                        l_oTmpItem.m_ComAccName = l_sAccountName;
                        l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                        l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                        l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                        string tmpAmount = "";
                        //借方金額
                        if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                        {
                            //借方金額 D
                            tmpAmount = row.Field<string>("DebitAmount");
                            l_oTmpItem.m_ComDAmount = tmpAmount;
                        }
                        else
                        {
                            //貸方金額 C
                            tmpAmount = row.Field<string>("CreditAmount");
                            l_oTmpItem.m_ComCAmount = tmpAmount;
                        }
                       
                        l_oTmpItem.m_FullName = ComAccountName;
                        continue;
                    }
                    ///////////////////////////////////////////////////////////////////////////

                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpnextamount = row.Field<string>("DebitAmount");
                        if(l_oTmpItem.m_ComDAmount == null)
                        { l_oTmpItem.m_ComDAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComDAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComDAmount = tmpnextamount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpnextamount = row.Field<string>("CreditAmount");
                        if (l_oTmpItem.m_ComCAmount == null)
                        { l_oTmpItem.m_ComCAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComCAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComCAmount = tmpnextamount;
                    }

                    ///tmpamount = Convert.ToString(Int32.Parse(tmpamount) + Int32.Parse(tmpnextamount));
                    ///l_oTmpItem.m_ComAmount = tmpamount; //累加值
                                                        //l_oComDataTable.Rows.Find(ComAccountNo)[1] = tmpamount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    ///l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                    l_oTmpItem.m_FullName = ComAccountName;
                    

                }




            }

            //Grouping 分群
            l_oRtnExporeExcel.GroupingBaseItems(m_AccInfoDBService);
            l_oRtnExporeExcel.Calc_GID4_OperaIncome();
            l_oRtnExporeExcel.Calc_GID5_TolCostOfCashExpend();
            l_oRtnExporeExcel.Calc_GID5_OperaCosts();
            l_oRtnExporeExcel.Calc_GID6_OperaExpense();
            l_oRtnExporeExcel.Calc_GID7_NonOperaIncome();
            l_oRtnExporeExcel.Calc_GID8_NonOperaExpense();
            l_oRtnExporeExcel.Calc_RestOthersVal();


            // 20201224 CCL 列印Excel
            l_iRowIndex = 0;


            /*
            // 20201224 CCL test
            l_iRowIndex = 0;
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oBaseAttrItems)
            {
                ++l_iRowIndex;
                //科目代號
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                //20201226 CCL+ 看是DAmount有值,還是CAmount有值 就挑哪一個
                
                if((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                } else if((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }
                
            }
            */



            /*
            l_iRowIndex = 0;
            foreach (TmpExcelItem row in l_oComAccNoAmount)
            {
                ++l_iRowIndex;
                //科目代號
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_ComAccountNo;
                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_ComAccName;
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComAmount;
            }
            */

            //Main
            /************************************************
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;
                //IXLCells cells = l_oWooksheet.Rows(l_iRowIndex.ToString()).Cells();
                //foreach(IXLCell cell in cells)
                //{
                //cell.Value = 
                //}


                //科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {


                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo") + "-" +
                        l_sDetailAccountNo;
                    string ComAccountNo = l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value.ToString();


                    l_oComDataTable.Rows.Add(new DataColumn(ComAccountNo));
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo");
                }

                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.Field<string>("SubjectName") + " " +
                    row.Field<string>("DetailSubjectName");


                if (l_iRowIndex == 1)
                {

                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }


                //借方金額
                if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("DebitAmount");
                }
                else
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("CreditAmount");
                }




            }
            ***********************************************/


            // 20201224 CCL test
            //Styling
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
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
            


            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }


        /////////////////////////////////////////////////////////////////////////////////////////////

        //20201227 CCL+ ////////////////////////////////////////////////////////////////////////////
        public static MERP_AccInfoExpore GenOneShopExcelList(string p_sShopNo, 
                                        DataSet p_oAllDataSet, IXLWorksheet p_oWooksheet)
        {
            //最終Expore要輸出的打包物件
            MERP_AccInfoExpore l_oRtnExporeExcel = new MERP_AccInfoExpore();
            l_oRtnExporeExcel.m_sShopId = p_sShopNo; //Shop ID

            TmpExcelItem l_oTmpItem = null;

            int l_iRowIndex = 0;

            //Compare 把同FullNo 不同天Day的DAmount值,CAmount值相加或新增Item到List
            foreach (DataRow row in p_oAllDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                string ComAccountName = "";

                //TmpExcelItem l_oProcessedItem = null;
                l_oTmpItem = null;
                //科目代號
                string l_sAccountNo = row.Field<string>("AccountNo");
                //明細科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                //科目名稱
                string l_sAccountName = row.Field<string>("SubjectName");
                //明細科目名稱
                string l_sDetailAccountName = row.Field<string>("DetailSubjectName");
                //傳票號碼 利用傳票號碼來區分 原料期初,原料期末(都叫原料存貨)
                string l_sSubpNo = row.Field<string>("SubpNo");

               
                //把AccNo-DtlAccNo相同的(不同天日期)相加               
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {
                   
                    ComAccountNo = l_sAccountNo + "-" + l_sDetailAccountNo;
                    ComAccountName = l_sAccountName + " " + l_sDetailAccountName;
                }
                else
                {
                   
                    ComAccountNo = l_sAccountNo;
                    ComAccountName = l_sAccountName;
                }

                if (l_iRowIndex == 1)
                {

                    //部門
                    l_oRtnExporeExcel.m_sShopName = row.Field<string>("DepartName");
                   
                }

                //如果有此[科目代碼+明細科目代碼]               
                if ((l_oRtnExporeExcel.m_oBaseAttrItems.Count() == 0) || (l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo) == null))
                {
                    
                    //建立新的Item
                    l_oTmpItem = new TmpExcelItem();
                   
                    l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                    l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                    l_oTmpItem.m_FullNo = ComAccountNo;

                    l_oTmpItem.m_ComAccName = l_sAccountName;
                    l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                    l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                    l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpAmount = row.Field<string>("DebitAmount");
                        l_oTmpItem.m_ComDAmount = tmpAmount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpAmount = row.Field<string>("CreditAmount");
                        l_oTmpItem.m_ComCAmount = tmpAmount;
                    }

                   
                    //科目名稱 - 明細科目名稱                   
                    l_oTmpItem.m_FullName = ComAccountName;

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";                   
                    l_oTmpItem = l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo);

                    //20201226 CCL+ 另外特別處理AccountNo == 1192 /////////////////////////
                    if ((row.Field<string>("AccountNo") == "1192") &&
                        (row.Field<string>("SubjectName") == "原料存貨"))
                    {
                        //改成不累加金額,而是新增一個新項目,之後用SubpNo傳票編號來判斷哪個是期初;哪個是期末
                        //建立新的Item
                        l_oTmpItem = new TmpExcelItem();
                        l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                        l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                        l_oTmpItem.m_FullNo = ComAccountNo;

                        l_oTmpItem.m_ComAccName = l_sAccountName;
                        l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                        l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                        l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                        string tmpAmount = "";
                        //借方金額
                        if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                        {
                            //借方金額 D
                            tmpAmount = row.Field<string>("DebitAmount");
                            l_oTmpItem.m_ComDAmount = tmpAmount;
                        }
                        else
                        {
                            //貸方金額 C
                            tmpAmount = row.Field<string>("CreditAmount");
                            l_oTmpItem.m_ComCAmount = tmpAmount;
                        }

                        l_oTmpItem.m_FullName = ComAccountName;
                        continue;
                    }
                    ///////////////////////////////////////////////////////////////////////////

                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpnextamount = row.Field<string>("DebitAmount");
                        if (l_oTmpItem.m_ComDAmount == null)
                        { l_oTmpItem.m_ComDAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComDAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComDAmount = tmpnextamount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpnextamount = row.Field<string>("CreditAmount");
                        if (l_oTmpItem.m_ComCAmount == null)
                        { l_oTmpItem.m_ComCAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComCAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComCAmount = tmpnextamount;
                    }

                    
                    //科目名稱 - 明細科目名稱                   
                    l_oTmpItem.m_FullName = ComAccountName;


                }


            }

            //Grouping 分群
            l_oRtnExporeExcel.GroupingBaseItems(m_AccInfoDBService);            
            l_oRtnExporeExcel.ReOrderByPrintOrder(); //20201229 CCL+ ReOrderBy PrintOrder
            l_oRtnExporeExcel.Calc_GID4_OperaIncome();
            l_oRtnExporeExcel.Calc_GID5_TolCostOfCashExpend();
            l_oRtnExporeExcel.Calc_GID5_OperaCosts();
            l_oRtnExporeExcel.Calc_GID6_OperaExpense();
            l_oRtnExporeExcel.Calc_GID7_NonOperaIncome();
            l_oRtnExporeExcel.Calc_GID8_NonOperaExpense();
            l_oRtnExporeExcel.Calc_RestOthersVal();

            return l_oRtnExporeExcel; //一家店Data
        }
        /////////////////////////////////////////////////////////////////////////////////////////////

        /* 20201229 CCL-
        //20201227 CCL+
        public static bool ExporeOneShopExcelList(int p_iShopIndex,
                                       MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                       String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 2; //科目NO,Name 兩行
            const int PRINT_COLS = 4; //科目代碼,科目名稱,實績,比率
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            if(p_iShopIndex == 0)
            {
                p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            } else
            {
                p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            }
            
            //部門名稱
            p_oWooksheet.Cell(7, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4+ l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolOperaCosts;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "原料期初";
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));

            }
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                else if (((row.m_ComDAmount != null) && (row.m_ComDAmount != "0")) &&
                        ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")))
                {
                    //原料進料 總部 DAmount和CAmount都有值要相減
                    double tmpVal = Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = tmpVal;

                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "減 原料期末";
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oEndRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oEndRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaMargin;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTopOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBtmOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //顯示 營業利益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBussInterest;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //顯示 非營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 非營業支出
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dConsuCurrentProfitLoss;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //顯示 空白
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolCostOfCashExpend;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dCashExpendForCurrPeriod;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }

        
        //20201227 CCL+ 不顯示科目代碼版本 //////////////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV1(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            if (p_iShopIndex == 0)
            {
                p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            }
            else
            {
                p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            }

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;  
            p_oWooksheet.Cell(7, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;            
            //顯示 營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolOperaCosts;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "原料期初";
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));

            }
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;

                if (((row.m_ComDAmount != null) && (row.m_ComDAmount != "0")) &&
                        ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")))
                {
                    //原料進料 總部 DAmount和CAmount都有值要相減
                    double tmpVal = Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = tmpVal;

                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                } else if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "減 原料期末";
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oEndRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oEndRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaMargin;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTopOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBtmOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBussInterest;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dConsuCurrentProfitLoss;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolCostOfCashExpend;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dCashExpendForCurrPeriod;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////
        */

        //20201228 CCL+ 不顯示科目代碼版本 //////////////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV2(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            if (p_iShopIndex == 0)
            {
                p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            }
            else
            {
                p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            }

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;  
            p_oWooksheet.Cell(7, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                /*
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                */

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 營業總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolOperaCosts;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "原料期初";
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = tmpVal;

                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
            } else 
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));

            }
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;

                /*
                if (((row.m_ComDAmount != null) && (row.m_ComDAmount != "0")) &&
                        ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")))
                {
                    //原料進料 總部 DAmount和CAmount都有值要相減
                    double tmpVal = Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = tmpVal;

                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                }
                else if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                */

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "減 原料期末";
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = tmpVal;

                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oEndRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oEndRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
            }
            //顯示 實際用量總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "實際用量總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dActualTolCosts;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dActualTolCosts);
            //顯示 營業毛利
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaMargin;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTopOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                /*
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                */

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBtmOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBussInterest;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                /*
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                */

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                /*
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                */

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                   
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dConsuCurrentProfitLoss;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolCostOfCashExpend;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dCashExpendForCurrPeriod;
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////

        
        //20201228 CCL+ 不顯示科目代碼版本 //////////////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV3(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            if (p_iShopIndex == 0)
            {
                p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            }
            else
            {
                p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            }

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;  
            p_oWooksheet.Cell(7, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                
                //if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                //{
                    //貸方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                //}
                //else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                //{
                    //借方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                //}
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 營業總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolOperaCosts);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "原料期初";
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
            }
            else
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));

            }
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;

                
                //if (((row.m_ComDAmount != null) && (row.m_ComDAmount != "0")) &&
                //        ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")))
                //{
                    //原料進料 總部 DAmount和CAmount都有值要相減
                //    double tmpVal = Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = tmpVal;

                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //}
                //else if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                //{
                    //貸方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                //}
                //else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                //{
                    //借方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                //}
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "減 原料期末";
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
            }
            //顯示 實際用量總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "實際用量總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dActualTolCosts);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dActualTolCosts);
            //顯示 營業毛利
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaMargin);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTopOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                
                //if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                //{
                    //貸方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                //}
                //else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                //{
                    //借方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                //}
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBtmOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBussInterest);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                
                //if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                //{
                    //貸方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                //}
                //else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                //{
                    //借方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                //}
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullName;
                
                //if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                //{
                    //貸方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComDAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                //}
                //else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                //{
                    //借方金額
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_ComCAmount;
                //    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                //}
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                }

            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////

        
        //20201228 CCL+ 不顯示科目代碼版本 //////////////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV4(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int HEAD_ROWS = 4; //新版表頭數目
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //if (p_iShopIndex == 0)
            //{
            //    p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //}
            //else
            //{
            //    p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            //}

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;              
            p_oWooksheet.Cell(3, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //營業收入 Item欄位填滿底色
            IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字
            if(AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF1 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF1.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group4 Item欄位名稱縮進
                IXLRange l_oRangeGID4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID4.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font                       
                        IXLRange l_oRangeRedF2 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF2.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
               

            }
                

            //顯示 營業總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolOperaCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolOperaCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF3 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF3.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "原料期初";
            //原料期初 Item欄位名稱縮進
            IXLRange l_oRangeBeginRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
            l_oRangeBeginRawM.Style.Alignment.SetIndent(2);
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group5 Item欄位名稱縮進
                IXLRange l_oRangeGID5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID5.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF5.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "減 原料期末";
            //減 原料期末 Item欄位名稱縮進
            IXLRange l_oRangeEndRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
            l_oRangeEndRawM.Style.Alignment.SetIndent(2);
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            //顯示 實際用量總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "實際用量總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dActualTolCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dActualTolCosts);
            //實際用量總成本 Item欄位填滿底色
            IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dActualTolCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF7.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaMargin);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaMargin))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF8.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTopOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTopOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF9 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF9.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group6 Item欄位名稱縮進
                IXLRange l_oRangeGID6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID6.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF10 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF10.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBtmOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //Bottom營業費用 Item欄位填滿底色
            IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBtmOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF11 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF11.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBussInterest);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBussInterest))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF12 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF12.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //非營業收入 Item欄位填滿底色
            IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF13 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF13.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group7 Item欄位名稱縮進
                IXLRange l_oRangeGID7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID7.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF14 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF14.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //非營業支出 Item欄位填滿底色
            IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF15 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF15.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group8 Item欄位名稱縮進
                IXLRange l_oRangeGID8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID8.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF16 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF16.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF17 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF17.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //現金支出總成本 Item欄位填滿底色
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolCostOfCashExpend))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF18 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF18.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF19 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF19.Style.Font.SetFontColor(XLColor.Red);
            }

            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS+1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(3, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////


        //20201228 CCL+ 不顯示科目代碼版本 //////////////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV5(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int HEAD_ROWS = 4; //新版表頭數目
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //if (p_iShopIndex == 0)
            //{
            //    p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //}
            //else
            //{
            //    p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            //}

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;              
            p_oWooksheet.Cell(3, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //營業收入 Item欄位填滿底色
            IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF1 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF1.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group4 Item欄位名稱縮進
                IXLRange l_oRangeGID4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID4.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font                       
                        IXLRange l_oRangeRedF2 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF2.Style.Font.SetFontColor(XLColor.Red);
                    }
                }


            }


            //顯示 營業總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolOperaCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolOperaCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF3 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF3.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "原料期初";
            //原料期初 Item欄位名稱縮進
            IXLRange l_oRangeBeginRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
            l_oRangeBeginRawM.Style.Alignment.SetIndent(2);
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }

            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group5 Item欄位名稱縮進
                IXLRange l_oRangeGID5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID5.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF5.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "減 原料期末";
            //減 原料期末 Item欄位名稱縮進
            IXLRange l_oRangeEndRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
            l_oRangeEndRawM.Style.Alignment.SetIndent(2);
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            //顯示 實際用量總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "實際用量總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dActualTolCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dActualTolCosts);
            //實際用量總成本 Item欄位填滿底色
            IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dActualTolCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF7.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaMargin);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaMargin))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF8.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTopOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTopOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF9 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF9.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group6 Item欄位名稱縮進
                IXLRange l_oRangeGID6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID6.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF10 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF10.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBtmOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //Bottom營業費用 Item欄位填滿底色
            IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBtmOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF11 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF11.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBussInterest);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBussInterest))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF12 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF12.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //非營業收入 Item欄位填滿底色
            IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF13 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF13.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group7 Item欄位名稱縮進
                IXLRange l_oRangeGID7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID7.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF14 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF14.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //非營業支出 Item欄位填滿底色
            IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF15 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF15.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = row.m_FullName;
                //Group8 Item欄位名稱縮進
                IXLRange l_oRangeGID8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address);
                l_oRangeGID8.Style.Alignment.SetIndent(2);


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF16 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF16.Style.Font.SetFontColor(XLColor.Red);
                    }
                }

            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF17 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF17.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //現金支出總成本 Item欄位填滿底色
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolCostOfCashExpend))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF18 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF18.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF19 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF19.Style.Font.SetFontColor(XLColor.Red);
            }

            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(3, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////

        //20201231 CCL+ 不顯示科目代碼版本 科目名稱唯一[頭部]////////////////////////////////////////////////////
        //顯示科目名稱 一行
        public static bool ExporeAccountNameInfoCol(
                                IXLWorksheet p_oWooksheet
                               , MERP_AccInfoTolExcel p_oAllShopNOData )
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int HEAD_ROWS = 4; //新版表頭數目
            int l_iRowIndex = 0;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            //MERP_AccInfoExpore l_oRtnExporeExcel = p_oOneShopNOData;

            
            //1.顯示 營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業收入";
            //營業收入 Item欄位填滿底色
            //IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
            //l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //2.顯示 Group4 Item欄位名稱         
            foreach (AccountInfo Info in p_oAllShopNOData.m_oCombineAccInfosGID4)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Info.AccountName + " " + Info.DetailAccName;
                //Group4 Item欄位名稱縮進
                IXLRange l_oRangeGID4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
                l_oRangeGID4.Style.Alignment.SetIndent(2);
            }

            //3.顯示 營業總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業總成本";

            //4.顯示 原料期初
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "原料期初";
            //原料期初 Item欄位名稱縮進
            IXLRange l_oRangeBeginRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
            l_oRangeBeginRawM.Style.Alignment.SetIndent(2);

            //5.顯示 GroupID5
            foreach (AccountInfo Info in p_oAllShopNOData.m_oCombineAccInfosGID5)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Info.AccountName + " " + Info.DetailAccName;                
                //Group5 Item欄位名稱縮進
                IXLRange l_oRangeGID5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
                l_oRangeGID5.Style.Alignment.SetIndent(2);
            }

            //6.顯示 減 原料期末
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "減 原料期末";
            //減 原料期末 Item欄位名稱縮進
            IXLRange l_oRangeEndRawM = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
            l_oRangeEndRawM.Style.Alignment.SetIndent(2);

            //7.顯示 實際用量總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "實際用量總成本";
            //實際用量總成本 Item欄位填滿底色
            //IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
            //l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //8.顯示 營業毛利
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業毛利";

            //9.顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業費用";

            //10.顯示 GroupID6
            foreach (AccountInfo Info in p_oAllShopNOData.m_oCombineAccInfosGID6)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Info.AccountName + " " + Info.DetailAccName;
                //Group6 Item欄位名稱縮進
                IXLRange l_oRangeGID6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
                l_oRangeGID6.Style.Alignment.SetIndent(2);
            }

            //11.顯示 營業費用
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業費用";
            //Bottom營業費用 Item欄位填滿底色
            //IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //12.顯示 營業利益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "營業利益";

            //13.顯示 非營業收入
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "非營業收入";
            //非營業收入 Item欄位填滿底色
            //IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //14.顯示 GroupID7
            foreach (AccountInfo Info in p_oAllShopNOData.m_oCombineAccInfosGID7)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 ).Value = Info.AccountName + " " + Info.DetailAccName;
                //Group7 Item欄位名稱縮進
                IXLRange l_oRangeGID7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 ).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
                l_oRangeGID7.Style.Alignment.SetIndent(2);
            }

            //15.顯示 非營業支出
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "非營業支出";
            //非營業支出 Item欄位填滿底色
            //IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //16.顯示 GroupID8
            foreach (AccountInfo Info in p_oAllShopNOData.m_oCombineAccInfosGID8)
            {
                ++l_iRowIndex;
                //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = Info.AccountName + " " + Info.DetailAccName;
                //Group8 Item欄位名稱縮進
                IXLRange l_oRangeGID8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
                l_oRangeGID8.Style.Alignment.SetIndent(2);

            }

            //17.顯示 實際用量本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "實際用量本期損益";
            //實際用量本期損益 Item欄位填滿底色
            //IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);

            //18.顯示 空白
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "";

            //19.顯示 現金支出總成本
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "現金支出總成本";
            //現金支出總成本 Item欄位填滿底色
            //IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);

            //20.顯示 現金支出本期損益
            ++l_iRowIndex;
            //l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Value = "現金支出本期損益";
            //現金支出本期損益 Item欄位填滿底色
            //IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            //l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);


            //20210104 CCL+ Cell符合文字大小
            //IXLRange l_oRangeAllCell = l_oWooksheet.Range(l_oWooksheet.Cell(1 + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address);
            //IXLColumn l_oCol = l_oWooksheet.Column(1);
            //l_oCol.Width = 20;


            return false;
        }
        // /////////////////////////////////////////////////////////////////////////////////////////////////////

        public static bool ExporeOneShopExcelListV6(int p_iShopIndex,
                               MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                               String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int HEAD_ROWS = 4; //新版表頭數目
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            //20201227 CCL- p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //if (p_iShopIndex == 0)
            //{
            //    p_oWooksheet.Cell(5, 1).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //}
            //else
            //{
            //    p_oWooksheet.Cell(5, 1).Value += ", " + l_oRtnExporeExcel.m_sShopName;
            //}

            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;              
            p_oWooksheet.Cell(3, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //營業收入 Item欄位填滿底色
            IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF1 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF1.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
               

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font                       
                        IXLRange l_oRangeRedF2 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF2.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }


            }


            //顯示 營業總成本
            ++l_iRowIndex;
            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolOperaCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolOperaCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF3 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF3.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();            
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }

            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF5.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            //顯示 實際用量總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dActualTolCosts);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dActualTolCosts);
            //實際用量總成本 Item欄位填滿底色
            IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dActualTolCosts))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF7.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業毛利
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dOperaMargin);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dOperaMargin))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF8.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業費用
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTopOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTopOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF9 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF9.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF10 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF10.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }
            //顯示 營業費用
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBtmOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //Bottom營業費用 Item欄位填滿底色
            IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBtmOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF11 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF11.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 營業利益
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dBussInterest);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dBussInterest))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF12 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF12.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 非營業收入
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaIncome);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //非營業收入 Item欄位填滿底色
            IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaIncome))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF13 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF13.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF14 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF14.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }
            //顯示 非營業支出
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dNonOperaExpense);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //非營業支出 Item欄位填滿底色
            IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dNonOperaExpense))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF15 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF15.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;               

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF16 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                        l_oRangeRedF16.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF17 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF17.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 空白
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //現金支出總成本 Item欄位填滿底色
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dTolCostOfCashExpend))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF18 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF18.Style.Font.SetFontColor(XLColor.Red);
            }
            //顯示 現金支出本期損益
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);
            //現金支出本期損益 Item欄位填滿底色
            IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF19 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                l_oRangeRedF19.Style.Font.SetFontColor(XLColor.Red);
            }

            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(3, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }

        /// ////////////////////////////////////////////////////
        public static bool ExporeOneShopExcelListV7(int p_iShopIndex,
                           MERP_NewOrderPtrTB p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            //20201231 CCL-改成兩行 const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int PRINT_COLS = 2; //金額,比率
            const int HEAD_ROWS = 4; //新版表頭數目
            const int TOLCOUNT_COL = 1; //合計 行//20210103 CCL+
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_NewOrderPtrTB l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

           
            //部門名稱
            //p_oWooksheet.Cell(7, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;              
            p_oWooksheet.Cell(3, 2 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //20201231 CCL+ 改成金額
            //l_oWooksheet.Cell(4, 2 + l_iPadColIndex).Value = "金額";
            //比率%
            //l_oWooksheet.Cell(4, 3 + l_iPadColIndex).Value = "比率%";

            //1.顯示 營業收入
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[0]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[0]));
            //營業收入 Item欄位填滿底色
            IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[0]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF1 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF1.Style.Font.SetFontColor(XLColor.Red);
            }


            //2.顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oNOAttrItemsGID4)
            {
                ++l_iRowIndex;
               

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount)));
                    //負數(XXX) 顯示紅色字
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                    //為負數設定Style Red Font                       
                        IXLRange l_oRangeRedF2 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                        l_oRangeRedF2.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }


            }


            //3.顯示 營業總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[1]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[1]));
            //負數(XXX) 顯示紅色字
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[1]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF3 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF3.Style.Font.SetFontColor(XLColor.Red);
            }

            //4.顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oNOAttrBeginRawMaterial.First();            
            //20201229 CCL+
            if (((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0")) &&
                       ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oBeginRawMat.m_ComDAmount) - Convert.ToInt32(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(tmpVal));
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount)));
                //負數(XXX) 顯示紅色字
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComDAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oBeginRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount)));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oBeginRawMat.m_ComCAmount))
                {
                    //為負數設定Style Red Font
                    IXLRange l_oRangeRedF4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF4.Style.Font.SetFontColor(XLColor.Red);
                }
            }

            //5.顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oNOAttrItemsGID5)
            {
                ++l_iRowIndex;
                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {
                    //貸方金額                    
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount)));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                        l_oRangeRedF5.Style.Font.SetFontColor(XLColor.Red);
                    }
                } else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }

            //6.顯示 減 原料期末
            ++l_iRowIndex;

            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oNOAttrEndRawMaterial.Last();
            //20201229 CCL+
            if (((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0")) &&
                       ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0")))
            {
                //原料進料 總部 DAmount和CAmount都有值要相減
                double tmpVal = Convert.ToInt32(l_oEndRawMat.m_ComDAmount) - Convert.ToInt32(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(tmpVal);

                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(tmpVal));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(tmpVal))
                {
                //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComDAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount)));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComDAmount))
                {
                //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oEndRawMat.m_ComCAmount);
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount)));
                //負數(XXX) 顯示紅色字               
                if (AmountNumProcess.ChkMinusNumRedFont(l_oEndRawMat.m_ComCAmount))
                {
                //為負數設定Style Red Font
                    IXLRange l_oRangeRedF6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                    l_oRangeRedF6.Style.Font.SetFontColor(XLColor.Red);
                }
            }



            //7.顯示 實際用量總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[2]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[2]));
            //實際用量總成本 Item欄位填滿底色
            IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[2]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF7.Style.Font.SetFontColor(XLColor.Red);
            }


            //8.顯示 營業毛利
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[3]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[3]));
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[3]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF8.Style.Font.SetFontColor(XLColor.Red);
            }

            //9.顯示 營業費用
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[4]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[4]));
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[4]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF9 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF9.Style.Font.SetFontColor(XLColor.Red);
            }


            //10.顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oNOAttrItemsGID6)
            {
                ++l_iRowIndex;                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount)));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF10 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                        l_oRangeRedF10.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }


            //11.顯示 營業費用
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[5]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[5]));
            //Bottom營業費用 Item欄位填滿底色
            IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[5]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF11 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF11.Style.Font.SetFontColor(XLColor.Red);
            }

            //12.顯示 營業利益
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[6]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[6]));
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[6]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF12 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF12.Style.Font.SetFontColor(XLColor.Red);
                //l_oRangeRedF12.Style.Font.FontColor = XLColor.Red;
            }

            //13.顯示 非營業收入
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[7]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[7]));
            //非營業收入 Item欄位填滿底色
            IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[7]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF13 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF13.Style.Font.SetFontColor(XLColor.Red);
            }

            //14.顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oNOAttrItemsGID7)
            {
                ++l_iRowIndex;                

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount)));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF14 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                        l_oRangeRedF14.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }

            //15.顯示 非營業支出
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[8]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[8]));
            //非營業支出 Item欄位填滿底色
            IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[8]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF15 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF15.Style.Font.SetFontColor(XLColor.Red);
            }

            //16.顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oNOAttrItemsGID8)
            {
                ++l_iRowIndex;               

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                {

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(row.m_ComAmount);
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComAmount)));
                    //負數(XXX) 顯示紅色字               
                    if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                    {
                        //為負數設定Style Red Font
                        IXLRange l_oRangeRedF16 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                        l_oRangeRedF16.Style.Font.SetFontColor(XLColor.Red);
                    }
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = 0;
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = 0;
                }

            }

            //17.顯示 實際用量本期損益
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[9]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[9]));
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[9]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF17 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF17.Style.Font.SetFontColor(XLColor.Red);
            }

            //18.顯示 空白
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = "";
            //20210204 CCL 修正空白行合併
            IXLRange l_oRangeSpace = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, TOPCOLS + 2 + l_iPadColIndex + TOLCOUNT_COL).Address);
            l_oRangeSpace.Merge();


            //19.顯示 現金支出總成本
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[10]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[10]));
            //現金支出總成本 Item欄位填滿底色
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[10]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF18 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF18.Style.Font.SetFontColor(XLColor.Red);
            }


            //20.顯示 現金支出本期損益
            ++l_iRowIndex;           
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_NewOrderItemPtrs[11]);
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Value = AmountNumProcess.ShowMinusPercent(l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_NewOrderItemPtrs[11]));
            //現金支出本期損益 Item欄位填滿底色
            IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
            l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_NewOrderItemPtrs[11]))
            {
                //為負數設定Style Red Font
                IXLRange l_oRangeRedF19 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 3 + l_iPadColIndex).Address);
                l_oRangeRedF19.Style.Font.SetFontColor(XLColor.Red);
            }


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //20210105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體";
            l_oRange4.Style.Font.FontName = "標楷體";
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(3, 3 + l_iPadColIndex).Address);
            l_oRange7.Merge();

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
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopNo);

            return false;
        }

        // //////////////////////////////////////////////////////////////////////////////////////////

        //20210103 CCL+ 顯示合計 行 [尾部]////////////////////////////////////////////////////////////
        public static bool ExporeAllShopColsAmount(int p_iShopIndex,
                          MERP_AccInfoTolExcel p_oAllShopNOData, IXLWorksheet p_oWooksheet,
                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 1; //科目Name 1行
            //20201231 CCL-改成兩行 const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int PRINT_COLS = 2; //金額,比率
            const int HEAD_ROWS = 4; //新版表頭數目
            const int TOLAMOUNT_COLS = 1; //20210103 CCL+ 合計
            int l_iShopValCount = 2; //金額,比率 一行


            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoTolExcel l_oRtnExporeExcel = p_oAllShopNOData;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;
            int l_iValIndex = 0;
            //
            IXLRange l_oRangeTolAmount = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(4, 2 + l_iPadColIndex).Address);
            l_oRangeTolAmount.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRangeTolAmount.Style.Font.FontColor = XLColor.White;
            //l_oRangeTolAmount.Style.Font.FontName = "微軟正黑體";
            l_oRangeTolAmount.Style.Font.FontName = "標楷體";
            l_oRangeTolAmount.Style.Font.FontSize = 11;
            l_oRangeTolAmount.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //合計                       
            p_oWooksheet.Cell(3, 2 + l_iPadColIndex).Value = "合計"; //第一行是科目名稱
            IXLRange l_oRangeTolName = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(4, 2 + l_iPadColIndex).Address);
            l_oRangeTolName.Merge();
            l_oRangeTolName.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            l_oRangeTolName.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //1.顯示 營業收入
            ++l_iRowIndex;            
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //營業收入 Item欄位填滿底色
            IXLRange l_oRangeOPIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeOPIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT1 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT1.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //2.顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_ShopNWOPtrTBs[0].m_oNOAttrItemsGID4)
            {
                ++l_iRowIndex;


                //20201228 CCL Modify 直接顯示計算好的金額Amount
                //if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                //{
                l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);

                //負數(XXX) 顯示紅色字
                //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
                //{
                    //為負數設定Style Red Font                       
                //    IXLRange l_oRangeRedT2 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                //    l_oRangeRedT2.Style.Font.SetFontColor(XLColor.Red);
                //}
                //}
                l_iValIndex++;
            }


            //3.顯示 營業總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT3 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT3.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //4.顯示 原料期初
            ++l_iRowIndex;                        
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT4 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT4.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //5.顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_ShopNWOPtrTBs[0].m_oNOAttrItemsGID5)
            {
                ++l_iRowIndex;

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                //if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                //{
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);

                //負數(XXX) 顯示紅色字               
                //if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                //{
                //為負數設定Style Red Font
                //    IXLRange l_oRangeRedT5 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                //    l_oRangeRedT5.Style.Font.SetFontColor(XLColor.Red);
                //}
                //}
                l_iValIndex++;

            }


            //6.顯示 減 原料期末
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT6 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT6.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //7.顯示 實際用量總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //實際用量總成本 Item欄位填滿底色
            IXLRange l_oRangeActTolCost = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeActTolCost.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT7 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT7.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //8.顯示 營業毛利
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT8 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT8.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //9.顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
                //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT9 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT9.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //10.顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_ShopNWOPtrTBs[0].m_oNOAttrItemsGID6)
            {
                ++l_iRowIndex;

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                //if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                //{

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
                //負數(XXX) 顯示紅色字               
                //if (AmountNumProcess.ChkMinusNumRedFont(row.m_ComAmount))
                //{
                //為負數設定Style Red Font
                //    IXLRange l_oRangeRedT10 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                //    l_oRangeRedT10.Style.Font.SetFontColor(XLColor.Red);
                //}
                //}
                l_iValIndex++;

            }


            //11.顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //Bottom營業費用 Item欄位填滿底色
            IXLRange l_oRangeBtmOpeExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeBtmOpeExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT11 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT11.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //12.顯示 營業利益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT12 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT12.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //13.顯示 非營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //非營業收入 Item欄位填滿底色
            IXLRange l_oRangeNonOpIncome = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeNonOpIncome.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT13 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT13.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //14.顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_ShopNWOPtrTBs[0].m_oNOAttrItemsGID7)
            {
                ++l_iRowIndex;

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                //if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                //{

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);

                //負數(XXX) 顯示紅色字               
                //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
                //{
                //為負數設定Style Red Font
                //    IXLRange l_oRangeRedT14 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                //    l_oRangeRedT14.Style.Font.SetFontColor(XLColor.Red);
                //}
                //}
                l_iValIndex++;

            }


            //15.顯示 非營業支出
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //非營業支出 Item欄位填滿底色
            IXLRange l_oRangeNonOpExpe = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeNonOpExpe.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT15 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT15.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //16.顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_ShopNWOPtrTBs[0].m_oNOAttrItemsGID8)
            {
                ++l_iRowIndex;

                //20201228 CCL Modify 直接顯示計算好的金額Amount
                //if ((row.m_ComAmount != null) && (row.m_ComAmount != "0"))
                //{

                    l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);

                //負數(XXX) 顯示紅色字               
                //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
                //{
                //為負數設定Style Red Font
                //    IXLRange l_oRangeRedT16 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
                //    l_oRangeRedT16.Style.Font.SetFontColor(XLColor.Red);
                //}
                //}
                l_iValIndex++;

            }


            //17.顯示 實際用量本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //實際用量本期損益 Item欄位填滿底色
            IXLRange l_oRangeConsuCurProLoss = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeConsuCurProLoss.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iRowIndex - 1]))
            //{
                //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT17 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT17.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;



            //18.顯示 空白
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = "";
            

            //19.顯示 現金支出總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //現金支出總成本 Item欄位填滿底色
            IXLRange l_oRangeTolCostOfCashExp = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeTolCostOfCashExp.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT18 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT18.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            //20.顯示 現金支出本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Value = AmountNumProcess.ShowAmountComma(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]);            
            //現金支出本期損益 Item欄位填滿底色
            IXLRange l_oRangeCashExpCurPeriod = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            l_oRangeCashExpCurPeriod.Style.Fill.SetBackgroundColor(XLColor.Orange);
            //負數(XXX) 顯示紅色字               
            //if (AmountNumProcess.ChkMinusNumRedFont(l_oRtnExporeExcel.m_RowItemCombineAmount[l_iValIndex]))
            //{
            //為負數設定Style Red Font
            //    IXLRange l_oRangeRedT19 = l_oWooksheet.Range(l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, 2 + l_iPadColIndex).Address);
            //    l_oRangeRedT19.Style.Font.SetFontColor(XLColor.Red);
            //}
            l_iValIndex++;


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(HEAD_ROWS + 1, 1).Address, l_oWooksheet.Cell(l_iRowIndex + HEAD_ROWS, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            //2021 0105 CCL- l_oRange4.Style.Font.FontName = "微軟正黑體"; //標楷體
            l_oRange4.Style.Font.FontName = "標楷體"; //標楷體
            l_oRange4.Style.Font.FontSize = 11;

            //20201227 CCL+ 部門名稱 Merge 2 Cols                     
            //IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 3 + l_iPadColIndex).Address, l_oWooksheet.Cell(7, 4 + l_iPadColIndex).Address);
            //20210103 CCL-
            ///IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 2 + l_iPadColIndex).Address, l_oWooksheet.Cell(3, 3 + l_iPadColIndex).Address);
            ///l_oRange7.Merge();
            

            ///IXLRange l_oRangeTolVal = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1 + l_iPadColIndex).Address, l_oWooksheet.Cell(4, 2 + l_iPadColIndex).Address);
            ///l_oRangeTolVal.Merge();
            ///l_oRangeTolVal.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            ///l_oRangeTolVal.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);



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
            Trace.WriteLine(l_oRtnExporeExcel.m_ShopNWOPtrTBs.Count);

            //return false;

            return true;
        }
        // //////////////////////////////////////////////////////////////////////////////////////////

        /* 20201229 CCL-
        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201227 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 2;
            const string TABLENAME = "損益表";
            const int PRINT_COLS = 4; //科目代碼,科目名稱,實績,比率

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;
            

            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, 1).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            l_oRange6.Style.Font.FontName = "微軟正黑體";
            l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, l_iPadColIndex).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;

            //202006月帳
            l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            l_oWooksheet.Cell(2, 1).Value = "多部門損益表";


            //起迄日期：2020-07-01 ~ 2020-07-31
            l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                int iCol3Index = (PRINT_COLS * i) + 3;
                int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                l_oWooksheet.Cell(7, iCol2Index).Value = "科目名稱";
                //實績
                l_oWooksheet.Cell(8, iCol3Index).Value = "實績";
                //比率%
                l_oWooksheet.Cell(8, iCol4Index).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();
            for(int i=0; i< p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol1Index).Address, l_oWooksheet.Cell(8, iMergCol1Index).Address);
                l_oRange2.Merge();
                IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                l_oRange3.Merge();
            }

            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {
                    
                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnExpExcelShops.Add(l_oRtnExporeExcel);
                l_iShopNoIndex++;
            }
            

            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID
            //顯示輸出
            foreach(MERP_AccInfoExpore OneShopExcel in l_oRtnExpExcelShops)
            {
                //20201227 CCL*
                if(OneShopExcel.m_oBaseAttrItems.Count() > 0)
                {
                    ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                        p_sPROG_ID, p_oServer);
                    l_iRowIndex++;
                }
                
                    
            }

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
        */

        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201227 CCL+ 不顯示科目代碼版本/////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7V1(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 1;
            const string TABLENAME = "損益表";
            const int PRINT_COLS = 3; //科目名稱,實績,比率

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, 1).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            l_oRange6.Style.Font.FontName = "微軟正黑體";
            l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, l_iPadColIndex).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;

            //202006月帳
            l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            l_oWooksheet.Cell(2, 1).Value = "多部門損益表";


            //起迄日期：2020-07-01 ~ 2020-07-31
            l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                int iCol3Index = (PRINT_COLS * i) + 3;
                //int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                //l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                l_oWooksheet.Cell(7, iCol1Index).Value = "科目名稱";
                //實績
                //l_oWooksheet.Cell(8, iCol2Index).Value = "實績";
                //20201229 CCL* 改成金額
                l_oWooksheet.Cell(8, iCol2Index).Value = "金額";
                //比率%
                l_oWooksheet.Cell(8, iCol3Index).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();
            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                //int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol1Index).Address, l_oWooksheet.Cell(8, iMergCol1Index).Address);
                l_oRange2.Merge();
                //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                //l_oRange3.Merge();
            }

            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {

                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnExpExcelShops.Add(l_oRtnExporeExcel);
                l_iShopNoIndex++;
            }


            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID
            //顯示輸出
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnExpExcelShops)
            {
                //ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                //                        p_sPROG_ID, p_oServer);
                if (OneShopExcel.m_oBaseAttrItems.Count() > 0)
                {
                    //20201228 CCL- ExporeOneShopExcelListV1(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //20201229 CCL- ExporeOneShopExcelListV2(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    ExporeOneShopExcelListV3(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                    p_sPROG_ID, p_oServer); //多Camma
                    //ExporeOneShopExcelListV4(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭

                    l_iRowIndex++;
                }
                

            }

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


        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201229 CCL+ 不顯示科目代碼版本,改標頭,改置中 ///////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7V2(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 1;
            const string TABLENAME = "損益表";
            const int PRINT_COLS = 3; //科目名稱,實績,比率

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, l_iPadColIndex).Address);            
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, l_iPadColIndex).Address);
            l_oRange6.Merge();
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, l_iPadColIndex).Address);
            l_oRange7.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //202006月帳
            //l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            //l_oWooksheet.Cell(2, 1).Value = "多部門損益表";

            //督導名-區域-類別 損益表 (區域:北部,中部,南部;  類別:直營,合營,加盟)
            l_oWooksheet.Cell(1, 1).Value = "南部直營損益表";

            //起迄日期：2020-07-01 ~ 2020-07-31
            //l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            l_oWooksheet.Cell(2, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " 一 " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                int iCol3Index = (PRINT_COLS * i) + 3;
                //int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                //l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                l_oWooksheet.Cell(3, iCol1Index).Value = "科目名稱";
                //實績
                //l_oWooksheet.Cell(8, iCol2Index).Value = "實績";
                //20201229 CCL* 改成金額
                l_oWooksheet.Cell(4, iCol2Index).Value = "金額";
                //比率%
                l_oWooksheet.Cell(4, iCol3Index).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();
            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                //int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, iMergCol1Index).Address, l_oWooksheet.Cell(4, iMergCol1Index).Address);
                l_oRange2.Merge();
                //20201229 CCL+
                l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                //l_oRange3.Merge();
            }

            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {

                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnExpExcelShops.Add(l_oRtnExporeExcel);
                l_iShopNoIndex++;
            }


            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID
            //顯示輸出
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnExpExcelShops)
            {
                //ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                //                        p_sPROG_ID, p_oServer);
                if (OneShopExcel.m_oBaseAttrItems.Count() > 0)
                {
                    //20201228 CCL- ExporeOneShopExcelListV1(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //20201229 CCL- ExporeOneShopExcelListV2(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //ExporeOneShopExcelListV3(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma
                    ExporeOneShopExcelListV4(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                    p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭

                    l_iRowIndex++;
                }


            }

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
        /// ////////////////////////////////////////////////////////////////////////////////////////////////


        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201229 CCL+ 不顯示科目代碼版本,改標頭,改置中 ///////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7V3(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 1;
            const string TABLENAME = "損益表";
            const int PRINT_COLS = 3; //科目名稱,實績,比率

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, l_iPadColIndex).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, l_iPadColIndex).Address);
            l_oRange6.Merge();
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, l_iPadColIndex).Address);
            l_oRange7.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            //202006月帳
            //l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            //l_oWooksheet.Cell(2, 1).Value = "多部門損益表";

            //督導名-區域-類別 損益表 (區域:北部,中部,南部;  類別:直營,合營,加盟)
            l_oWooksheet.Cell(1, 1).Value = "南部直營損益表";

            //起迄日期：2020-07-01 ~ 2020-07-31
            //l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            l_oWooksheet.Cell(2, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " 一 " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                int iCol3Index = (PRINT_COLS * i) + 3;
                //int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                //l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                l_oWooksheet.Cell(3, iCol1Index).Value = "科目名稱";
                //實績
                //l_oWooksheet.Cell(8, iCol2Index).Value = "實績";
                //20201229 CCL* 改成金額
                l_oWooksheet.Cell(4, iCol2Index).Value = "金額";
                //比率%
                l_oWooksheet.Cell(4, iCol3Index).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();
            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                //int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, iMergCol1Index).Address, l_oWooksheet.Cell(4, iMergCol1Index).Address);
                l_oRange2.Merge();
                //20201229 CCL+
                l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                //l_oRange3.Merge();
            }

            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {

                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            //List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            MERP_AccInfoTolExcel l_oRtnTolExcelShops = new MERP_AccInfoTolExcel();
            l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init

            //取的各店Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnTolExcelShops.m_oRtnExpExcelShops.Add(l_oRtnExporeExcel);
                

                l_iShopNoIndex++;
            }

            //20201230 CCL+ 要先全部處理完
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {
                //20201230 CCL+ 處理Combine GID4,GID5,GID6,GID7,GID8
                l_oRtnTolExcelShops.CombineAllShopGIDAccInfo(OneShopExcel);

            }
            //20201230 CCL+ 重新排序
            l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();
            //20201230 CCL+
            l_oRtnTolExcelShops.CombineAllShopAccInfoTB(); //產生All科目名稱
 
            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID
            //顯示輸出
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {
                //ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                //                        p_sPROG_ID, p_oServer);

               
                if (OneShopExcel.m_oBaseAttrItems.Count() > 0)
                {
                    //20201228 CCL- ExporeOneShopExcelListV1(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //20201229 CCL- ExporeOneShopExcelListV2(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //ExporeOneShopExcelListV3(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma
                    ExporeOneShopExcelListV4(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                    p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭
                    
                    l_iRowIndex++;
                }


            }

            

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
        /// ////////////////////////////////////////////////////////////////////////////////////////////////

        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201231 CCL+ 不顯示科目代碼版本,改標頭,改置中 科目名稱唯一 ///////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7V4(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 1;
            const string TABLENAME = "損益表";
            //20201231 CCL-改成兩行 const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int PRINT_COLS = 2; //金額,比率

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;


            m_oWorkbook = new XLWorkbook();

            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;
            l_oRange5.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

           
            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(1, TOPCOLS + l_iPadColIndex).Address);
            l_oRange6.Merge();
            IXLRange l_oRange7 = l_oWooksheet.Range(l_oWooksheet.Cell(2, 1).Address, l_oWooksheet.Cell(2, TOPCOLS + l_iPadColIndex).Address);
            l_oRange7.Merge();

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202006月帳
            //l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            //l_oWooksheet.Cell(2, 1).Value = "多部門損益表";

            //督導名-區域-類別 損益表 (區域:北部,中部,南部;  類別:直營,合營,加盟)
            l_oWooksheet.Cell(1, 1).Value = "南部直營損益表";

            //起迄日期：2020-07-01 ~ 2020-07-31
            //l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            l_oWooksheet.Cell(2, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " 一 " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            l_oWooksheet.Cell(3, 1).Value = "科目名稱"; //改單一行唯一科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                //20201231 CCL-改單一行唯一科目名稱 int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                //int iCol3Index = iCol2Index + 1;
                //int iCol3Index = (PRINT_COLS * i) + 3;
                //int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                //l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                //20201231 CCL-改單一行唯一科目名稱 l_oWooksheet.Cell(3, iCol1Index).Value = "科目名稱";
                //實績
                //l_oWooksheet.Cell(8, iCol2Index).Value = "實績";
                //20201229 CCL* 改成金額
                l_oWooksheet.Cell(4, iCol2Index).Value = "金額";
                //比率%
                l_oWooksheet.Cell(4, iCol2Index + 1).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            /* 20201231 CCL- 科目名稱改成唯一一行
            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                //int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, iMergCol1Index).Address, l_oWooksheet.Cell(4, iMergCol1Index).Address);
                l_oRange2.Merge();
                //20201229 CCL+
                l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                //l_oRange3.Merge();
            }
            */

            //20201231 CCL+ 科目名稱改成唯一一行
            IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            l_oRange2.Merge();
            l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {

                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            //List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            MERP_AccInfoTolExcel l_oRtnTolExcelShops = new MERP_AccInfoTolExcel();
            l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init

            //取的各店Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnTolExcelShops.m_oRtnExpExcelShops.Add(l_oRtnExporeExcel);


                l_iShopNoIndex++;
            }

            //20201230 CCL+ 要先全部處理完
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {
                //20201230 CCL+ 處理Combine GID4,GID5,GID6,GID7,GID8
                l_oRtnTolExcelShops.CombineAllShopGIDAccInfo(OneShopExcel);
                
            }
            //20201230 CCL+ 重新排序
            l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();
            //20201230 CCL+
            //l_oRtnTolExcelShops.CombineAllShopAccInfoTB(); //產生All科目名稱

            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID

            //20201230 CCL+ 處理Combine GID4,GID5,GID6,GID7,GID8
            int l_iIndexNo = 0;
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {

                l_oRtnTolExcelShops.CompareNOAccInfo(OneShopExcel);
                l_oRtnTolExcelShops.SetShopNoNaToNewOrderTB(OneShopExcel, l_iIndexNo); //設店名到New Order TB
                l_iIndexNo++;
            }
            //20210103 CCL+ 計算所有店的合計
            l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //顯示輸出 //MERP_NewOrderPtrTB
            //20201231 CCL- foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            foreach (MERP_NewOrderPtrTB OneShopExcel in l_oRtnTolExcelShops.m_ShopNWOPtrTBs)
            {
                //ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                //                        p_sPROG_ID, p_oServer);


                //if (OneShopExcel.m_oBaseAttrItems.Count() > 0)
                if (OneShopExcel.m_oNOAttrItemsGID4.Count() > 0)
                {
                    //20201228 CCL- ExporeOneShopExcelListV1(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //20201229 CCL- ExporeOneShopExcelListV2(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //ExporeOneShopExcelListV3(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma
                    //ExporeOneShopExcelListV4(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭
                    //ExporeOneShopExcelListV6(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭 //科目名稱唯一一行
                    ExporeOneShopExcelListV7(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                    p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭 //科目名稱唯一一行

                    l_iRowIndex++;
                }



            }

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops, 
                                    l_oWooksheet, p_sPROG_ID, p_oServer);


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
        /// ////////////////////////////////////////////////////////////////////////////////////////////////

        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20210103 CCL+ 不顯示科目代碼版本,改標頭,改置中 科目名稱唯一 ///////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions7V5(MERP_ProcessExcelOptions p_oOption,
                                           String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 1;
            const string TABLENAME = "損益表";
            //20201231 CCL-改成兩行 const int PRINT_COLS = 3; //科目名稱,實績,比率
            const int PRINT_COLS = 2; //金額,比率
            const int TOLCOUNT_COL = 1; //合計 行//20210103 CCL+
            const int HEAD_ROWS = 4; //新版表頭數目

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            //20201227 CCL+ 看店家多少Padding多少 /////////////////////////////////////////
            int l_iShopValCount = 2; //金額一行,比率一行
            l_iShopCount = p_oOption.m_iShopCount;
            //一家店用到顯示幾行Column, p_iShopIndex以0為起始            
            int l_iPadColIndex = PRINT_COLS * l_iShopCount;


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

            //IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            //l_oRange6.Style.Font.FontName = "微軟正黑體";
            //l_oRange6.Style.Font.FontSize = 11;
            
            //Styleing
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            //20201227 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 4).Address);
            //20201231 CCL- IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, l_iPadColIndex).Address);
            //IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, TOPCOLS + l_iPadColIndex).Address);            
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            //20210105 CCL- l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontName = "標楷體";
            l_oRange.Style.Font.FontSize = 11;
            l_oRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //202006月帳
            //l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            //l_oWooksheet.Cell(2, 1).Value = "多部門損益表";

            //督導名-區域-類別 損益表 (區域:北部,中部,南部;  類別:直營,合營,加盟)
            //20210106 CCL- l_oWooksheet.Cell(1, 1).Value = "南部直營損益表";
            //20210222 CCL-去掉區域 l_oWooksheet.Cell(1, 1).Value = p_oOption.m_sManager + "  " + "南部直營損益表";
            l_oWooksheet.Cell(1, 1).Value = p_oOption.m_sManager + "  " + "損益表";

            //起迄日期：2020-07-01 ~ 2020-07-31  
            //l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            //20210105 CCL- l_oWooksheet.Cell(2, 1).Value = "" + p_oOption.m_sStartDate + " 一 " + p_oOption.m_sEndDate;
            l_oWooksheet.Cell(2, 1).Value = "" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;

            //科目代號
            //l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            l_oWooksheet.Cell(3, 1).Value = "科目名稱"; //改單一行唯一科目名稱
            //l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            //l_oWooksheet.Cell(8, 3).Value = "實績";
            //比率%
            //l_oWooksheet.Cell(8, 4).Value = "比率%";

            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                //20201231 CCL-改單一行唯一科目名稱 int iCol1Index = (PRINT_COLS * i) + 1;
                int iCol2Index = (PRINT_COLS * i) + 2;
                //int iCol3Index = iCol2Index + 1;
                //int iCol3Index = (PRINT_COLS * i) + 3;
                //int iCol4Index = (PRINT_COLS * i) + 4;

                //科目代號
                //l_oWooksheet.Cell(7, iCol1Index).Value = "科目代號";
                //科目名稱
                //20201231 CCL-改單一行唯一科目名稱 l_oWooksheet.Cell(3, iCol1Index).Value = "科目名稱";
                //實績
                //l_oWooksheet.Cell(8, iCol2Index).Value = "實績";
                //20201229 CCL* 改成金額
                l_oWooksheet.Cell(4, iCol2Index).Value = "金額";
                //比率%
                l_oWooksheet.Cell(4, iCol2Index + 1).Value = "比率%";
            }


            //IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            //l_oRange2.Merge();
            //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            //l_oRange3.Merge();

            /* 20201231 CCL- 科目名稱改成唯一一行
            for (int i = 0; i < p_oOption.m_iShopCount; i++)
            {
                int iMergCol1Index = (PRINT_COLS * i) + 1;
                //int iMergCol2Index = (PRINT_COLS * i) + 2;

                IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, iMergCol1Index).Address, l_oWooksheet.Cell(4, iMergCol1Index).Address);
                l_oRange2.Merge();
                //20201229 CCL+
                l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                //IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, iMergCol2Index).Address, l_oWooksheet.Cell(8, iMergCol2Index).Address);
                //l_oRange3.Merge();
            }
            */

            //20201231 CCL+ 科目名稱改成唯一一行
            IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(3, 1).Address, l_oWooksheet.Cell(4, 1).Address);
            l_oRange2.Merge();
            l_oRange2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            l_oRange2.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);


            //找出該店且區間的Data //改用FA_JournalV1_SqlGetDataListByOptions3找出多店舖
            //List<DataSet> l_oDataSetList = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
            List<DataSet> l_oDataSetList = new List<DataSet>();
            if (p_oOption.m_iShopCount > 0)
            {
                foreach (string shopNo in p_oOption.m_sShopList)
                {

                    p_oOption.m_sTmpShopNo = shopNo;
                    l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                    l_oDataSetList.Add(l_oDataSet);
                }
            }
            else
            {
                l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions3(p_oOption);
                l_oDataSetList.Add(l_oDataSet);
            }

            //l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);

            int l_iShopNoIndex = 0;
            //最終Expore要輸出的多店家打包物件
            //List<MERP_AccInfoExpore> l_oRtnExpExcelShops = new List<MERP_AccInfoExpore>();
            MERP_AccInfoTolExcel l_oRtnTolExcelShops = new MERP_AccInfoTolExcel();
            l_oRtnTolExcelShops.GetBaseItemsFromAccInfos(m_AccInfoDBService); //Init

            //取的各店Data
            foreach (DataSet l_oDTItemData in l_oDataSetList)
            {
                string shopNo = p_oOption.m_sShopList[l_iShopNoIndex];
                MERP_AccInfoExpore l_oRtnExporeExcel = GenOneShopExcelList(shopNo,
                                                                           l_oDTItemData,
                                                                           l_oWooksheet);
                l_oRtnTolExcelShops.m_oRtnExpExcelShops.Add(l_oRtnExporeExcel);


                l_iShopNoIndex++;
            }

            //20201230 CCL+ 要先全部處理完
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {
                //20201230 CCL+ 處理Combine GID4,GID5,GID6,GID7,GID8
                l_oRtnTolExcelShops.CombineAllShopGIDAccInfo(OneShopExcel);

            }
            //20201230 CCL+ 重新排序
            l_oRtnTolExcelShops.ReOrderAllShopGIDAccInfo();
            //20201230 CCL+
            //l_oRtnTolExcelShops.CombineAllShopAccInfoTB(); //產生All科目名稱

            //l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID

            //20201230 CCL+ 處理Combine GID4,GID5,GID6,GID7,GID8
            int l_iIndexNo = 0;
            foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            {

                l_oRtnTolExcelShops.CompareNOAccInfo(OneShopExcel);
                l_oRtnTolExcelShops.SetShopNoNaToNewOrderTB(OneShopExcel, l_iIndexNo); //設店名到New Order TB
                l_iIndexNo++;
            }
            //20210103 CCL+ 計算所有店的合計
            l_oRtnTolExcelShops.Calc_RowItemAllColAmount();

            //20201231 CCL+ 先顯示標頭和一行科目名稱
            ExporeAccountNameInfoCol(l_oWooksheet, l_oRtnTolExcelShops);

            //顯示輸出 //MERP_NewOrderPtrTB
            //20201231 CCL- foreach (MERP_AccInfoExpore OneShopExcel in l_oRtnTolExcelShops.m_oRtnExpExcelShops)
            foreach (MERP_NewOrderPtrTB OneShopExcel in l_oRtnTolExcelShops.m_ShopNWOPtrTBs)
            {
                //ExporeOneShopExcelList(l_iRowIndex, OneShopExcel, l_oWooksheet,
                //                        p_sPROG_ID, p_oServer);


                //if (OneShopExcel.m_oBaseAttrItems.Count() > 0)
                if (OneShopExcel.m_oNOAttrItemsGID4.Count() > 0)
                {
                    //20201228 CCL- ExporeOneShopExcelListV1(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //20201229 CCL- ExporeOneShopExcelListV2(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer);
                    //ExporeOneShopExcelListV3(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma
                    //ExporeOneShopExcelListV4(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭
                    //ExporeOneShopExcelListV6(l_iRowIndex, OneShopExcel, l_oWooksheet,
                    //                p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭 //科目名稱唯一一行
                    ExporeOneShopExcelListV7(l_iRowIndex, OneShopExcel, l_oWooksheet,
                                    p_sPROG_ID, p_oServer); //多Camma //置中,加改標頭 //科目名稱唯一一行

                    l_iRowIndex++;
                }



            }

            //20210103 CCL+ 顯示合計///////////////////////////////////////////////////
            ExporeAllShopColsAmount(l_iShopCount, l_oRtnTolExcelShops,
                                    l_oWooksheet, p_sPROG_ID, p_oServer);



            //20210104 CCL+, 列印的顯示邊框/////////////////////////////////////////
            int l_iTolRowCount = 0;
            l_iTolRowCount = l_oRtnTolExcelShops.m_ShopNWOPtrTBs[0].m_TBTolRowCount;
            //+1 ==> 1行空白Row
            IXLRange l_oRangeAll = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(4+l_iTolRowCount +1, TOPCOLS + l_iPadColIndex + TOLCOUNT_COL).Address);
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
        /// ////////////////////////////////////////////////////////////////////////////////////////////////



        /**** 20201227 CCL-
        //20201227 CCL+
        public static bool ExporeOneShopExcelList(int p_iShopIndex,
                                       MERP_AccInfoExpore p_oRtnExporeExcel, IXLWorksheet p_oWooksheet,
                                       String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            const int TOPCOLS = 2; //科目NO,Name 兩行
            //int l_iShopValCount = 1; //金額一行
            int l_iShopValCount = 2; //金額一行,比率一行

            //一家店用到顯示幾行Column, p_iShopIndex以0為起始
            //int l_iPadColIndex = 3 * p_iShopIndex;
            int l_iPadColIndex = 4 * p_iShopIndex;

            IXLWorksheet l_oWooksheet = p_oWooksheet;
            MERP_AccInfoExpore l_oRtnExporeExcel = p_oRtnExporeExcel;
            // 20201224 CCL 列印Excel           
            int l_iRowIndex = 0;

            //20201226 CCL+
            //顯示店名
            //部門：            
            p_oWooksheet.Cell(5, 1+ l_iPadColIndex).Value = "部門：" + l_oRtnExporeExcel.m_sShopName;
            //部門名稱
            p_oWooksheet.Cell(7, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_sShopName;
            //顯示 營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4+ l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaIncome);
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolOperaCosts;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolOperaCosts);
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "原料期初";
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComDAmount));
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oBeginRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oBeginRawMat.m_ComCAmount));

            }
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
                else if (((row.m_ComDAmount != null) && (row.m_ComDAmount != "0")) &&
                        ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")))
                {
                    //原料進料 總部 DAmount和CAmount都有值要相減
                    double tmpVal = Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = tmpVal;

                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(tmpVal);
                }

            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2+ l_iPadColIndex).Value = "減 原料期末";
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oEndRawMat.m_ComDAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComDAmount));
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3+ l_iPadColIndex).Value = l_oEndRawMat.m_ComCAmount;
                l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(l_oEndRawMat.m_ComCAmount));
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1+ l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dOperaMargin;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dOperaMargin);
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTopOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTopOperaExpense);
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBtmOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBtmOperaExpense);
            //顯示 營業利益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dBussInterest;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dBussInterest);
            //顯示 非營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaIncome;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaIncome);
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 非營業支出
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dNonOperaExpense;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dNonOperaExpense);
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComDAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComDAmount));
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = row.m_ComCAmount;
                    l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(Convert.ToDouble(row.m_ComCAmount));
                }
            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dConsuCurrentProfitLoss;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dConsuCurrentProfitLoss);
            //顯示 空白
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dTolCostOfCashExpend;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dTolCostOfCashExpend);
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1 + l_iPadColIndex).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2 + l_iPadColIndex).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3 + l_iPadColIndex).Value = l_oRtnExporeExcel.m_dCashExpendForCurrPeriod;
            l_oWooksheet.Cell(l_iRowIndex + 8, 4 + l_iPadColIndex).Value = l_oRtnExporeExcel.Calc_PercentVal(l_oRtnExporeExcel.m_dCashExpendForCurrPeriod);


            // 20201224 CCL test
            //Styling
            //IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, l_iPadColIndex + TOPCOLS + l_iShopValCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
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



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        } 
        *****/
        /////////////////////////////////////////////////////////////////////////////////////////////


        //依據對應表的屬性設定去相加,相減       
        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201224 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelV1ByOptions6(MERP_ProcessExcelOptions p_oOption,
                                       String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {
            //把Mod的砍掉Row只留下1,2,3,4,5,6,7,8; Col只留下1[科目代碼],2[科目名稱],3[店名,實績]
            //然後依店家Id找出會科
            //依店家ID讀取DB
            //Origin
            int l_iRowIndex = 0;
            int l_iShopCount = 1;
            const int TOPCOLS = 2;
            const string TABLENAME = "損益表";

            DataSet l_oDataSet = new DataSet();
            List<AccountInfo> l_oAccInfoMapTB = null;
            AccountInfo l_oTmpAccInfoItem = null;
            //要比對的DataSet            
            ///List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();


            m_oWorkbook = new XLWorkbook();
            //找出該店且區間的Data
            l_oDataSet = m_FaJournalV1DBService.FA_JournalV1_SqlGetDataListByOptions2(p_oOption);
            //抓出AccInfo Map Table全部Data


            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


            IXLRange l_oRange5 = l_oWooksheet.Range(l_oWooksheet.Cell(1, 1).Address, l_oWooksheet.Cell(2, 1).Address);
            l_oRange5.Style.Font.FontName = "微軟正黑體";
            l_oRange5.Style.Font.FontSize = 14;

            IXLRange l_oRange6 = l_oWooksheet.Range(l_oWooksheet.Cell(4, 1).Address, l_oWooksheet.Cell(5, 1).Address);
            l_oRange6.Style.Font.FontName = "微軟正黑體";
            l_oRange6.Style.Font.FontSize = 11;

            //Styleing
            IXLRange l_oRange = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 3).Address);
            l_oRange.Style.Fill.BackgroundColor = XLColor.BallBlue;
            l_oRange.Style.Font.FontColor = XLColor.White;
            l_oRange.Style.Font.FontName = "微軟正黑體";
            l_oRange.Style.Font.FontSize = 11;

            //202006月帳
            l_oWooksheet.Cell(1, 1).Value = "202006月帳";
            //多部門損益表
            l_oWooksheet.Cell(2, 1).Value = "多部門損益表";


            //起迄日期：2020-07-01 ~ 2020-07-31
            l_oWooksheet.Cell(4, 1).Value = "起迄日期：" + p_oOption.m_sStartDate + " ~ " + p_oOption.m_sEndDate;
            //科目代號
            l_oWooksheet.Cell(7, 1).Value = "科目代號";
            //科目名稱
            l_oWooksheet.Cell(7, 2).Value = "科目名稱";
            //實績
            l_oWooksheet.Cell(8, 3).Value = "實績";

            IXLRange l_oRange2 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 1).Address, l_oWooksheet.Cell(8, 1).Address);
            l_oRange2.Merge();
            IXLRange l_oRange3 = l_oWooksheet.Range(l_oWooksheet.Cell(7, 2).Address, l_oWooksheet.Cell(8, 2).Address);
            l_oRange3.Merge();

            //最終Expore要輸出的打包物件
            MERP_AccInfoExpore l_oRtnExporeExcel = new MERP_AccInfoExpore();
            l_oRtnExporeExcel.m_sShopId = p_oOption.m_sShop; //Shop ID

            TmpExcelItem l_oTmpItem = null;

            /*
            //Save To Expore Object
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;
                //Tmp 
                l_oTmpItem = new TmpExcelItem();

                //如果該row的AccNo,DtlAccNo不在AccInfo Map內,不存入 輸出打包物件
                //找出並儲存AccInfo 資訊Item
                l_oTmpAccInfoItem = m_AccInfoDBService.AccountInfo_GetDataByAccNoDtlAccNo(row.Field<string>("AccountNo"),
                                                                                          row.Field<string>("DetailAccountNo"));

                if(l_oTmpAccInfoItem != null)
                {
                    //科目代號
                    l_oTmpItem.m_ComAccountNo = row.Field<string>("AccountNo");
                    //明細科目代號
                    l_oTmpItem.m_ComDetailAccNo = row.Field<string>("DetailAccountNo");
                    //代碼全名
                    l_oTmpItem.m_FullNo = l_oTmpItem.m_ComAccountNo + "-" + l_oTmpItem.m_ComDetailAccNo;

                    //科目名稱
                    l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName");
                    //明細科目名稱
                    l_oTmpItem.m_ComDetailAccName = row.Field<string>("DetailSubjectName");
                    //名稱全名
                    l_oTmpItem.m_FullName = l_oTmpItem.m_ComAccName + " " + l_oTmpItem.m_ComDetailAccName;

                    //有在AccInfo Map內 存入 輸出打包物件
                    l_oRtnExporeExcel.m_oMapedAccInfos.Add(l_oTmpAccInfoItem);


                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpAmount = row.Field<string>("DebitAmount");
                        l_oTmpItem.m_ComDAmount = tmpAmount;

                    }
                    else
                    {
                        //貸方金額
                        tmpAmount = row.Field<string>("CreditAmount");
                        l_oTmpItem.m_ComCAmount = tmpAmount;
                    }

                    //儲存 -> Expore Object
                    l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                }

            }
            */


            //Compare 把同FullNo 不同天Day的DAmount值,CAmount值相加或新增Item到List
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                string ComAccountName = "";

                //TmpExcelItem l_oProcessedItem = null;
                l_oTmpItem = null;
                //科目代號
                string l_sAccountNo = row.Field<string>("AccountNo");
                //明細科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                //科目名稱
                string l_sAccountName = row.Field<string>("SubjectName");
                //明細科目名稱
                string l_sDetailAccountName = row.Field<string>("DetailSubjectName");
                //傳票號碼 利用傳票號碼來區分 原料期初,原料期末(都叫原料存貨)
                string l_sSubpNo = row.Field<string>("SubpNo");

                //每次找出AccNo,DtlAccNo 就抓出該Map的Info


                //把AccNo-DtlAccNo相同的(不同天日期)相加
                ///if (!string.IsNullOrEmpty(l_oProcessedItem.m_ComDetailAccNo))
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {
                    //string ComAccountNo = row.Field<string>("AccountNo");

                    /// ComAccountNo = l_oProcessedItem.m_ComAccountNo + "-" +
                    ///     l_oProcessedItem.m_ComDetailAccNo;
                    ComAccountNo = l_sAccountNo + "-" + l_sDetailAccountNo;
                    ComAccountName = l_sAccountName + " " + l_sDetailAccountName;
                }
                else
                {
                    ///ComAccountNo = l_oProcessedItem.m_ComAccountNo;
                    ComAccountNo = l_sAccountNo;
                    ComAccountName = l_sAccountName;
                }

                if (l_iRowIndex == 1)
                {

                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }

                //如果有此[科目代碼+明細科目代碼]
                //if ((l_oComDataTable.Rows.Count == 0) || (l_oComDataTable.Rows.Find(l_iRowIndex-1) == null))
                //if ( (l_oComAccountNo.Count() == 0) || (l_oComAccountNo.Contains(ComAccountNo) == false) )
                ///if ((l_oComAccNoAmount.Count() == 0) || (l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo) == null))
                if ((l_oRtnExporeExcel.m_oBaseAttrItems.Count() == 0) || (l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo) == null))
                {
                    //DataRow tmprow = new DataRow();                       
                    //

                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.NewRow();
                    //tmprow["AccountNo"] = ComAccountNo;
                    //建立新的Item
                    l_oTmpItem = new TmpExcelItem();
                    ///l_oTmpItem.m_ComAccountNo = ComAccountNo;
                    ///l_oComAccNoAmount.Add(l_oTmpItem);
                    l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                    l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                    l_oTmpItem.m_FullNo = ComAccountNo;

                    l_oTmpItem.m_ComAccName = l_sAccountName;
                    l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                    l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                    l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpAmount = row.Field<string>("DebitAmount");
                        l_oTmpItem.m_ComDAmount = tmpAmount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpAmount = row.Field<string>("CreditAmount");
                        l_oTmpItem.m_ComCAmount = tmpAmount;
                    }

                    //tmprow["Amount"] = tmpAmount;
                    //l_oComDataTable.Rows.Add(tmprow);
                    //l_oComAmount.Add(tmpAmount);
                    ///l_oTmpItem.m_ComAmount = tmpAmount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    ///l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                    l_oTmpItem.m_FullName = ComAccountName;

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";
                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.Rows.Find(ComAccountNo);
                    //tmpamount = tmprow[1].ToString();
                    //tmpamount = l_oComDataTable.Rows.Find(ComAccountNo).Field<string>(1);
                    ///l_oTmpItem = l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo);
                    ///tmpamount = l_oTmpItem.m_ComAmount;
                    l_oTmpItem = l_oRtnExporeExcel.m_oBaseAttrItems.Find(m => m.m_FullNo == ComAccountNo);

                    //20201226 CCL+ 另外特別處理AccountNo == 1192 /////////////////////////
                    if ((row.Field<string>("AccountNo") == "1192") &&
                        (row.Field<string>("SubjectName") == "原料存貨"))
                    {
                        //改成不累加金額,而是新增一個新項目,之後用SubpNo傳票編號來判斷哪個是期初;哪個是期末
                        //建立新的Item
                        l_oTmpItem = new TmpExcelItem();
                        l_oTmpItem.m_ComAccountNo = l_sAccountNo;
                        l_oTmpItem.m_ComDetailAccNo = l_sDetailAccountNo;
                        l_oTmpItem.m_FullNo = ComAccountNo;

                        l_oTmpItem.m_ComAccName = l_sAccountName;
                        l_oTmpItem.m_ComDetailAccName = l_sDetailAccountName;
                        l_oTmpItem.m_ComSubpNo = l_sSubpNo;

                        l_oRtnExporeExcel.m_oBaseAttrItems.Add(l_oTmpItem);

                        string tmpAmount = "";
                        //借方金額
                        if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                        {
                            //借方金額 D
                            tmpAmount = row.Field<string>("DebitAmount");
                            l_oTmpItem.m_ComDAmount = tmpAmount;
                        }
                        else
                        {
                            //貸方金額 C
                            tmpAmount = row.Field<string>("CreditAmount");
                            l_oTmpItem.m_ComCAmount = tmpAmount;
                        }

                        l_oTmpItem.m_FullName = ComAccountName;
                        continue;
                    }
                    ///////////////////////////////////////////////////////////////////////////

                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額 D
                        tmpnextamount = row.Field<string>("DebitAmount");
                        if (l_oTmpItem.m_ComDAmount == null)
                        { l_oTmpItem.m_ComDAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComDAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComDAmount = tmpnextamount;
                    }
                    else
                    {
                        //貸方金額 C
                        tmpnextamount = row.Field<string>("CreditAmount");
                        if (l_oTmpItem.m_ComCAmount == null)
                        { l_oTmpItem.m_ComCAmount = "0"; }
                        tmpnextamount = Convert.ToString(Int32.Parse(l_oTmpItem.m_ComCAmount) + Int32.Parse(tmpnextamount));
                        l_oTmpItem.m_ComCAmount = tmpnextamount;
                    }

                    ///tmpamount = Convert.ToString(Int32.Parse(tmpamount) + Int32.Parse(tmpnextamount));
                    ///l_oTmpItem.m_ComAmount = tmpamount; //累加值
                                                        //l_oComDataTable.Rows.Find(ComAccountNo)[1] = tmpamount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    ///l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                    l_oTmpItem.m_FullName = ComAccountName;


                }




            }

            //Grouping 分群
            l_oRtnExporeExcel.GroupingBaseItems(m_AccInfoDBService);
            l_oRtnExporeExcel.Calc_GID4_OperaIncome();
            l_oRtnExporeExcel.Calc_GID5_TolCostOfCashExpend();
            l_oRtnExporeExcel.Calc_GID5_OperaCosts();
            l_oRtnExporeExcel.Calc_GID6_OperaExpense();
            l_oRtnExporeExcel.Calc_GID7_NonOperaIncome();
            l_oRtnExporeExcel.Calc_GID8_NonOperaExpense();
            l_oRtnExporeExcel.Calc_RestOthersVal();


            // 20201224 CCL 列印Excel           
            l_iRowIndex = 0;

            //20201226 CCL+
            //顯示 營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "4";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dOperaIncome;
            //顯示 GroupID4
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID4)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }
            }
            //顯示 營業總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "5";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dTolOperaCosts;
            //顯示 原料期初
            ++l_iRowIndex;
            TmpExcelItem l_oBeginRawMat = l_oRtnExporeExcel.m_oAttrBeginRawMaterial.First();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = l_oBeginRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "原料期初";            
            if ((l_oBeginRawMat.m_ComDAmount != null) && (l_oBeginRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oBeginRawMat.m_ComDAmount;
            }
            else if ((l_oBeginRawMat.m_ComCAmount != null) && (l_oBeginRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oBeginRawMat.m_ComCAmount;
            } 
            //顯示 GroupID5
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID5)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                    
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                    
                } else if( ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0") ) &&
                          ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0")) )
                {
                    //原料進料 總部 DAmount和CAmount都有值要相減
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value =
                        Convert.ToInt32(row.m_ComDAmount) - Convert.ToInt32(row.m_ComCAmount);
                }
               
            }
            //顯示 減 原料期末
            ++l_iRowIndex;
            TmpExcelItem l_oEndRawMat = l_oRtnExporeExcel.m_oAttrEndRawMaterial.Last();
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = l_oEndRawMat.m_ComAccountNo;
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "減 原料期末";
            if ((l_oEndRawMat.m_ComDAmount != null) && (l_oEndRawMat.m_ComDAmount != "0"))
            {
                //貸方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oEndRawMat.m_ComDAmount;
            }
            else if ((l_oEndRawMat.m_ComCAmount != null) && (l_oEndRawMat.m_ComCAmount != "0"))
            {
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oEndRawMat.m_ComCAmount;
            }
            //顯示 營業毛利
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業毛利";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dOperaMargin;
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "6";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dTopOperaExpense;
            //顯示 GroupID6
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID6)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }
            }
            //顯示 營業費用
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業費用";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dBtmOperaExpense;
            //顯示 營業利益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "營業利益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dBussInterest;
            //顯示 非營業收入
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "非營業收入";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dNonOperaIncome;
            //顯示 GroupID7
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID7)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }
            }
            //顯示 非營業支出
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "8";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "非營業支出";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dNonOperaExpense;
            //顯示 GroupID8
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oAttrItemsGID8)
            {
                ++l_iRowIndex;
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                if ((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                }
                else if ((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }
            }
            //顯示 實際用量本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "7";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "實際用量本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dConsuCurrentProfitLoss;
            //顯示 空白
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = "";
            //顯示 現金支出總成本
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "現金支出總成本";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dTolCostOfCashExpend;
            //顯示 現金支出本期損益
            ++l_iRowIndex;
            l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = "";
            l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = "現金支出本期損益";
            l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = l_oRtnExporeExcel.m_dCashExpendForCurrPeriod;


            /*
            // 20201224 CCL test
            l_iRowIndex = 0;
            foreach (TmpExcelItem row in l_oRtnExporeExcel.m_oBaseAttrItems)
            {
                ++l_iRowIndex;
                //科目代號
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_FullNo;
                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_FullName;
                //20201226 CCL+ 看是DAmount有值,還是CAmount有值 就挑哪一個

                if((row.m_ComDAmount != null) && (row.m_ComDAmount != "0"))
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComDAmount;
                } else if((row.m_ComCAmount != null) && (row.m_ComCAmount != "0"))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComCAmount;
                }

            }
            */



            /*
            l_iRowIndex = 0;
            foreach (TmpExcelItem row in l_oComAccNoAmount)
            {
                ++l_iRowIndex;
                //科目代號
                l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.m_ComAccountNo;
                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.m_ComAccName;
                //借方金額
                l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.m_ComAmount;
            }
            */

            //Main
            /************************************************
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;
                //IXLCells cells = l_oWooksheet.Rows(l_iRowIndex.ToString()).Cells();
                //foreach(IXLCell cell in cells)
                //{
                //cell.Value = 
                //}


                //科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {


                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo") + "-" +
                        l_sDetailAccountNo;
                    string ComAccountNo = l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value.ToString();


                    l_oComDataTable.Rows.Add(new DataColumn(ComAccountNo));
                }
                else
                {
                    l_oWooksheet.Cell(l_iRowIndex + 8, 1).Value = row.Field<string>("AccountNo");
                }

                //科目名稱 - 明細科目名稱
                l_oWooksheet.Cell(l_iRowIndex + 8, 2).Value = row.Field<string>("SubjectName") + " " +
                    row.Field<string>("DetailSubjectName");


                if (l_iRowIndex == 1)
                {

                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }


                //借方金額
                if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("DebitAmount");
                }
                else
                {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("CreditAmount");
                }




            }
            ***********************************************/


            // 20201224 CCL test
            //Styling
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS + l_iShopCount).Address);
            l_oRange4.Style.Font.FontColor = XLColor.Black;
            l_oRange4.Style.Font.FontName = "微軟正黑體";
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



            //Testing
            Trace.WriteLine(l_oRtnExporeExcel.m_sShopId);

            return false;
        }


        /////////////////////////////////////////////////////////////////////////////////////////////

    }
}