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

namespace MACU_WEB.BIServices
{
    public static class MERP_ExcelBIService
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
        public static MERP_FA_FaJournalDBService m_FaJournalDBService = new MERP_FA_FaJournalDBService();

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
            } else
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
            if(m_oImpTable != null)
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
                                if(l_iCellIndex > 2)
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

                if(l_sToDelIndexStr != "")
                {
                    l_sToDelIndexStr = l_sToDelIndexStr.Remove(l_sToDelIndexStr.Length-1, 1); //去除最後的","
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
            if((m_oWorkbook != null) && (m_oModTable != null))
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
            } else
            {

                return false;
            }


        }


        public static string ImportExcelToFA_DayContentDB(String p_sFullFilePath)
        {
            //載入的會計期別
            string l_sAccountPeriod = "";
            
            //載入上傳的Excel
            ImportExcelCommon(p_sFullFilePath);
            //判斷這個月會計期別是否已在DB有資料,有的話,刪除舊的
            if(m_oImpTable != null)
            {
                //從上傳Excel中抓出AccountPeriod
                l_sAccountPeriod = m_oImpTable.Cell(3, 6).Value.ToString();
                //找出資料庫是否有本月
                Boolean l_bIsExistData = m_FaJournalDBService.FA_FaJournal_FindDataByMonthVal(l_sAccountPeriod.Trim());
                if(l_bIsExistData)
                {
                    //刪除舊的
                    //m_FaJournalDBService.FA_FaJournal_DBDeleteByPeriod(l_sAccountPeriod.Trim());
                    m_FaJournalDBService.FA_FaJournal_SqlDBDeleteByPeriod(l_sAccountPeriod.Trim());
                }
                //,並且匯入DataBase
                //20201217 CCL- m_FaJournalDBService.FA_FaJournal_DBCreate(m_oImpTable);
                m_FaJournalDBService.FA_FaJournal_SqlDBCreate(m_oImpTable); //改用ADO.NET提升速度
            }
            //return true;
            return l_sAccountPeriod;
        }


        public static List<FA_FaJournal> GetImportExcelInDB_PeriodData(string p_sVal)
        {
            return m_FaJournalDBService.FA_FaJournal_GetDataByMonthVal(p_sVal);
        }

        public static List<FA_FaJournal> GetImportExcelInDB_PeriodDataPage(string p_sVal, int p_iPageing)
        {
            return m_FaJournalDBService.FA_FaJournal_GetDataByMonthValPage(p_sVal, p_iPageing);
        }




        //20201218 CCL+ For Processing 區間日期Excel 商業logical ////////////////////////////////////
        public static List<FA_FaJournal> TransDataTableToList(DataSet p_oDataSet)
        {
            return m_FaJournalDBService.FA_FaJournal_DataTableTo_FaJournalsList(p_oDataSet);
           
        }

        public static List<FA_FaJournal> ProcessImportExcelFromDB(MERP_ProcessExcelOptions p_oOption)
        {
            DataSet l_oDataSet = null;
            List<FA_FaJournal> l_RtnList = null;

            if (p_oOption != null)
            {
                l_oDataSet = m_FaJournalDBService.FA_FaJournal_SqlGetDataListByOptions(p_oOption);
                l_RtnList = m_FaJournalDBService.FA_FaJournal_DataTableTo_FaJournalsList(l_oDataSet);
                return l_RtnList;
            }
                

            return null;

        }

        //全部項目印出
        //20201221 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelByOptions(MERP_ProcessExcelOptions p_oOption,
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
            l_oDataSet = m_FaJournalDBService.FA_FaJournal_SqlGetDataListByOptions2(p_oOption);
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
                if(!string.IsNullOrEmpty(l_sDetailAccountNo))
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

                
                if(l_iRowIndex == 1)
                {
   
                    //部門：
                    l_oWooksheet.Cell(5, 1).Value = "部門：" + row.Field<string>("DepartName");
                    //部門名稱
                    l_oWooksheet.Cell(7, 3).Value = row.Field<string>("DepartName");
                }
                

                //借方金額
                if(!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                {
                    //借方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("DebitAmount");
                } else {
                    //貸方金額
                    l_oWooksheet.Cell(l_iRowIndex + 8, 3).Value = row.Field<string>("CreditAmount");
                }
                
                


            }

            //Styling
            IXLRange l_oRange4 = l_oWooksheet.Range(l_oWooksheet.Cell(9, 1).Address, l_oWooksheet.Cell(l_iRowIndex + 8, TOPCOLS+l_iShopCount).Address );
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
            } catch(Exception ex)
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
        
        //取AccountNo,AccountName,DetailAccountNo 而非取自AccountSubjects
        //20201222 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelByOptions2(MERP_ProcessExcelOptions p_oOption,
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
            //要比對的DataSet
            //List<string> l_oComAccountNo = new List<string>();
            //List<string> l_oComAmount = new List<string>();

            List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
            /*
            DataTable l_oComDataTable = new DataTable(TABLENAME);
            DataColumn column;
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountNo";
            l_oComDataTable.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "Amount";
            l_oComDataTable.Columns.Add(column);
            */

            m_oWorkbook = new XLWorkbook();
            l_oDataSet = m_FaJournalDBService.FA_FaJournal_SqlGetDataListByOptions2(p_oOption);
            
            //讀取第一個Sheet 
            IXLWorksheet l_oWooksheet = m_oWorkbook.Worksheets.Add(TABLENAME);


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

            //Compare
            foreach(DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                TmpExcelItem l_oTmpItem = null;
                //科目代號
                string l_sDetailAccountNo = row.Field<string>("DetailAccountNo");
                if (!string.IsNullOrEmpty(l_sDetailAccountNo))
                {
                    //string ComAccountNo = row.Field<string>("AccountNo");

                    ComAccountNo = row.Field<string>("AccountNo") + "-" +
                        l_sDetailAccountNo;
                }
                else
                {
                    ComAccountNo = row.Field<string>("AccountNo");
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
                if ((l_oComAccNoAmount.Count() == 0) || (l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo) == null))
                {
                    //DataRow tmprow = new DataRow();                       
                    //

                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.NewRow();
                    //tmprow["AccountNo"] = ComAccountNo;
                    l_oTmpItem = new TmpExcelItem();
                    l_oTmpItem.m_ComAccountNo = ComAccountNo;
                    l_oComAccNoAmount.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpAmount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpAmount = row.Field<string>("CreditAmount");
                    }

                    //tmprow["Amount"] = tmpAmount;
                    //l_oComDataTable.Rows.Add(tmprow);
                    //l_oComAmount.Add(tmpAmount);
                    l_oTmpItem.m_ComAmount = tmpAmount;

                    //科目名稱 - 明細科目名稱
                    l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                        row.Field<string>("DetailSubjectName");

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";
                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.Rows.Find(ComAccountNo);
                    //tmpamount = tmprow[1].ToString();
                    //tmpamount = l_oComDataTable.Rows.Find(ComAccountNo).Field<string>(1);
                    l_oTmpItem = l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo);
                    tmpamount = l_oTmpItem.m_ComAmount;


                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpnextamount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpnextamount = row.Field<string>("CreditAmount");
                    }

                    tmpamount = Convert.ToString(Int32.Parse(tmpamount) + Int32.Parse(tmpnextamount));
                    l_oTmpItem.m_ComAmount = tmpamount; //累加值
                                                        //l_oComDataTable.Rows.Find(ComAccountNo)[1] = tmpamount;

                    //科目名稱 - 明細科目名稱
                    l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                        row.Field<string>("DetailSubjectName");
                }

                
        

            }

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

            //Main
            /*
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
            */

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
            try {

                m_oWorkbook.Worksheet(1).CopyTo(m_oModWorkbook, m_oWorkbook.Worksheet(1).Name);

            } catch(Exception ex)
            {
                string message = ex.Message.ToString();
            }
            


            SaveAsExcel(p_sPROG_ID, p_oServer);//

            return false;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////

        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo
        //20201223 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelByOptions3(MERP_ProcessExcelOptions p_oOption,
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
            //要比對的DataSet            
            List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();
           

            m_oWorkbook = new XLWorkbook();
            l_oDataSet = m_FaJournalDBService.FA_FaJournal_SqlGetDataListByOptions2(p_oOption);

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

            //Compare
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                TmpExcelItem l_oTmpItem = null;
                TmpExcelItem l_oProcessedItem = null;
                //科目代號
                string l_sAccountSubjects = row.Field<string>("AccountSubjects"); //處理Subjects字串
                l_oProcessedItem = ProcessAccountSubjects(l_sAccountSubjects);

                if (!string.IsNullOrEmpty(l_oProcessedItem.m_ComDetailAccNo))
                {
                    //string ComAccountNo = row.Field<string>("AccountNo");

                    ComAccountNo = l_oProcessedItem.m_ComAccountNo + "-" +
                        l_oProcessedItem.m_ComDetailAccNo;
                }
                else
                {
                    ComAccountNo = l_oProcessedItem.m_ComAccountNo;
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
                if ((l_oComAccNoAmount.Count() == 0) || (l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo) == null))
                {
                    //DataRow tmprow = new DataRow();                       
                    //

                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.NewRow();
                    //tmprow["AccountNo"] = ComAccountNo;
                    l_oTmpItem = new TmpExcelItem();
                    l_oTmpItem.m_ComAccountNo = ComAccountNo;
                    l_oComAccNoAmount.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpAmount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpAmount = row.Field<string>("CreditAmount");
                    }

                    //tmprow["Amount"] = tmpAmount;
                    //l_oComDataTable.Rows.Add(tmprow);
                    //l_oComAmount.Add(tmpAmount);
                    l_oTmpItem.m_ComAmount = tmpAmount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";
                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.Rows.Find(ComAccountNo);
                    //tmpamount = tmprow[1].ToString();
                    //tmpamount = l_oComDataTable.Rows.Find(ComAccountNo).Field<string>(1);
                    l_oTmpItem = l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo);
                    tmpamount = l_oTmpItem.m_ComAmount;


                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpnextamount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpnextamount = row.Field<string>("CreditAmount");
                    }

                    tmpamount = Convert.ToString(Int32.Parse(tmpamount) + Int32.Parse(tmpnextamount));
                    l_oTmpItem.m_ComAmount = tmpamount; //累加值
                                                        //l_oComDataTable.Rows.Find(ComAccountNo)[1] = tmpamount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                }




            }

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

            //Main
            /*
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
            */

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

            return false;
        }

        //從AccountSubjects中分離出AccountNo, DetailAccountNo, AccountName
        public static TmpExcelItem ProcessAccountSubjects(string p_sAccSubjects)
        {
            TmpExcelItem l_oTmpItem = new TmpExcelItem();
            string l_sTmpStr = "";
            //string l_sAccountNo = "", l_sDetailAccountNo = "", l_sAccountName = "";

            //如果空格前面只有4碼
            l_sTmpStr = p_sAccSubjects.Substring(0,p_sAccSubjects.IndexOf(" ") );
            //4碼
            l_oTmpItem.m_ComAccountNo = p_sAccSubjects.Substring(0, 4);
            if(l_sTmpStr.Length > 4)
            {
                //如果空格前面大於4碼,便要取出AccountNo
                l_oTmpItem.m_ComDetailAccNo = p_sAccSubjects.Substring(4, l_sTmpStr.Length-4);
            }
            l_oTmpItem.m_ComAccName = p_sAccSubjects.Substring(p_sAccSubjects.IndexOf(" ")+1).ToString();


            return l_oTmpItem;
        }
        /////////////////////////////////////////////////////////////////////////////////////////////


        //改從AccountSubjects字串取得AccountNo,AccountName,DetailAccountNo; Amount金額改從LocalCurrencyAmount
        //20201223 CCL+ /////////////////////////////////////////////////////////////////////////////
        public static bool SaveAsExcelByOptions4(MERP_ProcessExcelOptions p_oOption,
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
            //要比對的DataSet            
            List<TmpExcelItem> l_oComAccNoAmount = new List<TmpExcelItem>();


            m_oWorkbook = new XLWorkbook();
            l_oDataSet = m_FaJournalDBService.FA_FaJournal_SqlGetDataListByOptions2(p_oOption);

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

            //Compare
            foreach (DataRow row in l_oDataSet.Tables[0].Rows)
            {
                ++l_iRowIndex;

                string ComAccountNo = "";
                TmpExcelItem l_oTmpItem = null;
                TmpExcelItem l_oProcessedItem = null;
                //科目代號
                string l_sAccountSubjects = row.Field<string>("AccountSubjects"); //處理Subjects字串
                l_oProcessedItem = ProcessAccountSubjects(l_sAccountSubjects);

                if (!string.IsNullOrEmpty(l_oProcessedItem.m_ComDetailAccNo))
                {
                    //string ComAccountNo = row.Field<string>("AccountNo");

                    ComAccountNo = l_oProcessedItem.m_ComAccountNo + "-" +
                        l_oProcessedItem.m_ComDetailAccNo;
                }
                else
                {
                    ComAccountNo = l_oProcessedItem.m_ComAccountNo;
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
                if ((l_oComAccNoAmount.Count() == 0) || (l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo) == null))
                {
                    //DataRow tmprow = new DataRow();                       
                    //

                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.NewRow();
                    //tmprow["AccountNo"] = ComAccountNo;
                    l_oTmpItem = new TmpExcelItem();
                    l_oTmpItem.m_ComAccountNo = ComAccountNo;
                    l_oComAccNoAmount.Add(l_oTmpItem);

                    string tmpAmount = "";
                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpAmount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpAmount = row.Field<string>("CreditAmount");
                    }

                    //tmprow["Amount"] = tmpAmount;
                    //l_oComDataTable.Rows.Add(tmprow);
                    //l_oComAmount.Add(tmpAmount);
                    l_oTmpItem.m_ComAmount = tmpAmount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;

                }
                else
                {
                    //找出該[科目代碼+明細科目代碼]的金額累加
                    string tmpamount = "", tmpnextamount = "";
                    //DataRow tmprow;
                    //tmprow = l_oComDataTable.Rows.Find(ComAccountNo);
                    //tmpamount = tmprow[1].ToString();
                    //tmpamount = l_oComDataTable.Rows.Find(ComAccountNo).Field<string>(1);
                    l_oTmpItem = l_oComAccNoAmount.Find(m => m.m_ComAccountNo == ComAccountNo);
                    tmpamount = l_oTmpItem.m_ComAmount;


                    //借方金額
                    if (!string.IsNullOrEmpty(row.Field<string>("DebitAmount")))
                    {
                        //借方金額
                        tmpnextamount = row.Field<string>("DebitAmount");

                    }
                    else
                    {
                        //貸方金額
                        tmpnextamount = row.Field<string>("CreditAmount");
                    }

                    tmpamount = Convert.ToString(Int32.Parse(tmpamount) + Int32.Parse(tmpnextamount));
                    l_oTmpItem.m_ComAmount = tmpamount; //累加值
                                                        //l_oComDataTable.Rows.Find(ComAccountNo)[1] = tmpamount;

                    //科目名稱 - 明細科目名稱
                    //l_oTmpItem.m_ComAccName = row.Field<string>("SubjectName") + " " +
                    //    row.Field<string>("DetailSubjectName");
                    l_oTmpItem.m_ComAccName = l_oProcessedItem.m_ComAccName;
                }




            }

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

            //Main
            /*
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
            */

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

            return false;
        }

    
        /////////////////////////////////////////////////////////////////////////////////////////////


    }
}