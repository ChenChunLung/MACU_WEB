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
    public class MERP_LaborHealthExcelBIService
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
        public static MERP_FA_LaborHealthInsDBService m_LHInsDBService = new MERP_FA_LaborHealthInsDBService();


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

        //處理要輸出列印的部分
        public static bool SaveAsExcelByOptions(MERP_TCF004_JournalsOptions p_oOption,
                                          String p_sPROG_ID, HttpServerUtilityBase p_oServer)
        {


            return true;
        }

        public static string ImportExcelTo_FA_LaborHealth(FileContent p_oSearchFile)
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
                Boolean l_bIsExistData = m_LHInsDBService.FA_LaborHealthIns_ChkDataByYearMon(l_iYear, l_iMonth);
                if (l_bIsExistData)
                {
                    //刪除舊的
                    m_LHInsDBService.FA_LaborHealthIns_DBDeleteByYearMon(l_iYear, l_iMonth);
                    
                }
                //,並且匯入DataBase , 傳入手動Key In 資料年月                
                m_LHInsDBService.FA_LaborHealthIns_SqlDBCreate(m_oImpTable, l_iYear, l_iMonth); //改用ADO.NET提升速度
            }
            //return true;
            return l_sLHInsMonth;


        }

        public static List<FA_LaborHealthIns> GetImportExcelInDB_YearMonthData(string p_sYear, string p_sMonth)
        {
            int l_iYear = Convert.ToInt32(p_sYear);
            int l_iMonth = Convert.ToInt32(p_sMonth);
            return m_LHInsDBService.FA_LaborHealthIns_GetDataByYearMon(l_iYear, l_iMonth).ToList();
        }

        public static List<FA_LaborHealthIns> GetImportExcelInDB_YearMonthDataPage(string p_sYear, 
                                                                                string p_sMonth, int p_iPageing)
        {

            return m_LHInsDBService.FA_LaborHealthIns_GetDataByYearMonthPage(p_sYear, p_sMonth, p_iPageing);
        }

    }
}