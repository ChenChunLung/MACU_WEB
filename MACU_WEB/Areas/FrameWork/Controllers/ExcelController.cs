using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.Data;


namespace MACU_WEB.Areas.FrameWork.Controllers
{
    public class ExcelController : Controller
    {


        IXLTable m_oImpTable = null;
        DataTable m_oExpTable = null;
        
        // GET: FrameWork/Excel
        public ActionResult Index()
        {
            return View();
        }

        //20201204 CCL+ 利用ClosedXML匯入Excel檔
        [HttpPost]
        public ActionResult ImportExcel(String p_sFullFilePath)
        {
            //判斷上傳的Excel是否存在
            if (p_sFullFilePath != null)
            {
                //
                XLWorkbook l_oWorkbook = new XLWorkbook(p_sFullFilePath);

                //讀取第一個Sheet
                IXLWorksheet l_oWooksheet = l_oWorkbook.Worksheet(1);
                //定義資料起始/結束Cell
                var l_oFirstCell = l_oWooksheet.FirstCellUsed();
                var l_oEndCell = l_oWooksheet.LastCellUsed();
                //使用資料起始/結束 Cell, 來定義一個資料範圍
                var l_oData = l_oWooksheet.Range(l_oFirstCell, l_oEndCell);

                //將資料範圍轉型
                m_oImpTable = l_oData.AsTable();

                //讀取資料
                string l_sExcel = "";
                l_sExcel = m_oImpTable.Cell(6, 1).Value.ToString();

                //寫入資料
                m_oImpTable.Cell(2, 1).Value = "test";

                //資料顯示
                Response.Write("<script language=javascript>alert(" + l_sExcel + ");</script>");


                //
            }

            return View("ExcelImport");
        }


        //20201204 CCL+ 利用ClosedXML匯入Excel檔
        [HttpPost]
        public ActionResult ExportExcel(String p_sFullFilePath)
        {
            //判斷下載的Excel檔名是否存在
            if (System.IO.File.Exists(p_sFullFilePath))
            {
                //
                XLWorkbook l_oWorkbook = new XLWorkbook();

                l_oWorkbook.Worksheets.Add(m_oExpTable);
                l_oWorkbook.SaveAs(p_sFullFilePath);
            }

            return View("Excel Export");
        }

    }
}