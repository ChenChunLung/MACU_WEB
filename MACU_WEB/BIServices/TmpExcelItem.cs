using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MACU_WEB.BIServices
{
    public class TmpExcelItem
    {
        public string m_ComAccountNo { get; set; }

        public string m_ComDetailAccNo { get; set; }

        public string m_ComAmount { get; set; } //Sum Total總和金額

        public string m_ComAccName { get; set; }

        //Ext

        public string m_FullNo { get; set; }

        public string m_FullName { get; set; }

        public string m_ComDetailAccName { get; set; }

        public string m_ComDAmount { get; set; } //Debit 借入金額
        public string m_ComCAmount { get; set; } //Credit 貸出金額

        public string m_ComSubpNo { get; set; } //利用傳票號碼來區分 原料期初,原料期末(都叫原料存貨)

        //20201229 CCL+ 把從AccountInfo查到的PrintOrder紀錄起來,以利後續排序
        public int m_PrintOrder { get; set; }

        //20201230 CCL+ 新的重排PrintOrder
        public int m_AllComPrintOrder { get; set; }

    }
}