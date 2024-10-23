using Example.NPOI;
using static Example.NPOI.ExcelOutputReport;
using System.Collections;
using System.Data;


#region new Logic
var dt = new DataTable();
var loopsheetDatas = new List<ExcelSheetRegion>();
if (dt.Rows.Count > 0)
{
    //正常交易區
    loopsheetDatas.Add(new ExcelSheetRegion
    {
        LoopSheetName = "Loop_101",
        LoopDataTables = new DataTable[] { dt },
    });
}


//Header Footer 共用變數
var arrayList = new ArrayList();
arrayList.Add(new string[] { "HEADER_1",  "Example - Test" });
arrayList.Add(new string[] { "HEADER_2", "戶名：00000000000(業務代碼：122)" });
arrayList.Add(new string[] { "HEADER_3", "製表日期：" + DateTime.Now.ToShortDateString() });
arrayList.Add(new string[] { "FOOTER_CENTER", "&C" + "101" + " 年度第 " + "2" + "  學期  第 &P 頁，共 &N 頁" });

ExcelSheetObject excelSheetObject = new ExcelSheetObject
{
    OutputSheetName = "101報表",
    LoopSheetData = loopsheetDatas
};

ExcelOutputReport excelOutputReport = new ExcelOutputReport();
string errMsg = "";
var output = excelOutputReport.OutputNPOIFees("101", arrayList, excelSheetObject, "rpt101", "ExcelOut", ref errMsg, "", FileType.xls.GetHashCode());  //check

#endregion