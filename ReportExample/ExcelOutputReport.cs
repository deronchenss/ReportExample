using EnumsNET;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;

///<summary>
///NPOI 報表
///<remarks>Create</remarks>
///<history>
///</history>
///</summary>

namespace Example.NPOI
{
    public static class EnumExtension
    {
        public static string GetDisplayName(this Enum value)
        {
            var displayAttribute = value.GetType()
                .GetField(value.ToString())
                ?.GetCustomAttributes(typeof(DisplayAttribute), false)
                .OfType<DisplayAttribute>()
                .FirstOrDefault();

            return displayAttribute?.Name ?? value.ToString();
        }
    }

    public class ExcelOutputReport
    {
        private int maxRow = 65536;
        private IWorkbook workbook;
        private int gi_ReportCount = 0;
        private NegativeNumberColorEnum g_negativeNumberColor = NegativeNumberColorEnum.Red;
        private Dictionary<string, Dictionary<int, ICellStyle>> loopDic = new Dictionary<string, Dictionary<int, ICellStyle>>();
        private Dictionary<string, Dictionary<string, ICellStyle>> dynamicDic = new Dictionary<string, Dictionary<string, ICellStyle>>();

        public enum FileType
        {
            [Display(Name = ".xls")]
            xls = 0,


            [Display(Name = ".xlsx")]
            xlsx = 1
        }
        private string negativeNumberColorStr
        {
            get
            {
                string colorStr = "";
                if(g_negativeNumberColor != NegativeNumberColorEnum.Black) //黑色直接給空字串即可
                {
                    //Red => [Red]
                    colorStr = $"[{Enum.GetName(g_negativeNumberColor.GetType(), g_negativeNumberColor)}]";
                }

                return colorStr;
            }
        }

        public int GetReportCount
        {
            get { return gi_ReportCount; }
        }

        /// <summary>
        /// 負值的顏色
        /// </summary>
        public NegativeNumberColorEnum NegativeNumberColor
        {
            get { return g_negativeNumberColor; }
            set { g_negativeNumberColor = value; }
        }

        public class ExcelSheetObject
        {
            /// <summary>
            /// 一定要有 沒有也要有空白頁
            /// </summary>
            public string OutputSheetName { get; set; }

            public List<ExcelSheetRegion> LoopSheetData { get; set; } = new List<ExcelSheetRegion>();

            /// <summary>
            /// 是否重製動態欄位寬度
            /// </summary>
            public bool isAutoDynamicSizeColumn { get; set; } = false;

            public bool isNeedReMergeHeader { get; set; } = false;

            public int headerDefaultColumn { get; set; } = 1;

        }

        /// <summary>
        /// 含
        /// </summary>
        public class ExcelSheetRegion
        {
            public string LoopSheetName { get; set; }
            public DataTable[] LoopDataTables { get; set; } = new DataTable[0];
            public ArrayList[] LoopArrayLists { get; set; } = new ArrayList[0];

            public Dictionary<string, List<string>> DicFees = new Dictionary<string, List<string>>();

            /// <summary>
            /// Looper 加一行空白
            /// </summary>
            public bool isLooperAddSpace { get; set; } = true;

            /// <summary>
            /// Footer 加一行空白
            /// </summary>
            public bool isFooterAddSpace { get; set; } = true;

        }


        //一般報表 by NPOI 可以動態產生欄位
        /// <summary>
        /// 一般報表，指定Templete Sheet，將來源DataTable產出至工作表中
        /// </summary>
        /// <param name="rptName">產出報表名稱</param>
        /// <param name="templatePath">範本報表路徑</param>
        /// <param name="outputPath">產出報表路徑</param>
        /// <param name="errMsg">傳出的錯誤訊息</param>
        /// <param name="programId">程式代號</param>
        /// <param name="extension"></param>
        public string[] OutputNPOIFees(string rptName, ArrayList otherAllData, ExcelSheetObject excelSheetObject, string templateName, string outputPath, 
                    ref string errMsg, string programId, int? fileType)
        {
            var result = new List<string>();
            string outputFileName = "";
            var datatable = new DataTable();
            int detailRow = 0, tempMAXROW = 0;
            bool isHasDynamicField = false;
            bool isAlreadyMerge = false; //根據動態欄位調整前六欄 只有第一次要

            var extension = FileType.xls;//預設xls
            if (fileType.HasValue && Enum.IsDefined(typeof(FileType), fileType.Value))
            {
                extension = (FileType)fileType;
            }

            int intFileNum = 1;

            string templatePath = Path.Combine("Template", extension.GetDisplayName().Replace(".", ""), templateName + extension.GetDisplayName());
            FileStream outputFile = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            GetWorkBook(templatePath, ref outputFile);

            try
            {
                //Step1 處理Header 
                ISheet outputSheet = GetNewHeader(otherAllData, excelSheetObject, programId, ref datatable, ref detailRow, ref tempMAXROW);

                //Step2 處理Detail 
                if (excelSheetObject.LoopSheetData.Count > 0)
                {
                    foreach (var data in excelSheetObject.LoopSheetData)
                    {
                        if (!string.IsNullOrEmpty(data.LoopSheetName))
                        {
                            string templateSheet = data.LoopSheetName;
                            var tmpSheetName = templateSheet + "_TMP";

                            ArrayList tmpTableNames = new ArrayList();
                            tmpTableNames.Add(tmpSheetName);

                            //處理單個datatable太大的話 分EXCEL
                            var tablesToInsert = new Dictionary<int, List<DataTable>>();
                            var sourceTables = data.LoopDataTables.ToList();
                            if (extension == FileType.xls)
                            {
                                var fixCnt = workbook.GetSheet(templateSheet).LastRowNum;
                                for (int i = 0; i < data.LoopDataTables.Count(); i++)
                                {
                                    var sourceTable = sourceTables[i];
                                    if (sourceTable.Rows.Count + 6 + fixCnt < maxRow)
                                    {
                                        continue;
                                    }

                                    var splitTables = new List<DataTable>();
                                    for (int startRow = 0; startRow < sourceTable.Rows.Count; startRow += (maxRow - 6 - fixCnt))
                                    {
                                        DataTable newTable = sourceTable.Clone();
                                        int endRow = Math.Min(startRow + maxRow - 6 - fixCnt, sourceTable.Rows.Count);
                                        for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                                        {
                                            newTable.ImportRow(sourceTable.Rows[rowIndex]);
                                        }
                                        splitTables.Add(newTable);
                                    }

                                    tablesToInsert.Add(i, splitTables);
                                }

                                var addCnt = 0;
                                foreach (var dic in tablesToInsert)
                                {
                                    sourceTables.RemoveAt(dic.Key + addCnt); 
                                    sourceTables.InsertRange(dic.Key + addCnt, dic.Value);
                                    addCnt = addCnt + dic.Value.Count - 1;
                                }
                            }

                            //迴圈建立Sheet 新增Data 複製到輸出sheet後刪除
                            for (int i = 0; i < sourceTables.Count(); i++)
                            {
                                var sourceTable = sourceTables[i];
                                if (sourceTable.Rows.Count > 0)
                                {
                                    workbook.CloneSheet(workbook.GetSheetIndex(templateSheet));
                                    workbook.SetSheetName(workbook.NumberOfSheets - 1, tmpSheetName);
                                    workbook.GetSheet(tmpSheetName).ForceFormulaRecalculation = true;

                                    var nextRow = outputSheet.LastRowNum + sourceTable.Rows.Count + workbook.GetSheet(tmpSheetName).LastRowNum;
                                    if (extension == FileType.xls && nextRow > maxRow)
                                    {
                                        outputFileName = MakeExcel(excelSheetObject, outputPath, $"{rptName}_{intFileNum.ToString("00")}{extension.GetDisplayName()}", ref outputSheet, ref intFileNum, templatePath);
                                        GetWorkBook(templatePath, ref outputFile);
                                        outputSheet = GetNewHeader(otherAllData, excelSheetObject, programId, ref datatable, ref detailRow, ref tempMAXROW);
                                        isAlreadyMerge = false;

                                        workbook.CloneSheet(workbook.GetSheetIndex(templateSheet));
                                        workbook.SetSheetName(workbook.NumberOfSheets - 1, tmpSheetName);
                                        workbook.GetSheet(tmpSheetName).ForceFormulaRecalculation = true;

                                        dynamicDic = new Dictionary<string, Dictionary<string, ICellStyle>>();
                                        loopDic = new Dictionary<string, Dictionary<int, ICellStyle>>();

                                        result.Add(outputFileName);
                                    }

                                    ISheet tmpSheet = workbook.GetSheet(tmpSheetName);

                                    ReplaceTemplateDynamic(ref sourceTable, ref tmpSheet, ref detailRow, ref tempMAXROW, programId, ref outputSheet, data, ref isHasDynamicField);

                                    if (!isAlreadyMerge && excelSheetObject.isNeedReMergeHeader)
                                    {
                                        for (var headerCnt = 0; headerCnt < 6; headerCnt++)
                                        {
                                            AddMergedCellWithCreate(outputSheet, headerCnt, 0, tmpSheet.GetRow(0).LastCellNum - 1);
                                        }
                                        isAlreadyMerge = true;
                                    }

                                    #region 預先處理特殊欄位及依資料筆數將sheet準備好

                                    var dataArray = data.LoopArrayLists != null && data.LoopArrayLists.Length >= i + 1 ? data.LoopArrayLists[i] : new ArrayList();

                                    //獲取TemplateSheet內容，替換所有變數欄位
                                    GetTemplateData(ref sourceTable, ref tmpSheet, ref detailRow, ref tempMAXROW, dataArray, programId);
                                    #endregion

                                    var lastRowNum = outputSheet.LastRowNum + (i == 0 ? 0 : (data.isLooperAddSpace ? 1 : 0));
                                    for (int j = 0; j < tmpSheet.LastRowNum + 1; j++)
                                    {
                                        lastRowNum += 1;
                                        CopyRow(ref tmpSheet, j, ref outputSheet, lastRowNum, copyRowHeight: false);
                                    }

                                    DeleteTags(ref tmpSheet);
                                    this.DeleteSheet(ref workbook, tmpTableNames, ref errMsg);

                                    outputSheet.SetRowBreak(outputSheet.LastRowNum);
                                }
                            }

                            //沒有loop的時候
                            //因為Loop沒有值得時候應該不顯示 但有些報表會有Array 先這樣解決
                            if (data.LoopDataTables.Count() == 0)
                            {
                                var loopArrayList = data.LoopArrayLists.Count() == 0 ? new ArrayList[] { new ArrayList() } : data.LoopArrayLists;
                                for (int i = 0; i < loopArrayList.Count(); i++)
                                {
                                    var nextRow = outputSheet.LastRowNum + workbook.GetSheet(data.LoopSheetName).LastRowNum;
                                    if (extension == FileType.xls && nextRow > maxRow)
                                    {
                                        outputFileName = MakeExcel(excelSheetObject, outputPath, $"{rptName}_{intFileNum.ToString("00")}{extension.GetDisplayName()}", ref outputSheet, ref intFileNum, templatePath);
                                        GetWorkBook(templatePath, ref outputFile);
                                        outputSheet = GetNewHeader(otherAllData, excelSheetObject, programId, ref datatable, ref detailRow, ref tempMAXROW);
                                        result.Add(outputFileName);

                                        dynamicDic = new Dictionary<string, Dictionary<string, ICellStyle>>();
                                        loopDic = new Dictionary<string, Dictionary<int, ICellStyle>>();
                                    }

                                    ISheet fixSheet = workbook.GetSheet(data.LoopSheetName);
                                    ReplaceTemplateDynamic(ref datatable, ref fixSheet, ref detailRow, ref tempMAXROW, programId, ref outputSheet, data, ref isHasDynamicField);

                                    GetTemplateData(ref datatable, ref fixSheet, ref detailRow, ref tempMAXROW, loopArrayList[i], programId);

                                    var lastRowNum = outputSheet.LastRowNum + (data.isFooterAddSpace ? 1 : 0);
                                    for (int j = 0; j < fixSheet.LastRowNum + 1; j++)
                                    {
                                        lastRowNum += 1;
                                        CopyRow(ref fixSheet, j, ref outputSheet, lastRowNum, copyRowHeight: false);
                                    }
                                }
                            }
                        }
                    }

                    if (!isAlreadyMerge && excelSheetObject.isNeedReMergeHeader)
                    {
                        for (var headerCnt = 0; headerCnt < 6; headerCnt++)
                        {
                            AddMergedCellWithCreate(outputSheet, headerCnt, 0, excelSheetObject.headerDefaultColumn);
                        }
                        isAlreadyMerge = true;
                    }

                    outputFileName = MakeExcel(excelSheetObject, outputPath, $"{rptName}_{intFileNum.ToString("00")}{extension.GetDisplayName()}", ref outputSheet, ref intFileNum, templatePath);
                    result.Add(outputFileName);
                }
                else
                {
                    if (!isAlreadyMerge && excelSheetObject.isNeedReMergeHeader)
                    {
                        for (var headerCnt = 0; headerCnt < 6; headerCnt++)
                        {
                            AddMergedCellWithCreate(outputSheet, headerCnt, 0, excelSheetObject.headerDefaultColumn);
                        }
                        isAlreadyMerge = true;
                    }

                    outputFileName = MakeExcel(excelSheetObject, outputPath, $"{rptName}_{intFileNum.ToString("00")}{extension.GetDisplayName()}", ref outputSheet, ref intFileNum, templatePath);
                    result.Add(outputFileName);
                }

                if (!string.IsNullOrEmpty(errMsg))
                {
                    throw new Exception(errMsg);
                }

                return result.ToArray();
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (outputFile != null)
                {
                    outputFile.Close();
                    outputFile.Dispose();
                }
            }
        }

        private ISheet GetNewHeader(ArrayList otherAllData, ExcelSheetObject excelSheetObject, string programId, ref DataTable datatable, ref int detailRow, ref int tempMAXROW)
        {
            ISheet outputSheet = workbook.GetSheet(excelSheetObject.OutputSheetName);
            GetTemplateData(ref datatable, ref outputSheet, ref detailRow, ref tempMAXROW, otherAllData, programId);
            return outputSheet;
        }

        private void GetWorkBook(string templatePath, ref FileStream outputFile)
        {
            if (outputFile != null)
            {
                outputFile.Close();
                outputFile.Dispose();
            }

            // 讀取範本EXCEL
            outputFile = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            workbook = WorkbookFactory.Create(outputFile);
        }

        private string MakeExcel(ExcelSheetObject excelSheetObject, string outputPath, string outputFileName, ref ISheet outputSheet, ref int intFileNum, string templatePath)
        {
            this.DeleteTags(ref outputSheet);

            ArrayList tableNames = new ArrayList();
            tableNames.Add(excelSheetObject.OutputSheetName);
            this.DeleteOtherSheet(ref workbook, tableNames);

            if (excelSheetObject.isAutoDynamicSizeColumn)
            {
                for (int j = 0; j < outputSheet.LastRowNum; j++)
                {
                    outputSheet.AutoSizeColumn(j);
                }
            }

            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }

            var outputFullPath = Path.Combine(outputPath, outputFileName);
            if (File.Exists(outputFullPath))
                File.Delete(outputFullPath);

            FileStream fileOut = new FileStream(outputFullPath, FileMode.Create);
            workbook.Write(fileOut);
            //fileOut.Flush();
            fileOut.Close();
            fileOut.Dispose();

            intFileNum++;

            return outputFileName;
        }
       
        public IWorkbook GetWorkbook(string s_aTemplatePath)
        {
            // 讀取範本EXCEL
            using (FileStream o_File = new FileStream(s_aTemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //取得範本
                workbook = WorkbookFactory.Create(o_File);
            }

            return workbook;
        }

        #region Private Method for Sheet
        public void ReplaceTemplateDynamic(ref System.Data.DataTable datas, ref ISheet wsTemplate, ref int detailRow, ref int tempMAXROW, string programId, ref ISheet outputTemplate, ExcelSheetRegion excelSheetRegion, ref bool isHasDynamicField)
        {
            Dictionary<string, List<string>> dicFees = excelSheetRegion.DicFees;
            if (dicFees == null || dicFees.Count() == 0)
            {
                return;
            }

            //設定搜尋資料範圍
            //獲取Excel Sheet中最大的可用列數
            tempMAXROW = wsTemplate.LastRowNum;
            detailRow = 0;

            var replaceCnt = 0;
            for (int i = 0; i <= tempMAXROW + datas.Rows.Count; i++)
            {
                if (wsTemplate.GetRow(i) == null)
                { continue; }
                int cols = wsTemplate.GetRow(i).LastCellNum;
                for (int j = 0; j < cols; j++)
                {
                    if (wsTemplate.GetRow(i).GetCell(j) != null)
                    {
                        if (wsTemplate.GetRow(i).GetCell(j).ToString() == "")
                        { continue; }

                        if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##DYNAMICFIELD") != -1)
                        {
                            isHasDynamicField = true;

                            //先建立需要的新欄位
                            for (int moveX = 0; moveX < dicFees.Count; moveX++)
                            {
                                //因為DynamicField本來就占一格 所以 0 不用建立
                                if (moveX == 0)
                                {
                                    continue;
                                }

                                wsTemplate.GetRow(i).CreateCell(cols + moveX - 1);
                                CopyCell(wsTemplate.GetRow(i).GetCell(j), wsTemplate.GetRow(i).GetCell(cols + moveX - 1), wsTemplate.SheetName, workbook);
                            }

                            //將現有欄位向右移
                            var fixField = cols - j - 1;
                            for (int moveX = fixField; moveX > 0; moveX--)
                            {
                                CopyCell(wsTemplate.GetRow(i).GetCell(j + moveX), wsTemplate.GetRow(i).GetCell(j + moveX + dicFees.Count - 1));
                            }

                            //將動態欄位賦值
                            var cnt = 0;
                            foreach (var dicFee in dicFees)
                            {
                                wsTemplate.GetRow(i).GetCell(j + cnt).SetCellValue(dicFee.Value.Count <= replaceCnt ? "" : dicFee.Value[replaceCnt]);
                                cnt++;
                            }

                            replaceCnt++;
                            break;
                        }
                    }
                }
            }
          
            return;
        }


        /// <summary>
        /// GetTemplateData
        /// 獲取Template Sheet的數據
        /// </summary>
        /// <param name="wsTemplate">報表Template Sheet</param>
        /// <param name="detailRow">DETAIL開始的列</param>
        /// <returns>獲取Excel Sheet中數據</returns>
        public void GetTemplateData(ref System.Data.DataTable datas, ref ISheet wsTemplate, ref int detailRow, ref int tempMAXROW, ArrayList sumDatas, string programId)
        {
            //設定搜尋資料範圍
            //獲取Excel Sheet中最大的可用列數
            tempMAXROW = wsTemplate.LastRowNum;
            detailRow = 0;

            for (int i = 0; i <= tempMAXROW + datas.Rows.Count; i++)
            {
                if (wsTemplate.GetRow(i) == null)
                { continue; }
                int i_Cols = wsTemplate.GetRow(i).LastCellNum;
                for (int j = 0; j < i_Cols; j++)
                {
                    if (wsTemplate.GetRow(i).GetCell(j) != null)
                    {
                        if (wsTemplate.GetRow(i).GetCell(j).ToString() == "")
                        { continue; }

                        if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##FIELD_") != -1)
                        {
                            //取得Template之欄位名稱
                            // 設定的類型如下
                            // 類型1: ##FIELD_<ColumnName>:I
                            // 類型2: ##FIELD_ColumnName:I
                            string s_TemplateColumnName = wsTemplate.GetRow(i).GetCell(j).ToString();
                            if (s_TemplateColumnName.IndexOf("<") != -1 && s_TemplateColumnName.IndexOf(">") != -1)
                            {
                                // 類型1: ##FIELD_<ColumnName>:I
                                int i_startIndex = s_TemplateColumnName.IndexOf("<");
                                int i_endIndex = s_TemplateColumnName.IndexOf(">");
                                s_TemplateColumnName = s_TemplateColumnName.Substring(i_startIndex + 1, i_endIndex - i_startIndex - 1);
                            }
                            else
                            {
                                // 類型2: ##FIELD_ColumnName:I
                                s_TemplateColumnName = s_TemplateColumnName.Replace("##FIELD_", "");
                                int i_endIndex = s_TemplateColumnName.IndexOf(":");
                                if (i_endIndex >= 0)
                                    s_TemplateColumnName = s_TemplateColumnName.Substring(0, i_endIndex);
                            }

                            for (int k = 0; k < sumDatas.Count; k++)
                            {
                                String[] temp = (String[])sumDatas[k];

                                if (s_TemplateColumnName == temp[0])
                                {
                                    if (temp[1] == string.Empty)
                                    {
                                        wsTemplate.GetRow(i).GetCell(j).SetCellValue("");
                                    }
                                    else
                                    {
                                        SetTagData(wsTemplate, i, j, temp[1]);
                                    }
                                    break;
                                }
                            }
                        }

                        string fieldvalue = wsTemplate.GetRow(i).GetCell(j).ToString().Trim();
                        if (fieldvalue.IndexOf("##LOOPFIELD_") != -1)
                        {
                            detailRow = SetRowData(datas, wsTemplate, sumDatas, programId, tempMAXROW, i, j, programId);  //填入主要報表內容
                        }

                        if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##SUM=") != -1)
                        {
                            SetRowSumFormulla(datas, wsTemplate, sumDatas, programId, i, j);    // 填入橫列的加總
                        }
                    }
                }
            }
            //改成外面自己呼叫 有可能有多次呼叫此function
            //SetHeaderFooter(ref wsTemplate, a_aSumDatas, s_aProgramId);   // 設定頁首頁尾
            return;
        }
      
        /// <summary>
        /// SetHeaderFooter
        /// 抓取頁首/頁尾
        /// </summary>
        /// <param name="wsTemplate">工作表</param>
        /// <param name="a_aSumDatas">從第幾列開始插入列</param>
        /// <param name="s_aProgramId">共需插入幾列</param>
        /// <returns></returns>
        private void SetHeaderFooter(ref ISheet wsTemplate, ArrayList a_aSumDatas, string s_aProgramId)
        {
            String[] a_HeaderCenter = wsTemplate.Header.Center.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            String[] a_HeaderLeft = wsTemplate.Header.Left.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            String[] a_HeaderRight = wsTemplate.Header.Right.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            String[] a_FooterCenter = wsTemplate.Footer.Center.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            String[] a_FooterLeft = wsTemplate.Footer.Left.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            String[] a_FooterRight = wsTemplate.Footer.Right.ToString().Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            String s_ReplaceHeaderCenter = String.Empty;
            String s_ReplaceHeaderLeft = String.Empty;
            String s_ReplaceHeaderRight = String.Empty;
            String s_ReplaceFooterCenter = String.Empty;
            String s_ReplaceFooterLeft = String.Empty;
            String s_ReplaceFooterRight = String.Empty;

            for (int i = 0; i < a_aSumDatas.Count; i++)
            {
                String[] s_Data = (String[])a_aSumDatas[i];

                for (int j = 0; j < a_HeaderCenter.Length; j++)
                {
                    if (a_HeaderCenter[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceHeaderCenter = s_ReplaceHeaderCenter + s_Data[1].ToString() + "\n";
                    }
                }
                for (int j = 0; j < a_HeaderLeft.Length; j++)
                {
                    if (a_HeaderLeft[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceHeaderLeft = s_ReplaceHeaderLeft + s_Data[1].ToString() + "\n";
                    }
                }
                for (int j = 0; j < a_HeaderRight.Length; j++)
                {
                    if (a_HeaderRight[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceHeaderRight = s_ReplaceHeaderRight + s_Data[1].ToString() + "\n";
                    }
                }
                for (int j = 0; j < a_FooterCenter.Length; j++)
                {
                    if (a_FooterCenter[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceFooterCenter = s_ReplaceFooterCenter + s_Data[1].ToString() + "\n";
                    }
                }
                for (int j = 0; j < a_FooterLeft.Length; j++)
                {
                    if (a_FooterLeft[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceFooterLeft = s_ReplaceFooterLeft + s_Data[1].ToString() + "\n";
                    }
                }
                for (int j = 0; j < a_FooterRight.Length; j++)
                {
                    if (a_FooterRight[j] == ("##FIELD_" + s_aProgramId + "_" + s_Data[0]))
                    {
                        s_ReplaceFooterRight = s_ReplaceFooterRight + s_Data[1].ToString() + "\n";
                    }
                }
            }

            if (a_HeaderCenter.Length > 0)
            {
                wsTemplate.Header.Center = wsTemplate.Header.Center.ToString().Replace(wsTemplate.Header.Center.ToString(), s_ReplaceHeaderCenter);
            }
            if (a_HeaderLeft.Length > 0)
            {
                wsTemplate.Header.Left = wsTemplate.Header.Left.ToString().Replace(wsTemplate.Header.Left.ToString(), s_ReplaceHeaderLeft);
            }
            if (a_HeaderRight.Length > 0)
            {
                wsTemplate.Header.Right = wsTemplate.Header.Right.ToString().Replace(wsTemplate.Header.Right.ToString(), s_ReplaceHeaderRight);
            }
            if (a_FooterCenter.Length > 0)
            {
                wsTemplate.Footer.Center = wsTemplate.Footer.Center.ToString().Replace(wsTemplate.Footer.Center.ToString(), s_ReplaceFooterCenter);
            }
            if (a_FooterLeft.Length > 0)
            {
                wsTemplate.Footer.Left = wsTemplate.Footer.Left.ToString().Replace(wsTemplate.Footer.Left.ToString(), s_ReplaceFooterLeft);
            }
            if (a_FooterLeft.Length > 0)
            {
                wsTemplate.Footer.Right = wsTemplate.Footer.Right.ToString().Replace(wsTemplate.Footer.Right.ToString(), s_ReplaceFooterRight);
            }
        }

        /// <summary>
        /// SetRowData
        /// 填入資料表內容
        /// </summary>
        /// <param name="wsTemplate">工作表</param>
        /// <param name="a_aSumDatas">從第幾列開始插入列</param>
        /// <param name="s_aProgramId">共需插入幾列</param>
        /// <param name="i">跑到第幾ROW</param>
        /// <param name="j">跑到第幾COLUMN</param>
        /// <param name="i_aTempMAXROW"></param>
        /// <param name="ProgramId">如有客製化需求，傳入的參數</param>
        /// <returns></returns>
        private int SetRowData(System.Data.DataTable datas, ISheet wsTemplate, ArrayList a_aSumDatas, string s_aProgramId, int i_aTempMAXROW, int i, int j, string ProgramId = null)
        {
            int i_DataCount = 0;
            int decimalPlace = 0;
            Double d_Sum = 0.0;
            bool chkNumeric = false;  // 判斷是否為數值
            bool chkSum = false;  //  判斷是否需要加總
            bool chkDate = false;  // 判斷是否為日期格式
            bool chkInteger = false;  // 判斷是否為整數
            bool isEmpty = false;
            //bool b_space = false;
            string nullDisplayValue = null; //設定Null的帶入值，預設為null
            bool percent = false;
            int style = 0;  // 判斷欄位樣式
            string color = ""; //判斷負數顏色是否有規定,若無則紅
            bool dollerSign = false;

            ArrayList a_KeepData = new ArrayList();
            a_KeepData.Clear();

            //var cellStyle = workbook.CreateCellStyle();
            ICellStyle cellStyle;
            if (loopDic.ContainsKey(wsTemplate.SheetName) && loopDic[wsTemplate.SheetName].ContainsKey(j))
            {
                cellStyle = loopDic[wsTemplate.SheetName][j];
            }
            else
            {
                cellStyle = workbook.CreateCellStyle();
                cellStyle.CloneStyleFrom(wsTemplate.GetRow(i).GetCell(j).CellStyle);

                if (!loopDic.ContainsKey(wsTemplate.SheetName))
                {
                    loopDic.Add(wsTemplate.SheetName, new Dictionary<int, ICellStyle> { { j, cellStyle } });
                }
                else
                {
                    loopDic[wsTemplate.SheetName].Add(j, cellStyle);
                }
            }
            
            var dataFormat = workbook.CreateDataFormat();
            //cellStyle = wsTemplate.GetRow(i).GetCell(j).CellStyle;
            //cellStyle.CloneStyleFrom(wsTemplate.GetRow(i).GetCell(j).CellStyle);
            //var dataFormat2 = workbook.CreateCellStyle();
            short sourceRowHeight = wsTemplate.GetRow(i).Height;

            int i_startIndex = wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("<");
            int i_endIndex = wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(">");
            string columnName = wsTemplate.GetRow(i).GetCell(j).ToString().Substring(i_startIndex + 1, i_endIndex - i_startIndex - 1);


            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N1") != -1)  // Template ":N1" = 小數一位
            {
                decimalPlace = 1;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N2") != -1)  // Template ":N2" = 小數二位
            {
                decimalPlace = 2;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N3") != -1)  // Template ":N3" = 小數三位
            {
                decimalPlace = 3;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N4") != -1)  // Template ":N4" = 小數四位
            {
                decimalPlace = 4;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N5") != -1)  // Template ":54" = 小數2位加上doller sign
            {
                decimalPlace = 5;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N6") != -1)  // Template ":N6" = 小數六位
            {
                decimalPlace = 6;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N7") != -1)  // Template ":N7" = 小數七位
            {
                decimalPlace = 7;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":P2") != -1)  // Template ":P2" = 百分比2位
            {
                decimalPlace = -1;

                percent = true;
                chkNumeric = true;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":Per3&N1") != -1)  // Template ":Per3&N1" = 百分比3位，且不進行四捨五入
            {

                /*
                四捨五入最小至小數點後第15位。
                需給i_DecimalPlace值，才不會進入1548行預設的四捨五入至整數位，
                但又不能給定目前有用到的值，不然會掉入1421行之後眾多i_DecimalPlace的格式設定，而不會進入1452行的我們要的設定。
                所以i_DecimalPlace設定成15，才不會掉入1421行之後的設定，同時也不會進入1548行預設的四捨五入至整數位。
                */
                decimalPlace = 15; 

                percent = true;
                chkNumeric = true;
            }

            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":SPACE") != -1)  // Template "::Space" = 數值Null顯示空白
            {
                //b_space = true;
                nullDisplayValue = "";
                if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":S1") != -1)  // Template "::S1" = 背景黑色
                {
                    style = 1;
                }
            }
            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":NULL1") != -1)  // Template ":NULL1" = 數值Null顯示"-"
            {
                nullDisplayValue = "-";
            }
            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":SUM") != -1)    // Template ":SUM" = 需要加總
            {
                chkSum = true;
                chkNumeric = true;
                chkInteger = true;
            }

            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":D") != -1)    // Template ":D" = 日期格式
            {
                chkDate = true;
            }
            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":I") != -1)    // Template ":I" = 整數
            {
                chkInteger = true;
            }
            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":IB") != -1)    // Template ":IB" = 整數, 若負數顏色依然黑
            {
                chkInteger = true;
                color = "BLACK";
            }
            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":$I") != -1)    // Template ":$I" = 整數 加上 doller sign($) 
            {
                chkInteger = true;
                dollerSign = true;
            }
     
            #endregion
            //if (style == 1)
            //{
            //    dataFormat2.CloneStyleFrom(cellStyle);
            //    dataFormat2.Alignment = HorizontalAlignment.Right;
            //    dataFormat2.FillForegroundColor = dataFormat2 is HSSFCellStyle ? HSSFColor.Black.Index : IndexedColors.Black.Index;
            //    dataFormat2.FillPattern = FillPattern.SolidForeground;
            //}

            if (chkNumeric || chkInteger)   //是否為數值
            {
                if (chkNumeric && decimalPlace == 1)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0;-#,##0.0");
                }
                else if (chkNumeric && decimalPlace == 2)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.00;-#,##0.00");
                }
                else if (chkNumeric && decimalPlace == 3)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000;-#,##0.000");
                }
                else if (chkNumeric && decimalPlace == 4)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0000;-#,##0.0000");
                }
                else if (chkNumeric && decimalPlace == 5)
                    cellStyle.DataFormat = dataFormat.GetFormat("$#,##0.00;-$#,##0.00");
                else if (chkNumeric && decimalPlace == 6)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000000;-#,##0.000000");
                }
                else if (chkNumeric && decimalPlace == 7)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0000000;-#,##0.0000000");
                }
                else if (chkNumeric && percent)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat($"#0.00%;-#0.00%");
                }
                else if (chkInteger && color == "BLACK")
                {
                    cellStyle.DataFormat = dataFormat.GetFormat("#,###0;-#,###0");
                }
                else if (chkInteger && dollerSign)
                {
                    cellStyle.DataFormat = dataFormat.GetFormat("$#,###0;-$#,###0");
                }
                else
                {
                    cellStyle.DataFormat = dataFormat.GetFormat("$#,###0;-$#,###0");
                    //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,###0;-#,###0");
                }
            }
            for (int k = 0; k < datas.Columns.Count; k++)
            {
                if (columnName.ToUpper() == datas.Columns[k].ColumnName.ToUpper())
                {
                    i_DataCount = datas.Rows.Count;
                    for (int m = i_aTempMAXROW; m > i; m--)
                    {
                        if (datas.Rows.Count > 1)  
                        {
                            var newRow = wsTemplate.GetRow(datas.Rows.Count + m - 1);   //目的
                            var sourceRow = wsTemplate.GetRow(m);   //來源
                            ICell sourceCell, newCell;

                            if (sourceRow == null)
                            { continue; }

                            if (newRow == null)
                            {
                                newRow = wsTemplate.CreateRow(datas.Rows.Count + m - 1);
                                if(newRow != null)
                                {
                                    newCell = newRow.CreateCell(j);

                                    newRow.Height = sourceRowHeight; // 建立新列時，帶入列高
                                }
                                else
                                {
                                    newCell = null;
                                }
                            }
                            else
                            {
                                newCell = newRow.GetCell(j);
                                if (newCell == null)
                                { newCell = newRow.CreateCell(j); }
                                else
                                { newCell = newRow.GetCell(j); }
                            }

                            sourceCell = sourceRow.GetCell(j);

                            if (sourceCell == null)
                            {
                                newCell = null;
                                continue;
                            }


                            if (sourceCell != null)
                            {
                                CopyCell(sourceCell, newCell);
                                sourceRow.CreateCell(j);
                            }
                        }
                    }

                    for (int m = 0; m < datas.Rows.Count; m++)
                    {
                        if (wsTemplate.GetRow(i + m) == null)
                        {
                            IRow row = wsTemplate.CreateRow(i + m);
                            row.Height = sourceRowHeight; // 建立新列時，帶入列高
                        }

                        if (chkNumeric || chkInteger)   //是否為數值
                        {
                            if (nullDisplayValue != null && (datas.Rows[m][k] == DBNull.Value)) //設定Null的帶入值
                            {
                                wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(nullDisplayValue);
                                //if (style == 1)
                                //{
                                //    wsTemplate.GetRow(i + m).GetCell(j).CellStyle = dataFormat2;
                                //}
                                //else
                                //{
                                    cellStyle.Alignment = HorizontalAlignment.Right;
                                    wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;
                                    //wsTemplate.GetRow(i + m).GetCell(j).CellStyle.Alignment = HorizontalAlignment.Right;
                                //}

                                wsTemplate.GetRow(i + m).GetCell(j).SetCellType(CellType.String);
                            }
                            else
                            {
                                double dTmp;

                                string strValue = datas.Rows[m][k] == DBNull.Value ? "0" : datas.Rows[m][k].ToString();
                                if (double.TryParse(strValue, out dTmp))
                                {
                                    // 數值四捨五入
                                    if (chkNumeric && decimalPlace >= 0)
                                        dTmp = RoundX(dTmp, decimalPlace);
                                    else if (chkInteger)
                                        dTmp = RoundX(dTmp, 0);

                                    if (percent)
                                    {
                                        wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(dTmp * 0.01);
                                    }
                                    else
                                    {
                                        wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(dTmp);
                                    }
                                }
                                else
                                {
                                    wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(0);
                                }
                                wsTemplate.GetRow(i + m).GetCell(j).SetCellType(CellType.Numeric);
                                cellStyle.Alignment = HorizontalAlignment.Right;
                                wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;

                                //wsTemplate.GetRow(i + m).GetCell(j).CellStyle.Alignment = HorizontalAlignment.Right;
                            }

                            if (chkSum)    //是否需要加總
                            {
                                if(wsTemplate.GetRow(i + m).GetCell(j).CellType == CellType.Numeric)
                                    d_Sum = d_Sum + wsTemplate.GetRow(i + m).GetCell(j).NumericCellValue;

                                if (wsTemplate.GetRow(i + m + 1) == null)
                                {
                                    IRow row = wsTemplate.CreateRow(i + m + 1);
                                    row.Height = sourceRowHeight; // 建立新列時，帶入列高
                                }
                                if (wsTemplate.GetRow(i + m + 1).GetCell(j) == null)
                                {
                                    wsTemplate.GetRow(i + m + 1).CreateCell(j);
                                }

                                wsTemplate.GetRow(i + m + 1).GetCell(j).SetCellValue(d_Sum);
                                wsTemplate.GetRow(i + m + 1).GetCell(j).SetCellType(CellType.Numeric);
                                
                                cellStyle.Alignment = HorizontalAlignment.Right;
                                wsTemplate.GetRow(i + m + 1).GetCell(j).CellStyle = cellStyle;
                                wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;
                                //wsTemplate.GetRow(i + m + 1).GetCell(j).CellStyle.DataFormat = cellStyle.DataFormat;
                                //wsTemplate.GetRow(i + m + 1).GetCell(j).CellStyle = o_CellStyle;
                            }
                        }
                        else if (chkDate)    //是否為日期
                        {
                            DateTime tTmp;
                            string s_Date = datas.Rows[m][k] == DBNull.Value ? "" : datas.Rows[m][k].ToString();
                            if (DateTime.TryParse(s_Date, out tTmp))
                                wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(tTmp.ToString("yyyy/MM/dd"));
                            else
                                wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue("");
                            wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;
                        }
                        else
                        {

                            string cellvalue = (datas.Rows[m][k] == DBNull.Value) ? "" : datas.Rows[m][k].ToString();
                            wsTemplate.GetRow(i + m).CreateCell(j).SetCellValue(cellvalue);  //寫入DATA
                            wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;

                        }
                    }
                }
            }

            return i_DataCount;
        }
       
        private bool CheckInRange(int maxValue, int minValue, int v)
        {
            return (maxValue >= v && v >= minValue);
        }

        /// <summary>
        /// 取得儲存格的合併設定
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private CellRangeAddress GetMergedSetting(ICell cell)
        {
            if (cell == null)
                return null;

            ISheet sheet = cell.Row.Sheet;
            int numOfSettings = sheet.NumMergedRegions;
            CellRangeAddress mergedSetttingData = null;

            for (int i = 0; i < numOfSettings; i++)
            {
                mergedSetttingData = sheet.GetMergedRegion(i);

                if (mergedSetttingData == null)
                    continue;

                if (mergedSetttingData.FirstRow == (cell.Row.RowNum - 1)
                    && mergedSetttingData.FirstColumn == cell.ColumnIndex)
                {
                    break;
                }
                else
                {
                    mergedSetttingData = null;
                }
            }

            return mergedSetttingData;
        }

        /// <summary>
        /// 移除合併儲存格設定
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="mergedSetting"></param>
        private void RemoveMergedCell(ISheet sheet, CellRangeAddress mergedSetting)
        {
            int numOfSettings = sheet.NumMergedRegions;

            for (int i = 0; i < numOfSettings; i++)
            {
                CellRangeAddress mergedSetttingData = sheet.GetMergedRegion(i);
                if (mergedSetttingData == null)
                    continue;

                if (mergedSetttingData.FirstRow == mergedSetting.FirstRow &&
                    mergedSetttingData.LastRow == mergedSetting.LastRow &&
                    mergedSetttingData.FirstColumn == mergedSetting.FirstColumn &&
                    mergedSetttingData.LastColumn == mergedSetting.LastColumn)
                {
                    sheet.RemoveMergedRegion(i);
                    break;
                }
            }
        }

        /// <summary>
        /// 將指定範圍儲存格合併為一個
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellStart"></param>
        /// <param name="cellEnd"></param>
        private void AddMergedCellWithCreate(ISheet sheet, int row, int cellStart, int cellEnd)
        {
            if (sheet == null)
                return;

            if (sheet.GetRow(row) == null)
            {
                sheet.CreateRow(row);
            }

            if (sheet.GetRow(row).GetCell(cellStart) == null)
            {
                sheet.GetRow(row).CreateCell(cellStart);
            }

            if (sheet.GetRow(row).GetCell(cellEnd) == null)
            {
                sheet.GetRow(row).CreateCell(cellEnd);
            }
           
            CellRangeAddress newCellRangeAddress = new CellRangeAddress(row, row, cellStart, cellEnd);
            sheet.AddMergedRegion(newCellRangeAddress);
        }

        /// <summary>
        /// 將指定範圍儲存格合併為一個
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellStart"></param>
        /// <param name="cellEnd"></param>
        private void AddMergedCell(ISheet sheet, ICell cellStart, ICell cellEnd)
        {
            if (sheet == null)
                return;

            if (cellStart == null)
                return;

            if (cellEnd == null)
                return;

            CellRangeAddress newCellRangeAddress = new CellRangeAddress(cellStart.Row.RowNum,
                                                                                 cellEnd.Row.RowNum,
                                                                                 cellStart.ColumnIndex,
                                                                                 cellEnd.ColumnIndex);
            sheet.AddMergedRegion(newCellRangeAddress);
        }

        private bool isMergedRow(ISheet sheet, int row, int column, out int FirstRow, out int LastRow)
        {
            int sheetMergeCount = sheet.NumMergedRegions;
            for (int i = 0; i < sheetMergeCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);

                if (range == null) // 發現會有null，而導致產檔之合併儲存格失效，故加入排除null的判斷
                    continue;

                int firstColumn = range.FirstColumn;
                int lastColumn = range.LastColumn;
                int firstRow = range.FirstRow;
                int lastRow = range.LastRow;

                if (row >= firstRow && row <= lastRow)
                {
                    if (column >= firstColumn && column <= lastColumn)
                    {
                        FirstRow = firstRow;
                        LastRow = lastRow;
                        return true;
                    }
                }
            }
            FirstRow = -1;
            LastRow = -1;

            return false;
        }

        /// <summary>
        /// 傳回指定儲存格是否為合併儲存格?
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="FirstColumn"></param>
        /// <param name="LastColumn"></param>
        /// <returns></returns>
        private bool IsMergedCell(ISheet sheet, int row, int column, out int FirstColumn, out int LastColumn)
        {
            int sheetMergeCount = sheet.NumMergedRegions;
            for (int i = 0; i < sheetMergeCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                if (range == null)
                    continue;

                int firstColumn = range.FirstColumn;
                int lastColumn = range.LastColumn;
                int firstRow = range.FirstRow;
                int lastRow = range.LastRow;

                if (row == firstRow && row == lastRow)
                {
                    if (column >= firstColumn && column <= lastColumn)
                    {
                        FirstColumn = firstColumn;
                        LastColumn = lastColumn;
                        return true;
                    }
                }
            }

            FirstColumn = -1;
            LastColumn = -1;

            return false;
        }

        /// <summary>
        /// 傳回工作表是否含有合併儲存格?
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private bool IsHasMergedRegions(ISheet sheet)
        {
            return (sheet?.NumMergedRegions ?? 0) > 0;
        }

        private void SetBreak(ISheet wsTemplate, int i_PageRowMax, int i_TitleBegin, int i_TitleRowsCount, bool bSpace = false, int iRealPageMaxRow = 42, int i_TitleRow = 0)
        {
            var source = wsTemplate;
            for (int iRow = i_PageRowMax + i_TitleRow; iRow < wsTemplate.LastRowNum; iRow += i_PageRowMax)
            {
                wsTemplate.ShiftRows(iRow + i_TitleRow, wsTemplate.LastRowNum, i_TitleRowsCount);
                for (int iTitleRow = i_TitleBegin - 1; iTitleRow < i_TitleRowsCount; iTitleRow++)
                {
                    CopyRow(ref source, i_TitleBegin - 1 + iTitleRow, ref source, iRow + iTitleRow);
                }
            }
        }

        private void ScanTitleToSetRowBreak(SheetInfo[] sInfo, bool bSpace = false)
        {
            int iRealPageRowMax = 46;
            string sIdnTitle = "";
            int iPageCount = 1;
            for (int i = 0; i < sInfo.Length; i++)
            {
                //if (sInfo[i].TargetSheetName == null) continue;
                //if (sInfo[i].Title == null) continue;
                ISheet ISourceSheet;
                if (sInfo[i].TargetSheetName != null)
                {
                    ISourceSheet = workbook.GetSheet(sInfo[i].TargetSheetName);
                }
                else
                {
                    ISourceSheet = workbook.GetSheet(sInfo[i].SheetName);
                }
                if (ISourceSheet == null) continue;
                for (int iRow = ISourceSheet.FirstRowNum + 1; iRow < ISourceSheet.LastRowNum; iRow++)
                {
                    IRow ir = ISourceSheet.GetRow(iRow);
                    if (ir == null) continue;
                    for (int iCol = ISourceSheet.GetRow(iRow).FirstCellNum; iCol < ISourceSheet.GetRow(iRow).LastCellNum; iCol++)
                    {
                        if (ISourceSheet.GetRow(iRow).GetCell(iCol) == null)
                        {
                            continue;
                        }
                        if (sInfo[i].Title != null && ISourceSheet.GetRow(iRow).GetCell(iCol).ToString().IndexOf(sInfo[i].Title) >= 0)
                        {
                            if (bSpace)
                            {
                                for (int iR = 0; iR < sInfo[i].PageRowMax - (iRow % sInfo[i].PageRowMax); iR++)
                                {
                                    ISourceSheet.ShiftRows(iRow, ISourceSheet.LastRowNum, 1);
                                }
                                iRow += sInfo[i].PageRowMax - (iRow % sInfo[i].PageRowMax);
                            }
                            else
                            {
                                ISourceSheet.SetRowBreak(iRow - 1);
                            }
                        }
                        if (sInfo[i].PageCountTitle == null) continue;
                        if (sInfo[i].IdnTitle == null) continue;
                        if (ISourceSheet.GetRow(iRow).GetCell(iCol).ToString().IndexOf(sInfo[i].IdnTitle) >= 0)
                        {
                            if (sIdnTitle == ISourceSheet.GetRow(iRow).GetCell(iCol + 1).ToString())
                            {
                                iPageCount += 1;
                            }
                            else
                            {
                                sIdnTitle = ISourceSheet.GetRow(iRow).GetCell(iCol + 1).ToString();
                                iPageCount = 1;
                            }
                        }
                        if (ISourceSheet.GetRow(iRow).GetCell(iCol).ToString().IndexOf(sInfo[i].PageCountTitle) >= 0)
                        {
                            ISourceSheet.GetRow(iRow).GetCell(iCol + 1).SetCellValue(iPageCount + ISourceSheet.GetRow(iRow).GetCell(iCol + 1).ToString());
                        }
                    }
                }
                if (sInfo[i].TargetSheetName != null)
                {
                    int index = workbook.GetSheetIndex(sInfo[i].SheetName);
                    workbook.RemoveSheetAt(index);
                }
            }
        }
        /// <summary>
        /// SetRowSumFormulla
        /// 填入橫向加總表內容
        /// </summary>
        /// <param name="wsTemplate">工作表</param>
        /// <param name="a_aSumDatas">從第幾列開始插入列</param>
        /// <param name="s_aProgramId">共需插入幾列</param>
        /// <param name="i">跑到第幾ROW</param>
        /// <param name="j">跑到第幾COLUMN</param>
        /// <returns></returns>
        private void SetRowSumFormulla(System.Data.DataTable o_aDatas, ISheet wsTemplate, ArrayList a_aSumDatas, string s_aProgramId, int i, int j)
        {
            string s_DelimStr = " =+";
            char[] a_Delimiter = s_DelimStr.ToCharArray();
            string s_Tag = wsTemplate.GetRow(i).GetCell(j).ToString().Replace("##SUM=", "");
            string[] a_Split = null;


            a_Split = s_Tag.Split(a_Delimiter);

            for (int m = 0; m < o_aDatas.Rows.Count; m++)
            {
                string s_Data = null;
                for (int k = 0; k < a_Split.Length; k++)
                {
                    s_Data = s_Data + a_Split[k] + (i + m + 1) + "+";
                }

                if(s_Data != null)
                {
                    s_Data = s_Data.TrimEnd('+'); 

                }

                var cellStyle = workbook.CreateCellStyle();
                var dataFormat = workbook.CreateDataFormat();

                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.00_);(#,##0.00)");

                wsTemplate.GetRow(i + m).CreateCell(j).SetCellFormula(s_Data);
                wsTemplate.GetRow(i + m).GetCell(j).CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// DeleteOtherSheet
        /// 刪除除了指定Sheet陣列以外的所有Sheet
        /// </summary>
        /// <param name="strSheetName">Sheet名稱(陣列)</param>
        /// <param name="strErrMsg">傳出的錯誤訊息</param>
        /// <returns>是否成功刪除（如果m_wbOutBook=null,也算成功）,true成功/false失敗</returns>
        public bool DeleteOtherSheet(ref IWorkbook workbook, ArrayList strSheetNames)
        {
            if (workbook != null)
            {
                //工作表個數(最多遍巡幾個工作表)
                int SheetCount = workbook.NumberOfSheets;
                //從第0個工作表開始確認
                int index = 0;
                //每個工作表名稱確認是否與指定名稱相同，不同者則刪除工作表
                for (int i = 0; i < SheetCount; i++)
                {
                    if (!strSheetNames.Contains(workbook.GetSheetName(index)))
                    {
                        if (workbook.GetSheetAt(index).RepeatingRows == null && workbook.GetSheetAt(index).RepeatingColumns == null)
                        {
                            workbook.RemoveSheetAt(index);
                        }
                        else
                        {
                            workbook.SetSheetHidden(index, SheetState.Hidden);
                        }
                    }
                    else
                    {
                        index++;
                    }
                }
            }
            return true;
        }

        public bool DeleteSheet(ref IWorkbook workbook, ArrayList strSheetNames, ref string strErrMsg)
        {
            try
            {
                if (workbook != null)
                {
                    //工作表個數(最多遍巡幾個工作表)
                    int SheetCount = workbook.NumberOfSheets;
                    //從第0個工作表開始確認
                    int index = 0;
                    //每個工作表名稱確認是否與指定名稱相同，不同者則刪除工作表
                    for (int i = 0; i < SheetCount; i++)
                    {
                        if (strSheetNames.Contains(workbook.GetSheetName(index)))
                        {
                            workbook.RemoveSheetAt(index);
                        }
                        else
                        {
                            index++;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                strErrMsg = ex.ToString();
                return false;
            }
        }

        /// <summary>
        /// DeleteTags
        /// 刪除Tags
        /// </summary> 
        public void DeleteTags(ref ISheet wsTemplate)
        {
            try
            {
                int i_aLastRow = wsTemplate.LastRowNum;
                for (int i = 0; i <= i_aLastRow; i++)
                {
                    if (wsTemplate.GetRow(i) == null)
                    { continue; }
                    else
                    {
                        int i_Cols = wsTemplate.GetRow(i).LastCellNum;
                        for (int j = 0; j < i_Cols; j++)
                        {
                            if (wsTemplate.GetRow(i).GetCell(j) == null)
                            { continue; }
                            else
                            {
                                if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##LOOPFIELD_") != -1 || wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##FIELD_") != -1 || wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf("##DYNAMICFIELD") != -1)
                                {
                                    wsTemplate.GetRow(i).GetCell(j).SetCellValue("");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                int i = 0;
            }
        }

        /// <summary>
        /// Copy Row
        /// 複製單列
        /// </summary>  
        /// <param name="sourceWorksheet">來源的Sheet (非WorkBook)</param>
        /// <param name="sourceRowNum">欲複製的起始列</param>
        /// <param name="destinationWorksheet">目標的Sheet</param>
        /// <param name="destinationRowNum">目標的起始列</param>
        /// <param name="IsRemoveSrcRow">是否清掉原列</param>
        /// <param name="copyRowHeight">複製行高到新列</param>
        /// <param name="resetOriginalRowHeight">重製原始列行高</param>        
        /// <returns>void</returns>
        public void CopyRow(ref ISheet sourceWorksheet, int sourceRowNum,
                             ref ISheet destinationWorksheet, int destinationRowNum,
                             bool IsRemoveSrcRow = false, bool copyRowHeight = true, bool resetOriginalRowHeight = true)
        {
            var newRow = destinationWorksheet.CreateRow(destinationRowNum);   //目的先Create
            var sourceRow = sourceWorksheet.GetRow(sourceRowNum);   //來原先Get
            ICell oldCell, newCell;
            int i;

            if (sourceRow == null)
            {
                return;
            }

            // Loop through source columns to add to new row
            for (i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                oldCell = sourceRow.GetCell(i);
                newCell = newRow.GetCell(i); 

                if (newCell == null)
                    newCell = newRow.CreateCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }

                // Copy style from old cell and apply to new cell
                if (newCell != null && oldCell != null)
                {
                    newCell.CellStyle = oldCell.CellStyle; 
                }
                // If there is a cell comment, copy
                if (newCell != null) newCell.CellComment = oldCell.CellComment;

                // If there is a cell hyperlink, copy
                if (newCell != null) newCell.Hyperlink = oldCell.Hyperlink;

                if (newCell != null)
                { 
                    // Set the cell data value
                    switch (oldCell.CellType)
                    {
                        case CellType.Blank:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            newCell.SetCellValue(oldCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                            break;
                        case CellType.Formula:
                            newCell.CellFormula = oldCell.CellFormula;
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue(oldCell.NumericCellValue);
                            break;
                        case CellType.String:
                            newCell.SetCellValue(oldCell.RichStringCellValue);
                            break;
                        case CellType.Unknown:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                    }
                }
                   
            }

            #region 合併儲存格
            // 將之前的合併儲存格的處理加回來
            // 跨欄的問題
            CellRangeAddress cellRangeAddress = null, newCellRangeAddress = null;
            for (i = 0; i < sourceWorksheet.NumMergedRegions; i++)
            {
                cellRangeAddress = sourceWorksheet.GetMergedRegion(i);
                if (cellRangeAddress == null)
                    continue;

                try
                {
                    if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                    {
                        if(newRow != null && cellRangeAddress != null)
                        {
                            newCellRangeAddress = new CellRangeAddress(newRow.RowNum, (newRow.RowNum + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)), cellRangeAddress.FirstColumn, cellRangeAddress.LastColumn);
                            //增加判斷是否已處理過[合併儲存格](如:WEB UI程式中已處理)。若重複處理[合併儲存格]，產出的檔案會出現[合併儲存格]錯誤
                            int FirstRow, LastRow;
                            if (!isMergedRow(destinationWorksheet, newRow.RowNum, cellRangeAddress.FirstColumn, out FirstRow, out LastRow))
                            {
                                destinationWorksheet.AddMergedRegion(newCellRangeAddress);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //nothing
                }
            }
            #endregion

            //複製行高到新列
            if (copyRowHeight)
                if(newRow != null && sourceRow != null)
                {
                    newRow.Height = sourceRow.Height; 
                }
            //重製原始列行高
            if (resetOriginalRowHeight)
            {
                if (sourceRow != null)
                {
                    sourceRow.Height = sourceWorksheet.DefaultRowHeight;
                }
            }
            //清掉原列
            if (IsRemoveSrcRow == true)
                sourceWorksheet.RemoveRow(sourceRow);
        }       

        /// <summary>
        /// HSSFRow Copy Command
        /// 
        /// Description:  Inserts a existing row into a new row, will automatically push down
        ///               any existing rows.  Copy is done cell by cell and supports, and the
        ///               command tries to copy all properties available (style, merged cells, values, etc...)
        /// </summary>
        /// <param name="workbook">Workbook containing the worksheet that will be changed</param>
        /// <param name="worksheet">WorkSheet containing rows to be copied</param>
        /// <param name="sourceRowNum">Source Row Number</param>
        /// <param name="destinationRowNum">Destination Row Number</param>
        public void CopyRow2(IWorkbook workbook, ISheet worksheet, int sourceRowNum, int destinationRowNum, bool IsRemoveSrcRow = false, bool copyRowHeight = true, bool resetOriginalRowHeight = true)
        {
            // Get the source / new row
            IRow newRow = worksheet.GetRow(destinationRowNum);
            IRow sourceRow = worksheet.GetRow(sourceRowNum);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1);
            }
            else
            {
                newRow = worksheet.CreateRow(destinationRowNum);
            }

            // Loop through source columns to add to new row
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }

                // Copy style from old cell and apply to new cell
                ICellStyle newCellStyle = workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
                newCell.CellStyle = newCellStyle;

                // If there is a cell comment, copy
                if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
                {
                    case CellType.Blank:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < worksheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);

                if (cellRangeAddress == null)
                    continue;
                try
                {
                    if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                    {
                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                                                    (newRow.RowNum +
                                                                                     (cellRangeAddress.FirstRow -
                                                                                      cellRangeAddress.LastRow)),
                                                                                    cellRangeAddress.FirstColumn,
                                                                                    cellRangeAddress.LastColumn);
                        worksheet.AddMergedRegion(newCellRangeAddress);
                    }
                }
                catch (Exception)
                {


                }

            }

            //複製行高到新列
            if (copyRowHeight)
                newRow.Height = sourceRow.Height;
            //重製原始列行高
            if (resetOriginalRowHeight)
                sourceRow.Height = worksheet.DefaultRowHeight;
            //清掉原列
            if (IsRemoveSrcRow == true)
                worksheet.RemoveRow(sourceRow);
        }
        public void CopySheet(ref ISheet sourceWorksheet,
                             ref ISheet destinationWorksheet, int destinationRowNum,
                             bool IsRemoveSrcRow = false, bool copyRowHeight = true, bool resetOriginalRowHeight = true)
        {

            for (int sourceRowNum = 0; sourceRowNum <= sourceWorksheet.LastRowNum; sourceRowNum++)
            {
                CopyRow(ref sourceWorksheet, sourceRowNum, ref destinationWorksheet, destinationRowNum,
                        IsRemoveSrcRow, copyRowHeight, resetOriginalRowHeight);

                destinationRowNum++;
            }
        }

        void SetCellValue(ICell cell, object value, bool b_chkNumeric = false, bool b_chkInteger = false, bool b_chkSum = false,
            bool b_chkDate = false,
            Action<double, int> sumRowProcess = null,
            double d_Sum = 0.0,
            int i_DecimalPlace = 0,
            ICellStyle o_CellStyle = null, IDataFormat o_DataFormat = null, bool b_isCustomFormat = false, int i_CustomType = -1)
        {
            if (cell == null)
                return;

            o_CellStyle = o_CellStyle ?? cell.Row.Sheet.Workbook.CreateCellStyle();
            o_DataFormat = o_DataFormat ?? cell.Row.Sheet.Workbook.CreateDataFormat();

            if (b_chkNumeric || b_chkInteger)   //是否為數值
            {

                if (b_chkNumeric && i_DecimalPlace == 1)
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0_);(#,##0.0)");
                }
                else if (b_chkNumeric && i_DecimalPlace == 2)
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.00_);(#,##0.00)");
                }
                else if (b_chkNumeric && i_DecimalPlace == 3)
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.000_);(#,##0.000)");
                }
                else if (b_chkNumeric && i_DecimalPlace == 4)
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0000_);(#,##0.0000)");
                }
                else if (b_chkNumeric && i_DecimalPlace == 5)
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat("$#,##0.00;-$#,##0.00");
                }
                else
                {
                    o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,###0;-#,###0");
                }

                cell.SetCellType(CellType.Numeric);
                cell.CellStyle = o_CellStyle;

                if (b_chkSum)    //是否需要加總
                {
                    d_Sum = d_Sum + cell.NumericCellValue;

                    if (sumRowProcess != null)
                    {
                        sumRowProcess(d_Sum, cell.RowIndex);
                    }
                }
            }
            string _sFormat = "";
            if (b_isCustomFormat)
            {
                switch (i_CustomType)
                {
                    case 1:
                        //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0 \"%\"_);(#,##0.0 \"%\")");
                        _sFormat = "###0.#";
                        cell.SetCellType(CellType.String);
                        cell.CellStyle = o_CellStyle;
                        break;
                    default:
                    case 2:
                        //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0# \"%\"_);(#,##0.0# \"%\")");
                        _sFormat = "###0.##";
                        cell.SetCellType(CellType.String);
                        cell.CellStyle = o_CellStyle;
                        break;
                    case 3:
                        //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0## \"%\"_);(#,##0.0## \"%\")");
                        _sFormat = "###0.###";
                        cell.SetCellType(CellType.String);
                        cell.CellStyle = o_CellStyle;
                        break;
                    case 4:
                        //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0### \"%\"_);(#,##0.0### \"%\")");
                        _sFormat = "###0.####";
                        cell.SetCellType(CellType.String);
                        cell.CellStyle = o_CellStyle;
                        break;
                    case 5:
                        //o_CellStyle.DataFormat = o_DataFormat.GetFormat($"#,##0.0#### \"%\"_);(#,##0.0#### \"%\")");
                        _sFormat = "###0.#####";
                        cell.SetCellType(CellType.String);
                        cell.CellStyle = o_CellStyle;
                        break;
                }
            }

            //寫入DATA
            switch (cell.CellType)
            {
                case CellType.Blank:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
                case CellType.Boolean:
                    cell.SetCellValue(Convert.ToBoolean(value));
                    break;
                case CellType.Error:
                    cell.SetCellErrorValue(Convert.ToByte(value));
                    break;
                case CellType.Formula:
                    cell.CellFormula = Convert.ToString(value);
                    break;
                case CellType.Numeric:
                    double dTmp;

                    string s_Value = value == DBNull.Value ? "0" : value.ToString();

                    if (double.TryParse(s_Value, out dTmp))
                        cell.SetCellValue(dTmp);
                    else
                        cell.SetCellValue(0);

                    break;
                case CellType.String:

                    if (b_chkDate)
                    {
                        DateTime tTmp;
                        string s_Date = value == DBNull.Value ? "" : value.ToString();

                        if (DateTime.TryParse(s_Date, out tTmp))
                            cell.SetCellValue(tTmp.ToString("yyyy/MM/dd"));
                        else
                            cell.SetCellValue("");
                    }
                    else
                    {
                        if (b_isCustomFormat)
                        {

                            string cellvalue = (value == DBNull.Value) ? "" : string.Format("{0:" + _sFormat + "}", value);
                            cellvalue = string.Format("{0} %", cellvalue.TrimEnd('.'));
                            cell.SetCellValue(cellvalue);
                        }
                        else
                        {
                            string cellvalue = (value == DBNull.Value) ? "" : value.ToString();
                            cell.SetCellValue(cellvalue);
                        }

                    }

                    break;
                case CellType.Unknown:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
            }
        }

        private void SetTagData(ISheet wsTemplate, int i, int j, string value)
        {
            //ICellStyle o_CellStyle = wsTemplate.GetRow(i).GetCell(j).CellStyle;
            //XSSFCellStyle o_CellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            var dataFormat = workbook.CreateDataFormat();
            //var cellStyle = workbook.CreateCellStyle();
            ICellStyle cellStyle = wsTemplate.GetRow(i).GetCell(j).CellStyle;
            //cellStyle.CloneStyleFrom(wsTemplate.GetRow(i).GetCell(j).CellStyle);

            if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N1") != -1)  // Template ":N1" = 小數一位
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0;-#,##0.0");
                double dValue = RoundX(Convert.ToDouble(value), 1); // 數值四捨五入
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric);
                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0;-#,##0.0");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N2") != -1)  // Template ":N2" = 小數二位
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,##0.00;-#,##0.00");
                double dValue = RoundX(Convert.ToDouble(value), 2); // 數值四捨五入
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric);
                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.00;-#,##0.00");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N3") != -1)  // Template ":N3" = 小數三位
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000;-#,##0.000");
                double dValue = RoundX(Convert.ToDouble(value), 3); // 數值四捨五入
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric);
                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000;-#,##0.000");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N4") != -1)  // Template ":N4" = 小數四位
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0000;-#,##0.0000");
                double dValue = RoundX(Convert.ToDouble(value), 4); // 數值四捨五入
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric);
                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.0000;-#,##0.0000");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":N6") != -1)  // Template ":N6" = 小數六位
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000000;-#,##0.000000");
                double dValue = RoundX(Convert.ToDouble(value), 6); // 數值四捨五入
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric);
                cellStyle.DataFormat = dataFormat.GetFormat($"#,##0.000000;-#,##0.000000");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":IB") != -1)    // Template ":IB" = 整數, 負數顏色為黑
            {
                //o_CellStyle.DataFormat = dataFormat.GetFormat("#,###0;-#,###0");
                double dValue = RoundX(Convert.ToDouble(value), 0);
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric); // 數值四捨五入
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.DataFormat = dataFormat.GetFormat("#,###0;-#,###0");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else if (wsTemplate.GetRow(i).GetCell(j).ToString().IndexOf(":I") != -1)    // Template ":I" = 整數
            {
                double dValue = RoundX(Convert.ToDouble(value), 0);
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(dValue);
                //o_CellStyle.DataFormat = dataFormat.GetFormat($"#,###0;-#,###0");
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.Numeric); // 數值四捨五入
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.DataFormat = dataFormat.GetFormat($"#,###0;-#,###0");
                wsTemplate.GetRow(i).GetCell(j).CellStyle = cellStyle;
            }
            else
            {
                wsTemplate.GetRow(i).GetCell(j).SetCellValue(value);
                wsTemplate.GetRow(i).GetCell(j).SetCellType(CellType.String);
            }
        }

        private void CopyCell(ICell sourceCell, ICell newCell, string sheetName = null, IWorkbook workbook = null)
        {
            // Copy style from old cell and apply to new cell
            var newCellKey = Convert.ToString(newCell.ColumnIndex) + Convert.ToString(newCell.RowIndex);
            if (workbook == null)
            {
                newCell.CellStyle = sourceCell.CellStyle;
            }
            else if (dynamicDic.ContainsKey(sheetName) && dynamicDic[sheetName].ContainsKey(newCellKey))
            {
                newCell.CellStyle = dynamicDic[sheetName][newCellKey];
            }
            else
            {
                ICellStyle newCellStyle = workbook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(sourceCell.CellStyle);
                newCell.CellStyle = newCellStyle;

                if (!dynamicDic.ContainsKey(sheetName))
                {
                    dynamicDic.Add(sheetName, new Dictionary<string, ICellStyle> { { newCellKey, newCellStyle } });
                }
                else
                {
                    dynamicDic[sheetName].Add(newCellKey, newCellStyle);
                }
            }

            //if (workbook == null || (sourceCell.))
            //{
            //    newCell.CellStyle = sourceCell.CellStyle;
            //}
            //else
            //{
            //    ICellStyle newCellStyle = workbook.CreateCellStyle();
            //    newCellStyle.CloneStyleFrom(sourceCell.CellStyle);
            //    newCell.CellStyle = newCellStyle;
            //}

            // If there is a cell comment, copy
            if (sourceCell.CellComment != null) newCell.CellComment = sourceCell.CellComment;

            // If there is a cell hyperlink, copy
            if (sourceCell.Hyperlink != null) newCell.Hyperlink = sourceCell.Hyperlink;


            // Set the cell data value
            switch (sourceCell.CellType)
            {
                case CellType.Blank:
                    newCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.CellFormula = sourceCell.CellFormula;
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.String:
                    newCell.SetCellValue(sourceCell.RichStringCellValue);
                    break;
                case CellType.Unknown:
                    newCell.SetCellValue(sourceCell.StringCellValue);
                    break;
            }

        }

        /// <summary>
		/// Copy Row
		/// 傳入檔案路徑設定Sheet頁首頁尾
		/// </summary>  
		/// <param name="s_destPath">目標檔案路徑</param>
		/// <param name="s_SheetsName">欲設定的Sheet名稱</param>
		/// <param name="s_SetSide">設定的位置 EX:HC = Header.Center</param>
		/// <param name="s_Type">設定的內容 EX: &P (頁碼)</param>
		/// <returns>void</returns>
        public void SetPageHeaderFooter(string s_destPath, string s_SheetsName, string s_SetSide, string s_Type)
        {
            try
            {
                if (File.Exists(s_destPath))
                {
                    FileStream o_File = new FileStream(s_destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    IWorkbook o_workbook = WorkbookFactory.Create(o_File);
                    var o_resultSheet = o_workbook.GetSheet(s_SheetsName);

                    if (!s_SetSide.Equals(""))
                    {
                        if(o_resultSheet != null)
                        {
                            switch (s_SetSide)
                            {
                                case ("HC"):
                                    o_resultSheet.Header.Center = s_Type;
                                    break;
                                case ("HR"):
                                    o_resultSheet.Header.Right = s_Type; 
                                    break;
                                case ("HL"):
                                    o_resultSheet.Header.Left = s_Type; 
                                    break;
                                case ("FC"):
                                    o_resultSheet.Footer.Center = s_Type; 
                                    break;
                                case ("FR"):
                                    o_resultSheet.Footer.Right = s_Type; 
                                    break;
                                case ("FL"):
                                    o_resultSheet.Footer.Left = s_Type; 
                                    break;
                            }
                        }
                        else
                        {
                            o_File.Close();
                            o_File.Dispose();
                            throw new Exception("傳入參數有誤");
                        }
                    }

                    FileStream o_FileOut = new FileStream(s_destPath, FileMode.Create);
                    o_workbook.Write(o_FileOut);

                    o_FileOut.Close();
                    o_FileOut.Dispose();

                    o_File.Close();
                    o_File.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Copy Row
        /// 傳入檔案路徑設定Sheet頁首頁尾
        /// </summary>  
        /// <param name="s_destPath">目標檔案路徑</param>
        /// <param name="s_SheetsName">欲設定的Sheet名稱</param>
        /// <param name="IsZip">是否"將工作表放入單一頁面"</param>
        /// <returns>void</returns>
        public void SetPagePrint(string s_destPath, string s_SheetsName, bool IsZip)
        {
            try
            {
                if (File.Exists(s_destPath))
                {
                    FileStream o_File = new FileStream(s_destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    IWorkbook o_workbook = WorkbookFactory.Create(o_File);
                    var o_resultSheet = o_workbook.GetSheet(s_SheetsName);

                    if (o_resultSheet != null && IsZip)
                    {
                        o_resultSheet.Autobreaks = true; //forti
                        o_resultSheet.PrintSetup.FitHeight = 1;
                        o_resultSheet.PrintSetup.FitWidth = 1;
                    }

                    FileStream o_FileOut = new FileStream(s_destPath, FileMode.Create);
                    o_workbook.Write(o_FileOut);

                    o_FileOut.Close();
                    o_FileOut.Dispose();

                    o_File.Close();
                    o_File.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetRowBreak(ref ISheet o_resultSheet, int i_PageRowMax, bool IsSetBrk)
        {
            try
            {
                if (i_PageRowMax - 1 < o_resultSheet.LastRowNum)
                {
                    for (int iRow = i_PageRowMax - 1; iRow < o_resultSheet.LastRowNum; iRow += i_PageRowMax)
                    {
                        o_resultSheet.SetRowBreak(iRow);
                    }
                }
                else
                {
                    if (IsSetBrk)
                    {
                        o_resultSheet.SetRowBreak(i_PageRowMax - 1);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ScanTitleToSetRowBreak_withoutHiddenText(SheetInfo[] sInfo, bool bSpace = false)
        {
            int iPageCount = 0;

            for (int i = 0; i < sInfo.Length; i++)
            {
                //if (sInfo[i].TargetSheetName == null) continue;
                //if (sInfo[i].Title == null) continue;
                ISheet ISourceSheet;
                if (sInfo[i].TargetSheetName != null)
                {
                    ISourceSheet = workbook.GetSheet(sInfo[i].TargetSheetName);
                }
                else
                {
                    ISourceSheet = workbook.GetSheet(sInfo[i].SheetName);
                }
                if (ISourceSheet == null) continue;
                for (int iRow = ISourceSheet.FirstRowNum + 1; iRow < ISourceSheet.LastRowNum; iRow++)
                {
                    IRow ir = ISourceSheet.GetRow(iRow);
                    if (ir == null) continue;
                    for (int iCol = ISourceSheet.GetRow(iRow).FirstCellNum; iCol < ISourceSheet.GetRow(iRow).LastCellNum; iCol++)
                    {
                        if (ISourceSheet.GetRow(iRow).GetCell(iCol) == null)
                        {
                            continue;
                        }
                        if (sInfo[i].Title != null && ISourceSheet.GetRow(iRow).GetCell(iCol).ToString().IndexOf(sInfo[i].Title) >= 0)
                        {
                            if (bSpace)
                            {
                                for (int iR = 0; iR < sInfo[i].PageRowMax - (iRow % sInfo[i].PageRowMax); iR++)
                                {
                                    ISourceSheet.ShiftRows(iRow, ISourceSheet.LastRowNum, 1);
                                }
                                iRow += sInfo[i].PageRowMax - (iRow % sInfo[i].PageRowMax);
                            }
                            else
                            {
                                ISourceSheet.SetRowBreak(iRow - 1);
                            }
                        }
                        if (sInfo[i].PageCountTitle == null) continue;
                        if (ISourceSheet.GetRow(iRow).GetCell(iCol).ToString().IndexOf(sInfo[i].PageCountTitle) >= 0)
                        {
                            iPageCount += 1;
                            ISourceSheet.GetRow(iRow).GetCell(iCol + 1).SetCellValue(iPageCount + ISourceSheet.GetRow(iRow).GetCell(iCol + 1).ToString());
                        }
                    }
                }
                if (sInfo[i].TargetSheetName != null)
                {
                    int index = workbook.GetSheetIndex(sInfo[i].SheetName);
                    workbook.RemoveSheetAt(index);
                }
            }
        }

        /// <summary>
        /// 四捨五入
        /// </summary>
        /// <param name="value"></param>
        /// <param name="scale"></param>
        /// <returns></returns>
        private double RoundX(double value, int scale)
        {
            return Math.Round(value, scale, MidpointRounding.AwayFromZero); //四捨五入
        }
    }

    /// <summary>
    /// <param name="SheetName">來源頁名稱</param>
    /// <param name="PageRowMax">每頁行數</param>
    /// <param name="HeaderStartIndex">標題起啟列</param>
    /// <param name="HeaderLenght">標題列數</param>
    /// <param name="TargetSheetName">複製目的頁名稱</param>
    /// <param name="Title">標題名稱</param>
    /// <param name="DelSheet">是否刪除本頁</param>      
    /// </summary>
    public class SheetInfo
    {
        public string SheetName { set; get; }
        public int PageRowMax { set; get; }
        public int HeaderStartIndex { set; get; }
        public int HeaderLenght { set; get; }
        public string TargetSheetName { set; get; }
        public string Title { set; get; }
        public string PageCountTitle { set; get; }
        public int PageCount { set; get; }
        public string IdnTitle { set; get; }
        //public bool DelSheet { set; get; } = false;
    }

    /// <summary>
    /// 負值的顏色
    /// </summary>
    public enum NegativeNumberColorEnum
    {
        Red,
        Black
    }
}