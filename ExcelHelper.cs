using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.Web.UI;
using System.Security;

namespace NPOIDemo
{
    public class ExcelHelper : IDisposable
    {
        #region 导入初始变量
        private string fileName = null;
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private int headCount = -1;

        #endregion

        #region 导出初始变量
        private XSSFWorkbook exWorkBook = null;//word2007
        private MemoryStream exMs = null;

        private ICellStyle exTitleCellStyle = null;
        private ICellStyle exCellStyle = null;
        private IFont exTitleFont = null;
        private IFont exFont = null;
        #endregion

        private bool disposed;

        //导入
        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            this.disposed = false;
        }
        //导入
        public ExcelHelper(string fileName, int headcount)
        {
            this.fileName = fileName;
            this.headCount = headcount;
        }

        //导出
        public ExcelHelper()
        {
            this.disposed = false;
        }

        public void Dispose()
        {
            //调用带参数的Dispose方法，释放托管和非托管资源
            Dispose(true);
            //手动调用了Dispose释放资源，那么析构函数就是不必要的了，这里阻止GC调用析构函数
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool IsDispose)
        {
            if (!this.disposed)
            {
                if (IsDispose)
                {
                    //释放导入的流资源
                    if (fs != null)
                    {
                        fs.Close();
                    }
                    //释放导出的流资源
                    if (exMs != null)
                    {
                        exMs.Close();
                    }
                }
            }
            fs =  null;
            exMs = null;
            this.disposed = true;
        }

        #region 导入

        #region 单行表头
        /// <summary>
        /// excel数据转datatable
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelToDataTable()
        {
            DataTable dt = new DataTable();
            ISheet sheet = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //HSSF适用2007以前的版本,XSSF适用2007版本及其以上的
                if (fileName.Substring(fileName.LastIndexOf('.')) == ".xlsx")//2007
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.Substring(fileName.LastIndexOf('.')) == ".xls")//2003
                {
                    workbook = new HSSFWorkbook(fs);
                }
                if (workbook != null)
                {
                    sheet = workbook.GetSheetAt(0);
                    if (sheet != null)
                    {
                        IRow firstRow = sheet.GetRow(0);//获取第FirstRowColumnNum行作为列头
                        if (firstRow != null)
                        {
                            int columnsCount = firstRow.LastCellNum;//总的列数(包含空的列)
                            int notNullColumnsCount = columnsCount;
                            for (int i = firstRow.FirstCellNum; i < columnsCount; i++)
                            {
                                ICell cell = firstRow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellValue = "";
                                    if (cell.CellType == CellType.Numeric)
                                    {
                                        cellValue = cell.NumericCellValue.ToString().Trim();
                                    }
                                    else
                                    {
                                        cellValue = cell.StringCellValue.ToString().Trim();
                                    }
                                    if (cellValue.Trim().Length > 0)
                                    {
                                        DataColumn dc = new DataColumn();
                                        dc.ColumnName = cellValue;
                                        dt.Columns.Add(dc);
                                    }
                                    else
                                    {
                                        notNullColumnsCount -= 1;
                                    }
                                }
                                else
                                {
                                    notNullColumnsCount -= 1;
                                }
                            }
                            //最后一行的标号
                            int rowCount = sheet.LastRowNum;
                            for (int j = firstRow.RowNum + 1; j <= rowCount; j++)
                            {
                                if (sheet.GetRow(j) != null)
                                {
                                    IRow row = sheet.GetRow(j);
                                    if (row == null)
                                    {
                                        continue;
                                    }
                                    DataRow dr = dt.NewRow();
                                    bool isNull = true;
                                    for (int k = 0; k < notNullColumnsCount; k++)
                                    {
                                        ICell cell = row.GetCell(k);
                                        if (cell != null)
                                        {
                                            if (cell.CellType == CellType.Formula)
                                            {
                                                cell.SetCellType(CellType.String);
                                            }
                                            string cellValue = GetTypeValueByCell(cell);
                                            if (cellValue.Trim().Length > 0)
                                            {
                                                isNull = false;
                                            }
                                            dr[k] = ChangeDataToString(cellValue);
                                        }
                                    }
                                    if (isNull)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        dt.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region 双行表头
        #region 双行表头-取第一行,第二行表头混合
        /// <summary>
        /// 两行表头excel数据转datatable
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelToDataTableByTwoHeads(List<ExcelHeadView> columnList)
        {
            DataTable dt = new DataTable();
            ISheet sheet = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //HSSF适用2007以前的版本,XSSF适用2007版本及其以上的
                if (fileName.Substring(fileName.LastIndexOf('.')) == ".xlsx")//2007
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.Substring(fileName.LastIndexOf('.')) == ".xls")//2003
                {
                    workbook = new HSSFWorkbook(fs);
                }
                if (workbook != null)
                {
                    sheet = workbook.GetSheetAt(0);
                    if (sheet != null)
                    {
                        int notNullColumnsCount = 0;
                        IRow firstRow = sheet.GetRow(0);//获取第1行的列头
                        IRow secondRow = sheet.GetRow(1);//获取第2行的列头
                        if (firstRow != null && secondRow != null)
                        {
                            //获取准确的列头
                            GetFirstColumn(dt, firstRow, columnList);
                            GetSecondColumn(dt, secondRow, columnList);
                            //调整列的顺序
                            AdjustColumnIndex(dt, columnList);
                            notNullColumnsCount = dt.Columns.Count;
                            //最后一行的标号
                            int rowCount = sheet.LastRowNum;
                            for (int j = secondRow.RowNum + 1; j <= rowCount; j++)
                            {
                                if (sheet.GetRow(j) != null)
                                {
                                    IRow row = sheet.GetRow(j);
                                    if (row == null)
                                    {
                                        continue;
                                    }
                                    DataRow dr = dt.NewRow();
                                    bool isNull = true;
                                    for (int k = 0; k < notNullColumnsCount; k++)
                                    {
                                        ICell cell = row.GetCell(k);
                                        if (cell != null)
                                        {
                                            if (cell.CellType == CellType.Formula)
                                            {
                                                cell.SetCellType(CellType.String);
                                            }
                                            string cellValue = GetTypeValueByCell(cell);
                                            if (cellValue.Trim().Length > 0)
                                            {
                                                isNull = false;
                                            }
                                            dr[k] = ChangeDataToString(cellValue);
                                        }
                                    }
                                    if (isNull)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        dt.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //获取第一行的列头
        private void GetFirstColumn(DataTable dt, IRow firstRow, List<ExcelHeadView> columnList)
        {
            int columnsCount = firstRow.LastCellNum;//总的列数(包含空的列)      
            for (int i = firstRow.FirstCellNum; i < columnsCount; i++)
            {
                ICell cell = firstRow.GetCell(i);
                if (cell != null)
                {
                    string cellValue = "";
                    if (cell.CellType == CellType.Numeric)
                    {
                        cellValue = cell.NumericCellValue.ToString().Trim();
                    }
                    else
                    {
                        cellValue = cell.StringCellValue.ToString().Trim();
                    }
                    if (cellValue.Trim().Length > 0)
                    {
                        if (columnList.Find(a => a.CellName == cellValue.Trim()) != null)
                        {
                            DataColumn dc = new DataColumn();
                            dc.ColumnName = cellValue;
                            dt.Columns.Add(dc);
                        }
                    }
                }
            }
        }

        //获取第二行的列头
        private void GetSecondColumn(DataTable dt, IRow SecondRow, List<ExcelHeadView> columnList)
        {
            int columnsCount = SecondRow.LastCellNum;//总的列数(包含空的列)      
            for (int i = SecondRow.FirstCellNum; i < columnsCount; i++)
            {
                ICell cell = SecondRow.GetCell(i);
                if (cell != null)
                {
                    string cellValue = "";
                    if (cell.CellType == CellType.Numeric)
                    {
                        cellValue = cell.NumericCellValue.ToString().Trim();
                    }
                    else
                    {
                        cellValue = cell.StringCellValue.ToString().Trim();
                    }
                    if (cellValue.Trim().Length > 0)
                    {
                        if (columnList.Find(a => a.CellName == cellValue.Trim()) != null)
                        {
                            DataColumn dc = new DataColumn();
                            dc.ColumnName = cellValue;
                            dt.Columns.Add(dc);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 调整列的顺序
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="singleColumnList"></param>
        private void AdjustColumnIndex(DataTable dt, List<ExcelHeadView> columnList)
        {

            foreach (var item in columnList)
            {
                if (dt.Columns.Contains(item.CellName))
                {
                    dt.Columns[item.CellName].SetOrdinal(item.index);
                }
            }
        }
        #endregion
        #endregion

        /// <summary>
        /// 通过单元格格式获取值
        /// </summary>
        /// <returns></returns>
        private String GetTypeValueByCell(ICell cell)
        {

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return "";
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString().Trim();
                    }
                    else
                    {
                        return cell.NumericCellValue.ToString().Trim();
                    }
                case CellType.Formula:
                    return cell.CellFormula.ToString().Trim();
                default:
                    return cell.StringCellValue.ToString().Trim();
            }
        }


        /// <summary>
        /// 将科学计数法转化为字符串
        /// </summary>
        /// <param name="strData"></param>
        /// <returns></returns>
        private string ChangeDataToString(string strData)
        {
            decimal dData = 0.0M;
            if (strData.Contains("E+"))
            {
                dData = Convert.ToDecimal(Decimal.Parse(strData.ToString(), System.Globalization.NumberStyles.Float));
                return dData.ToString().Replace('M', ' ');
            }
            else
            {
                return strData;
            }
        }

        /// <summary>
        /// 批量导入数据
        /// </summary>
        /// <param name="connectionString">目标连接字符</param>
        /// <param name="TableName">目标表</param>
        /// <param name="dt">源数据</param>
        public void SqlBulkCopyBatch(string connKey, string TableName, DataTable dt)
        {

            string connectionString = Common.DataCache.Inst().GetConnection(connKey);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
                {
                    try
                    {
                        sqlbulkcopy.DestinationTableName = TableName;
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sqlbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                        }
                        sqlbulkcopy.WriteToServer(dt);
                    }
                    catch (System.Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }

        #endregion


        #region 导出


        #region 单Sheet,多Sheet
        /// <summary>
        /// datatable数据转excel多Sheet  
        /// </summary>
        /// <returns></returns>
        public Stream DataTableToExcelBySheets(List<DataTable> dataTableList, List<string> sheetNameList, List<List<ExcelHeadView>> headArrList)
        {
            try
            {
                exWorkBook = new XSSFWorkbook();//2007   
                //初始化单元格样式
                exTitleCellStyle = exWorkBook.CreateCellStyle();
                exCellStyle = exWorkBook.CreateCellStyle();
                //初始化单元格字体
                exFont = exWorkBook.CreateFont();
                exTitleFont = exWorkBook.CreateFont();
                ISheet[] sheets = new ISheet[sheetNameList.Count];
                for (int i = 0; i < sheets.Length; i++)
                {
                    sheets[i] = exWorkBook.CreateSheet(sheetNameList[i].ToString().Trim());
                }
                exMs = new MemoryStream();
                for (int a = 0; a < sheets.Length; a++)
                {
                    if (dataTableList[a] != null)
                    {
                        //两行表头
                        if (headArrList != null && headArrList.Count() > 0 && dataTableList.Count == headArrList.Count)
                        {
                            for (int i = 0; i < headArrList.Count; i++)
                            {
                                if (headArrList[i][0].TableName == sheetNameList[a])
                                {
                                    DataTableToExcelByTwoHeads(sheets[a], dataTableList[a], headArrList[i]);
                                }
                            }
                        }
                        //单行表头,两行表头混合的多sheet
                        else if (headArrList != null && headArrList.Count() > 0 && dataTableList.Count != headArrList.Count)
                        {
                            if (headArrList.Find(k => k[0].TableName == sheetNameList[a]) != null)
                            {
                                List<ExcelHeadView> list = headArrList.Find(k => k[0].TableName == sheetNameList[a]);
                                DataTableToExcelByTwoHeads(sheets[a], dataTableList[a], list);
                            }
                            else
                            {
                                DataTableToExcelBySingleHead(sheets[a], dataTableList[a]);
                            }
                        }
                        //单行表头
                        else
                        {
                            DataTableToExcelBySingleHead(sheets[a], dataTableList[a]);
                        }
                    }
                }
                exWorkBook.Write(exMs);
                exMs.Flush();
                return exMs;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// datatable数据转excel多Sheet(单行表头)
        /// </summary>
        /// <returns></returns>
        private void DataTableToExcelBySingleHead(ISheet xSheet, DataTable dt)
        {
            try
            {

                XSSFRow headRow = (XSSFRow)xSheet.CreateRow(0);
                foreach (DataColumn dc in dt.Columns)
                {
                    headRow.CreateCell(dc.Ordinal).SetCellValue(dc.ColumnName.ToString().Trim());
                    SetCellStyle(headRow.Cells[dc.Ordinal], "title");
                }
                int rowIndex = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    XSSFRow xssfRow = (XSSFRow)xSheet.CreateRow(rowIndex);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        xssfRow.CreateCell(dc.Ordinal).SetCellValue(dr[dc].ToString().Trim());
                        SetCellStyle(xssfRow.Cells[dc.Ordinal], "");
                    }
                    rowIndex += 1;
                }
                //列宽自适应,只对英文和数字有效
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSheet.SetColumnWidth(i, 15 * 256);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// datatable数据转excel多Sheet(两行表头)
        /// </summary>
        /// <param name="xWorkBook"></param>
        /// <param name="xSheet"></param>
        /// <param name="dt"></param>
        private void DataTableToExcelByTwoHeads(ISheet xSheet, DataTable dt, List<ExcelHeadView> xHeadList)
        {
            try
            {
                XSSFRow firstHeadRow = (XSSFRow)xSheet.CreateRow(0);
                XSSFRow secondHeadRow = (XSSFRow)xSheet.CreateRow(1);
                //创建表头
                foreach (DataColumn dc in dt.Columns)
                {
                    SetHeadRow(firstHeadRow, dc, xHeadList);
                    SetHeadRow(secondHeadRow, dc, xHeadList);
                }
                //调整表头
                ChangeHeadRow(xSheet, firstHeadRow, xHeadList);
                int rowIndex = 2;
                foreach (DataRow dr in dt.Rows)
                {
                    XSSFRow xssfRow = (XSSFRow)xSheet.CreateRow(rowIndex);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dr.IsNull(dc))
                        {
                            xssfRow.CreateCell(dc.Ordinal).SetCellValue("");
                            SetCellStyle(xssfRow.Cells[dc.Ordinal], "");
                        }
                        else
                        {
                            xssfRow.CreateCell(dc.Ordinal).SetCellValue(dr[dc].ToString().Trim());
                            SetCellStyle(xssfRow.Cells[dc.Ordinal], "");
                        }
                    }
                    rowIndex += 1;
                }
                //设置列宽
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSheet.SetColumnWidth(i, 15 * 256);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        #region 调整表头样式
        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="headRow"></param>
        /// <param name="dc"></param>
        private void SetHeadRow(XSSFRow headRow, DataColumn dc, List<ExcelHeadView> xHeadList)
        {
            int rowIndex = headRow.RowNum;
            ExcelHeadView excelHeadView = null;
            if (xHeadList.Count > 0)
            {
                excelHeadView = xHeadList.Find(a => a.CellName == dc.ColumnName);
            }
            //第一层表头
            if (rowIndex == 0)
            {
                if (excelHeadView != null)
                {
                    //创建
                    headRow.CreateCell(dc.Ordinal).SetCellValue(excelHeadView.FatherCellName);
                }
                else
                {
                    headRow.CreateCell(dc.Ordinal).SetCellValue(dc.ColumnName.ToString());
                }
            }
            //第二层表头
            else
            {
                if (excelHeadView != null)
                {
                    headRow.CreateCell(dc.Ordinal).SetCellValue(excelHeadView.CellName);
                }
                else
                {
                    headRow.CreateCell(dc.Ordinal).SetCellValue("");
                }
            }
            //设置样式 
            SetCellStyle(headRow.Cells[dc.Ordinal], "title");
        }
        /// <summary>
        /// 调整表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headRow"></param>
        private void ChangeHeadRow(ISheet sheet, XSSFRow headRow, List<ExcelHeadView> xHeadList)
        {
            //取第一个字表头
            List<ExcelHeadView> firstSonCellList = xHeadList.FindAll(a => a.IsFirstSonCell);
            for (int i = 0; i < headRow.Cells.Count; i++)
            {
                string span = firstSonCellList.Find(a => a.FatherCellName == headRow.Cells[i].ToString()) == null ? "" :
                    firstSonCellList.Find(a => a.FatherCellName == headRow.Cells[i].ToString()).Span.ToString();
                //合并列
                if (span.Length > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, i, i + (int.Parse(span) - 1)));
                }
                //合并行
                else
                {
                    sheet.AddMergedRegion(new CellRangeAddress(0, 1, i, i));
                }
            }
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="cell"></param>
        private void SetCellStyle(ICell cell, string xTitle)
        {

            if (xTitle.Length > 0)
            {
                exTitleCellStyle.BorderBottom = BorderStyle.Thin;//实线表格
                exTitleCellStyle.BorderLeft = BorderStyle.Thin;
                exTitleCellStyle.BorderRight = BorderStyle.Thin;
                exTitleCellStyle.BorderTop = BorderStyle.Thin;
                exTitleCellStyle.Alignment = HorizontalAlignment.Center; //水平居中对齐
                exTitleCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中对齐
                exTitleFont.FontHeightInPoints = 10;
                exTitleFont.FontName = "宋体";
                exTitleFont.Boldweight = (short)FontBoldWeight.Bold;
                exTitleCellStyle.SetFont(exTitleFont);
                cell.CellStyle = exTitleCellStyle;
            }
            else
            {
                exCellStyle.BorderBottom = BorderStyle.Thin;//实线表格
                exCellStyle.BorderLeft = BorderStyle.Thin;
                exCellStyle.BorderRight = BorderStyle.Thin;
                exCellStyle.BorderTop = BorderStyle.Thin;
                exCellStyle.Alignment = HorizontalAlignment.Center; //水平居中对齐
                exCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中对齐
                exCellStyle.SetFont(exFont);
                cell.CellStyle = exCellStyle;
            }
        }

        #endregion




        #endregion


        #endregion

        public Stream DataTableToExcelBySheet(DataTable dt)
        {
            try
            {
                exWorkBook = new XSSFWorkbook();//2007   
                //初始化单元格样式
                exTitleCellStyle = exWorkBook.CreateCellStyle();
                exCellStyle = exWorkBook.CreateCellStyle();
                //初始化单元格字体
                exFont = exWorkBook.CreateFont();
                exTitleFont = exWorkBook.CreateFont();
                ISheet sheet = exWorkBook.CreateSheet();
                exMs = new MemoryStream();
                if (dt != null)
                {
                    XSSFRow headRow = (XSSFRow)sheet.CreateRow(0);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        headRow.CreateCell(dc.Ordinal).SetCellValue(dc.ColumnName.ToString().Trim());
                        SetCellStyle(headRow.Cells[dc.Ordinal], "title");
                    }
                    int rowIndex = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        XSSFRow xssfRow = (XSSFRow)sheet.CreateRow(rowIndex);
                        foreach (DataColumn dc in dt.Columns)
                        {
                            xssfRow.CreateCell(dc.Ordinal).SetCellValue(dr[dc].ToString().Trim());
                            SetCellStyle(xssfRow.Cells[dc.Ordinal], "");
                        }
                        rowIndex += 1;
                    }
                    //列宽自适应,只对英文和数字有效
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        sheet.SetColumnWidth(i, 15 * 256);
                    }
                }
                exWorkBook.Write(exMs);
                exMs.Flush();
                return exMs;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

    }

}