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
        #region �����ʼ����
        private string fileName = null;
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private int headCount = -1;

        #endregion

        #region ������ʼ����
        private XSSFWorkbook exWorkBook = null;//word2007
        private MemoryStream exMs = null;

        private ICellStyle exTitleCellStyle = null;
        private ICellStyle exCellStyle = null;
        private IFont exTitleFont = null;
        private IFont exFont = null;
        #endregion

        private bool disposed;

        //����
        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            this.disposed = false;
        }
        //����
        public ExcelHelper(string fileName, int headcount)
        {
            this.fileName = fileName;
            this.headCount = headcount;
        }

        //����
        public ExcelHelper()
        {
            this.disposed = false;
        }

        public void Dispose()
        {
            //���ô�������Dispose�������ͷ��йܺͷ��й���Դ
            Dispose(true);
            //�ֶ�������Dispose�ͷ���Դ����ô�����������ǲ���Ҫ���ˣ�������ֹGC������������
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool IsDispose)
        {
            if (!this.disposed)
            {
                if (IsDispose)
                {
                    //�ͷŵ��������Դ
                    if (fs != null)
                    {
                        fs.Close();
                    }
                    //�ͷŵ���������Դ
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

        #region ����

        #region ���б�ͷ
        /// <summary>
        /// excel����תdatatable
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelToDataTable()
        {
            DataTable dt = new DataTable();
            ISheet sheet = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //HSSF����2007��ǰ�İ汾,XSSF����2007�汾�������ϵ�
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
                        IRow firstRow = sheet.GetRow(0);//��ȡ��FirstRowColumnNum����Ϊ��ͷ
                        if (firstRow != null)
                        {
                            int columnsCount = firstRow.LastCellNum;//�ܵ�����(�����յ���)
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
                            //���һ�еı��
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

        #region ˫�б�ͷ
        #region ˫�б�ͷ-ȡ��һ��,�ڶ��б�ͷ���
        /// <summary>
        /// ���б�ͷexcel����תdatatable
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelToDataTableByTwoHeads(List<ExcelHeadView> columnList)
        {
            DataTable dt = new DataTable();
            ISheet sheet = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //HSSF����2007��ǰ�İ汾,XSSF����2007�汾�������ϵ�
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
                        IRow firstRow = sheet.GetRow(0);//��ȡ��1�е���ͷ
                        IRow secondRow = sheet.GetRow(1);//��ȡ��2�е���ͷ
                        if (firstRow != null && secondRow != null)
                        {
                            //��ȡ׼ȷ����ͷ
                            GetFirstColumn(dt, firstRow, columnList);
                            GetSecondColumn(dt, secondRow, columnList);
                            //�����е�˳��
                            AdjustColumnIndex(dt, columnList);
                            notNullColumnsCount = dt.Columns.Count;
                            //���һ�еı��
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

        //��ȡ��һ�е���ͷ
        private void GetFirstColumn(DataTable dt, IRow firstRow, List<ExcelHeadView> columnList)
        {
            int columnsCount = firstRow.LastCellNum;//�ܵ�����(�����յ���)      
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

        //��ȡ�ڶ��е���ͷ
        private void GetSecondColumn(DataTable dt, IRow SecondRow, List<ExcelHeadView> columnList)
        {
            int columnsCount = SecondRow.LastCellNum;//�ܵ�����(�����յ���)      
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
        /// �����е�˳��
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
        /// ͨ����Ԫ���ʽ��ȡֵ
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
        /// ����ѧ������ת��Ϊ�ַ���
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
        /// ������������
        /// </summary>
        /// <param name="connectionString">Ŀ�������ַ�</param>
        /// <param name="TableName">Ŀ���</param>
        /// <param name="dt">Դ����</param>
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


        #region ����


        #region ��Sheet,��Sheet
        /// <summary>
        /// datatable����תexcel��Sheet  
        /// </summary>
        /// <returns></returns>
        public Stream DataTableToExcelBySheets(List<DataTable> dataTableList, List<string> sheetNameList, List<List<ExcelHeadView>> headArrList)
        {
            try
            {
                exWorkBook = new XSSFWorkbook();//2007   
                //��ʼ����Ԫ����ʽ
                exTitleCellStyle = exWorkBook.CreateCellStyle();
                exCellStyle = exWorkBook.CreateCellStyle();
                //��ʼ����Ԫ������
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
                        //���б�ͷ
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
                        //���б�ͷ,���б�ͷ��ϵĶ�sheet
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
                        //���б�ͷ
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
        /// datatable����תexcel��Sheet(���б�ͷ)
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
                //�п�����Ӧ,ֻ��Ӣ�ĺ�������Ч
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
        /// datatable����תexcel��Sheet(���б�ͷ)
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
                //������ͷ
                foreach (DataColumn dc in dt.Columns)
                {
                    SetHeadRow(firstHeadRow, dc, xHeadList);
                    SetHeadRow(secondHeadRow, dc, xHeadList);
                }
                //������ͷ
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
                //�����п�
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


        #region ������ͷ��ʽ
        /// <summary>
        /// ������ͷ
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
            //��һ���ͷ
            if (rowIndex == 0)
            {
                if (excelHeadView != null)
                {
                    //����
                    headRow.CreateCell(dc.Ordinal).SetCellValue(excelHeadView.FatherCellName);
                }
                else
                {
                    headRow.CreateCell(dc.Ordinal).SetCellValue(dc.ColumnName.ToString());
                }
            }
            //�ڶ����ͷ
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
            //������ʽ 
            SetCellStyle(headRow.Cells[dc.Ordinal], "title");
        }
        /// <summary>
        /// ������ͷ
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headRow"></param>
        private void ChangeHeadRow(ISheet sheet, XSSFRow headRow, List<ExcelHeadView> xHeadList)
        {
            //ȡ��һ���ֱ�ͷ
            List<ExcelHeadView> firstSonCellList = xHeadList.FindAll(a => a.IsFirstSonCell);
            for (int i = 0; i < headRow.Cells.Count; i++)
            {
                string span = firstSonCellList.Find(a => a.FatherCellName == headRow.Cells[i].ToString()) == null ? "" :
                    firstSonCellList.Find(a => a.FatherCellName == headRow.Cells[i].ToString()).Span.ToString();
                //�ϲ���
                if (span.Length > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, i, i + (int.Parse(span) - 1)));
                }
                //�ϲ���
                else
                {
                    sheet.AddMergedRegion(new CellRangeAddress(0, 1, i, i));
                }
            }
        }

        /// <summary>
        /// ���õ�Ԫ����ʽ
        /// </summary>
        /// <param name="cell"></param>
        private void SetCellStyle(ICell cell, string xTitle)
        {

            if (xTitle.Length > 0)
            {
                exTitleCellStyle.BorderBottom = BorderStyle.Thin;//ʵ�߱��
                exTitleCellStyle.BorderLeft = BorderStyle.Thin;
                exTitleCellStyle.BorderRight = BorderStyle.Thin;
                exTitleCellStyle.BorderTop = BorderStyle.Thin;
                exTitleCellStyle.Alignment = HorizontalAlignment.Center; //ˮƽ���ж���
                exTitleCellStyle.VerticalAlignment = VerticalAlignment.Center;//��ֱ���ж���
                exTitleFont.FontHeightInPoints = 10;
                exTitleFont.FontName = "����";
                exTitleFont.Boldweight = (short)FontBoldWeight.Bold;
                exTitleCellStyle.SetFont(exTitleFont);
                cell.CellStyle = exTitleCellStyle;
            }
            else
            {
                exCellStyle.BorderBottom = BorderStyle.Thin;//ʵ�߱��
                exCellStyle.BorderLeft = BorderStyle.Thin;
                exCellStyle.BorderRight = BorderStyle.Thin;
                exCellStyle.BorderTop = BorderStyle.Thin;
                exCellStyle.Alignment = HorizontalAlignment.Center; //ˮƽ���ж���
                exCellStyle.VerticalAlignment = VerticalAlignment.Center;//��ֱ���ж���
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
                //��ʼ����Ԫ����ʽ
                exTitleCellStyle = exWorkBook.CreateCellStyle();
                exCellStyle = exWorkBook.CreateCellStyle();
                //��ʼ����Ԫ������
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
                    //�п�����Ӧ,ֻ��Ӣ�ĺ�������Ч
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