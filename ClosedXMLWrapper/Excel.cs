﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXMLWrapper
{
    public class Excel
    {
        private ClosedXML.Excel.IXLWorksheet workSheet = null;
        private ClosedXML.Excel.XLWorkbook workbook = null;
        private ClosedXML.Excel.IXLRange range = null;
        private ClosedXML.Excel.IXLTable table = null;

        public Excel(FileInfo fullExcelPath)
        {
            this.range = workSheet.Range(workSheet.FirstCellUsed(), workSheet.LastCellUsed());
            this.workbook = new ClosedXML.Excel.XLWorkbook(fullExcelPath.DirectoryName);
            this.workSheet = workbook.Worksheet(1);
            this.table = range.AsTable();
        }

        public Excel()
        { }

        public DataTable ConvertToDataTable()
        {
            return null;
        }

        public static Excel CreateExcel(List<string> columnsName, List<List<string>> rowsValue, FileInfo excelFile)
        {
            DataTable dt = new DataTable();

            ClosedXML.Excel.XLWorkbook workbook = CreateTable(dt, columnsName, rowsValue);

            workbook.SaveAs(excelFile.FullName);

            return CreateExcel(excelFile);
        }

        private static Excel CreateExcel(FileInfo excelFile)
        {
            Excel excel = new Excel(excelFile);

            return excel;
        }

        public class Read
        {
            public Dictionary<string, string> SearchForColumn()
            {
                return null;
            }

            public Dictionary<string, string> SearchForRow()
            {
                return null;
            }

            public Dictionary<string, string> SearchForValue()
            {
                return null;
            }
        }

        public class Update
        {
            public DataTable AddColumnToDataTable()
            {
                return null;
            }

            public void UpdateColumn()
            { }

            public void InsertColumn()
            { }

            public void UpdateRow()
            { }

            public void InsertRow()
            { }

            public void ReplaceValue()
            { }
        }

        public class Delete
        {
            public void DeleteColumn()
            { }

            public void DeleteRow()
            { }

            public void DeleteValue()
            { }
        }

        private static ClosedXML.Excel.XLWorkbook CreateTable(DataTable dt, List<string> columnsName, List<List<string>> rowsValue)
        {
            AddColumn(dt, columnsName);

            AddRow(dt, rowsValue);

            ClosedXML.Excel.IXLWorksheet ws = new ClosedXML.Excel.XLWorkbook().Worksheets.Add(dt);

            return ws.Workbook;
        }

        private static DataTable AddColumn(DataTable dt, List<string> columnsName)
        {
            foreach (string column in columnsName)
                dt.Columns.Add(column, typeof(String));

            return dt;
        }

        private static DataTable AddRow(DataTable dt, List<List<string>> rowsValue)
        {
            for (int i = 0; i < rowsValue.Count; i++)
            {
                DataRow dataRow = dt.NewRow();

                for (int x = 0; x < rowsValue[i].Count; x++)
                    dataRow[x] = rowsValue[i][x];

                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        private Dictionary<string, string> SearchColumn(ClosedXML.Excel.XLWorkbook workbook, string columnName)
        {

        }

        private Dictionary<string, string> SearchRow()
        { }

        private Dictionary<string, string> Search()
        {

        }

        private void ErrorHandler()
        { }
    }
}
