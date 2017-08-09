using System;
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
        private ClosedXML.Excel.XLWorkbook workbook = null;

        private ClosedXML.Excel.IXLWorksheet workSheet = null;

        private ClosedXML.Excel.IXLRange range = null;

        private ClosedXML.Excel.IXLTable table = null;

        public Excel(FileInfo fullExcelPath)
        {
            this.workbook = new ClosedXML.Excel.XLWorkbook(fullExcelPath.DirectoryName);
            this.workSheet = workbook.Worksheet(1);
            this.range = workSheet.Range(workSheet.FirstCellUsed(), workSheet.LastCellUsed());
            this.table = range.AsTable();
        }

        public DataTable AddColumnToDataTable()
        {
            return null;
        }

        public DataTable AddToDataTable()
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

        public void InsertRow()
        {
            
        }

        public void DeleteRow()
        {

        }

        public void DeleteValue()
        {

        }

        public static void Create()
        {

        }
    }
}
