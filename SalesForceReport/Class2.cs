using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace SalesForceReport
{
    class ExcelReader
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public ExcelReader(string path, int Sheet) {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = (Worksheet)wb.Worksheets[Sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j] != null)
                return ws.Cells[i, j];
            else
                return "";
        }
        
    }
}
