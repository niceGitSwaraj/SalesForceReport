using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;


namespace SalesForceReport
{
    class ExcelApiTest
    {
        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        public string xlFilePath;

        public ExcelApiTest(string xlFilePath)
        {
            this.xlFilePath = xlFilePath;
        }

        public void OpenExcel()
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets = new Hashtable();
            int count = 1;
            // Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }

        public void CloseExcel()
        {
            workbook.Close(false, xlFilePath, null); // Close the connection to workbook
            Marshal.FinalReleaseComObject(workbook); // Release unmanaged object references.
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        public List<string> GetCellData(string sheetName, List<string> colName, List<int> rowNumbers)
        {
            OpenExcel();

            List<string> value = new List<string>();
            int sheetValue = 0;
            int colNumber = 0;

            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;
                for (var row = rowNumbers[0]; row <= rowNumbers[1]; row++) {
                    List<string> temp = new List<string>();
                    foreach (var column in colName)
                    {
                        for (int i = 1; i <= range.Columns.Count; i++)
                        {
                            string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);

                            if (colNameValue.ToLower() == column.ToLower())
                            {
                                colNumber = i;
                                break;
                            }
                        }
                        temp.Add(Convert.ToString((range.Cells[row, colNumber] as xl.Range).Value2)+",");                        
                    }
                    string concat = "";
                    foreach (var i in temp) {
                        concat += i;
                    }
                    value.Add(concat);
                }
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }

        public List<int> GetDateRows(string sheetName,string dateColumn ,List<string> Dates)
        {
            OpenExcel();

            string colToCheck = dateColumn;//get this as user input
            int sheetValue = 0;
            List<int> result = new List<int>();
            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range resultRange;// = worksheet.UsedRange;
                var colRange = worksheet.Range[colToCheck];//get the range object where you want to search from
                foreach (var date in Dates)
                {
                    resultRange = colRange.Find(

                                    What: date,
                                    LookIn: xl.XlFindLookIn.xlValues,
                                    LookAt: xl.XlLookAt.xlPart,
                                    SearchOrder: xl.XlSearchOrder.xlByRows,
                                    SearchDirection: xl.XlSearchDirection.xlNext
                    );
                    result.Add(resultRange.Row);
                }
            }
            CloseExcel();
            return result;
        }

    }
}