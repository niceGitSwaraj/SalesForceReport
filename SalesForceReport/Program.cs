
using System;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SalesForceReport
{
    class CheckReport
    {
        public class ReportData
        {
            public string ID { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
        }
        static void Main(string[] args)
        {
            ExcelApiTest exp = new ExcelApiTest(@"C:\\ExcelFiles\\TestData.xlsx");
            List<string> Dates = new List<string>();
            string dateColumn = "D1";
            Console.Write("Please enter the Start Date: ");            
            string startDate = Console.ReadLine();
            Console.Write("Please enter the End Date: ");
            string endDate = Console.ReadLine();
            Dates.Add(startDate);
            Dates.Add(endDate); 
            List<int> dateRows = exp.GetDateRows("Sheet1",dateColumn,Dates);
            
            List<string> columnNames = new List<string>();
            columnNames.Add("ID");
            columnNames.Add("FirstName");
            columnNames.Add("LastName");
            columnNames.Add("StartDate");
            columnNames.Add("EndDate");
            List<string> cellValues = exp.GetCellData("Sheet1", columnNames, dateRows);
            List<ReportData> _data = new List<ReportData>();
            foreach (var row in cellValues){
                
                string[] dataBlock = row.Split(",");
                
                _data.Add(new ReportData
                {
                    ID = dataBlock[0],
                    FirstName = dataBlock[1],
                    LastName = dataBlock[2],
                    StartDate = DateTime.FromOADate(Convert.ToDouble(dataBlock[3])),
                    EndDate = DateTime.FromOADate(Convert.ToDouble(dataBlock[4]))
                });
            }
             string json = JsonSerializer.Serialize(_data);
             //File.WriteAllText(@"D:\path.json", json);
             Console.WriteLine(json);


        }

    }
}
