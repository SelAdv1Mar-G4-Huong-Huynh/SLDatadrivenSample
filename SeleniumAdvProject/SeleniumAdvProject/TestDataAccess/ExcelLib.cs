using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.TestDataAccess
{
    class ExcelLib
    {
        private static DataTable ExcelToDataTable(string fileName)
        {
            //Open File and returns as Stream
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //set the frist row as column name
            excelReader.IsFirstRowAsColumnNames = true;
            // Get all table
            DataSet result = excelReader.AsDataSet();
            DataTableCollection table = result.Tables;
            DataTable resultTable = table["DataSet"];
            return resultTable;
        }
        static List<DataCollection> dataCollection = new List<DataCollection>();
        public static void PopulateInCollection(string fileName)
        {
            DataTable table = ExcelToDataTable(fileName);
            // iterate through the rows and columns of the table
            for (int row = 1; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    DataCollection dtTable = new DataCollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString(),
                    };
                    dataCollection.Add(dtTable);
                }

            }
        }

        public static string ReadData(int rowNumber, string columnName)
        {
            try
            {
                string data = (from colData in dataCollection
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).SingleOrDefault();
                return data.ToString();
            }
            catch (Exception e) { return String.Format("Error!{0}",e.Message);}
        }
    }
    public class DataCollection
    {
        public int rowNumber { get; set; }
        public string colName {get;set;}
        public string colValue{get;set;}
    }


}
