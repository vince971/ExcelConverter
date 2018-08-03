using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    /// <summary>
    /// Converting Xls To Csv Class
    /// </summary>
    public static class XlsToCsv
    {
        /// <summary>
        /// Converting a XLS File to a CSV File
        /// </summary>
        /// <param name="excelFilePath">Full xls file path</param>
        /// <param name="csvOutputFile">full csv file output path</param>
        /// <param name="worksheetNumber">Worksheet number in the xls file</param>
        public static void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            // Checking if source file exist
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException(excelFilePath);

            // Checking if output file exist
            if (File.Exists(csvOutputFile))
                csvOutputFile = csvOutputFile.Replace(".csv", $@"{DateTime.Now.ToShortDateString().Replace(@"/", "")}.csv");
            
            // Connection String
            var conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            var con = new OleDbConnection(conStr);
            var dataTable = new DataTable();

            try
            {
                con.Open();

                var schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber)
                    throw new ArgumentException($"The worksheet number provided cannot be found in the spreadsheet");

                var worksheet = schemaTable.Rows[worksheetNumber - 1][$"table_name"].ToString().Replace("'", "");
                var sqlRequest = $"SELECT * FROM [{worksheet}]";
                var dataAdapter = new OleDbDataAdapter(sqlRequest, con);

                dataAdapter.Fill(dataTable);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
            finally
            {
                con.Close();
            }

            // Create output file.
            var fs = new FileStream(csvOutputFile, FileMode.Create);

            using (var wtr = new StreamWriter(fs, Encoding.Default))
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        if(!firstLine)
                            wtr.Write($";");
                        else
                            firstLine = false;

                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(data);
                    }
                    wtr.WriteLine();
                }
            }
        }
    }
}
