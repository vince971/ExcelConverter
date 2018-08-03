using System.IO;

namespace Converter
{
    /// <summary>
    /// Converting XLS to XLSX class
    /// </summary>
    public static class XlsToXlsx
    {
        /// <summary>
        /// Converting XLS file to XLSX
        /// </summary>
        /// <param name="excelFilePath">Full file path</param>
        public static void ConvertXlsToXlsx(string excelFilePath)
        {
            // Checking if source file exist
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException(excelFilePath);

            var app = new Microsoft.Office.Interop.Excel.Application();

            var workBook = app.Workbooks.Open(excelFilePath);

            // Adding new extension
            var xlsxFile = excelFilePath + "x";

            // Saving new file with xlsx format
            workBook.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

            workBook.Close();
            app.Quit();
        }
    }
}
