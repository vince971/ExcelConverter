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

            // Starting Excel app
            var app = new Microsoft.Office.Interop.Excel.Application();

            // Opening the XLS File
            var workBook = app.Workbooks.Open(excelFilePath);

            // Adding new extension
            var outPutFile = Path.Combine(Path.GetDirectoryName(excelFilePath), System.IO.Path.GetFileNameWithoutExtension(excelFilePath) + ".xlsx");
            
            // Saving new file on XLSX Format
            workBook.SaveAs(Filename: outPutFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

            // Closing XLS File
            workBook.Close();

            // Exiting the Excel App
            app.Quit();
        }
    }
}
