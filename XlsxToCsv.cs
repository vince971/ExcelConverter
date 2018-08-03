using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Converter
{ 
    /// <summary>
    /// Converting Xlsx To Csv Class
    /// </summary>
    public class XlsxToCsv
    {
        /// <summary>
        /// Alphabet to find columns
        /// </summary>
        private const string ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        /// <summary>
        /// Converting a XLSX File to a CSV File
        /// </summary>
        /// <param name="excelFilePath">Full xlsx file path</param>
        public static bool Converter_XlsxToCsv(string excelFilePath)
        {
            var fileToWrite = System.IO.Path.Combine(Path.GetDirectoryName(excelFilePath), System.IO.Path.GetFileNameWithoutExtension(excelFilePath) + ".csv");
            var outputFile = new StreamWriter(fileToWrite, false, Encoding.UTF8);

            try
            {
                using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(excelFilePath, true))
                {
                    WorkbookPart workbookPart = myDoc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.GetPartById("rId1") as WorksheetPart;

                    int dateIndex = -1;
                    int styleIndex = 0;

                    List<int> countDateForm = new List<int>();

                    foreach (CellFormat format in workbookPart.WorkbookStylesPart.Stylesheet.CellFormats)
                    {
                        //if (format.NumberFormatId == 14 )
                        //{
                        //    dateIndex = styleIndex;
                        //    break;
                        //}
                        if (format.NumberFormatId == 14)
                            countDateForm.Add(styleIndex);

                        ++styleIndex;
                    }

                    var sharedDico = new Dictionary<string, string>();


                    int i = 0;

                    workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ToList().ForEach(e => sharedDico.Add(i++.ToString(), e.InnerText.Replace(";", ":").Replace("\n", "").Replace("'", " ")));

                    StringBuilder lineBuilder = null;
                    bool isSharedString = false;
                    char previousColumn = default(char);
                    bool mustAddValue = false;
                    string[] Col;
                    int nb_column = 0;
                    int column = 0;
                    bool firstDetec = false;
                    bool isDate = false;
                    bool isCellFormul = false;

                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                    while (reader.Read())
                    {
                        if (reader.LocalName == "row")
                        {
                            if (reader.IsStartElement)
                            {
                                lineBuilder = new StringBuilder();
                                mustAddValue = false;
                            }
                            else
                            {
                                Col = lineBuilder.ToString().Split(';');
                                column = Col.Count();

                                if (firstDetec == false)
                                {
                                    nb_column = column;
                                    firstDetec = true;
                                }

                                if (nb_column == column && firstDetec == true)
                                    outputFile.WriteLine(lineBuilder.ToString());

                                else if (nb_column > column && firstDetec == true)
                                {
                                    int nbColMissing = nb_column - column;
                                    string sep = null;

                                    for (int ajoutCol = 0; ajoutCol < nbColMissing; ajoutCol++)
                                        sep += ';';

                                    outputFile.WriteLine(lineBuilder.ToString() + sep);
                                }
                                else if (nb_column < column && firstDetec == true)
                                {
                                    int nbColMissing = column - nb_column;
                                    int countLastChar = 0;

                                    for (int ajoutCol = lineBuilder.Length - 1; ajoutCol >= 0; ajoutCol--)
                                    {
                                        if (lineBuilder[ajoutCol] != ';') break;
                                        countLastChar++;
                                    }

                                    if (countLastChar == nbColMissing)
                                        outputFile.WriteLine(lineBuilder.ToString().Substring(0, lineBuilder.Length - nbColMissing));
                                    else
                                        outputFile.WriteLine(lineBuilder.ToString());
                                }
                                else
                                    outputFile.WriteLine(lineBuilder.ToString());
                            }
                            continue;
                        }

                        if (reader.IsEndElement)
                            continue;

                        if (reader.LocalName == "c")
                        {
                            if (mustAddValue)
                                lineBuilder.Append(";");

                            isSharedString = reader.Attributes.Any(a => a.LocalName == "t");
                            //isDate = reader.Attributes.Any(a => a.LocalName == "s" && a.Value == dateIndex.ToString());
                            isDate = reader.Attributes.Any(a => a.LocalName == "s" && countDateForm.Contains(Convert.ToInt32(a.Value)));

                            OpenXmlAttribute codeCellule = reader.Attributes.FirstOrDefault(a => a.LocalName == "r");
                            if (codeCellule != null)
                            {
                                char cellColumn = Regex.Replace(codeCellule.Value, "[0-9]", "").Last();
                                int indexCol = ALPHABET.IndexOf(cellColumn);
                                int indexPreviousCol = ALPHABET.IndexOf(previousColumn);
                                int diff = indexCol - indexPreviousCol;
                                if (diff > 1)
                                {
                                    for (int j = 1; j < diff; j++)
                                        lineBuilder.Append(";");
                                }
                                previousColumn = cellColumn;
                            }
                            mustAddValue = true;
                            continue;
                        }
                        if (reader.LocalName == "f")
                            isCellFormul = true;
                        if (reader.LocalName == "v")
                        {
                            string value = reader.GetText();

                            if (value.StartsWith("#"))
                                lineBuilder.Append(string.Empty);
                            else if (isSharedString && !isCellFormul)
                                lineBuilder.Append(sharedDico[value]);
                            else if (isSharedString && !isDate)
                                lineBuilder.Append(value);
                            else if (isDate && !value.Contains('-') && !value.Contains('+') && !value.Contains('.') && !value.Contains(','))
                                lineBuilder.Append(DateTime.FromOADate(double.Parse(value)).ToShortDateString());
                            else if (value.Contains("E-") || value.Contains("E+"))
                                lineBuilder.Append(decimal.Parse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture));
                            else
                            {
                                if (value.Contains('.') || value.Contains(','))
                                    lineBuilder.Append(double.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture));
                                else
                                    lineBuilder.Append(value);
                            }
                            lineBuilder.Append(";");
                            isCellFormul = false;
                            mustAddValue = false;
                            continue;
                        }
                    }
                }
                outputFile.Close();
                return true;
            }
            catch (OleDbException)
            {
                return false;
            }
        }

        public void SerializeDataSet(string filename, char separator)
        {
            string fileToWrite = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filename), System.IO.Path.GetFileNameWithoutExtension(filename) + ".xml");

            XmlSerializer ser = new XmlSerializer(typeof(DataSet));

            DataSet ds = new DataSet("Fichier");
            DataTable t = new DataTable("Colonne");
            DataColumn c = new DataColumn("Valeur");
            t.Columns.Add(c);
            ds.Tables.Add(t);
            DataRow r;

            foreach (string line in File.ReadLines(filename, Encoding.Default))
            {
                string[] splitedLine = line.Split(separator);

                foreach (string value in splitedLine)
                {
                    r = t.NewRow();
                    r[0] = value;
                    t.Rows.Add(r);
                }
            }

            using (TextWriter writer = File.CreateText(fileToWrite))
            {
                ser.Serialize(writer, ds);
                writer.Close();
            }
        }
    }
}
