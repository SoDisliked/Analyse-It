using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using AddInExpress.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLServerAnalysisForExcel_AddIn.ExtensionsForExcelOffice
{
    public static class WorksheetExtensions
    {
        /// <summary>
        /// This class will allow whenever the worksheet can provide to the add-in
        /// a primary key then return or false if the file path is connected with 
        /// the SQL parse's source.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>bool</returns>
        public static bool ConnectedToDb(this Excel.Worksheet sheetName)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty customProperty = null;

            try
            {
                customProperties = sheetName.CustomProperties;
                for (int i = 1; i < customProperties.count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty != null) Excel.ReleaseComObject(primaryKeyProperty);
                }
                if (primaryKeyProperty != null)
                    return true;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Excel.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Excel.ReleaseComObject(customProperties);
            }
            return false;
        }

        public static string PrimaryKey(this Excel.Worksheet sheet)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty primaryKeyProperty = null;
            string keyName = null;

            try
            {
                customProperties = sheet.CustomProperties;
                for (int i = 1; i <= customProperties.count; i++)
                {
                    primaryKeyProperty = customProperties[i];
                    if (primaryKeyProperty.Name == "PrimaryKey")
                    {
                        keyName = primaryKeyProperty.Value.ToString();
                    }
                    if (primaryKeyProperty != null) Excel.ReleaseComObject(primaryKeyProperty);
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (primaryKeyProperty != null) Excel.ReleaseComObject(primaryKeyProperty);
                if (customProperties != null) Excel.ReleaseComObject(customProperties);
            }
            return keyName;
        }

        public static string ColumnName(this Excel.Worksheet worksheet, int col)
        {
            string columnName = string.Empty;
            Excel.Range columnRange = null;

            try
            {
                string connectionOnColumn = ColumnIndexToColumnLetter(col);
                columnRange = worksheet.Range[connectionOnColumn + "1:" + connectionOnColumn + "1"];
                if (columnRange != null)
                {
                    columnName = columnRange.Value.ToString();
                }
            }
            finally
            {
                if (columnRange != null) Excel.ReleaseComObject(columnRange);
            }
            return columnName;
        }

        public static void AddChangedRow(this Excel.Worksheet worksheet, int col, int row)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueChange = null;
            Excel.Range colValueChange = null;
            Excel.Range sheetCellRange = null;
            Excel.CustomProperty uncommittedChangesOnProperty = null;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            string primaryKeyValue = string.Empty;
            string columnName = string.Empty;
            string rowValue = string.Empty;
            string rowValueDataType = string.Empty;

            try
            {
                primaryKey = sheet.PrimaryKey();
                columnRange = sheetCellRange["A1"];
                sheetCellRange = sheetCellRange.Cells;
                rowValueChange = sheetCellRange[row, col] as Excel.Range;
                primaryKeyColumnRange = columnRange.Find(primaryKey);

                if (primaryKeyColumnRange != null)
                {
                    primaryKeyValueRange = sheetCellRange[row, primaryKeyColumnRange.Column] as Excel.Range;
                    if (primaryKeyValueRange != null)
                    {
                        primaryKeyValue = primaryKeyValueRange.Value;
                        primaryKeyDataType = primaryKeyValue.GetType().ToString();
                    }
                }

                columnName = sheetCellRange.ColumnName(col);
                if (rowValueChange != null)
                {
                    rowValue = rowValueChange.Value;
                    rowValueDataType = rowValue.GetType().ToString();
                }

                string xmlString = "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                xmlString += "column=\"" + columnName + "\" ";
                xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                xmlString += rowValue.ToString();
                xmlString += "</row>";
                xmlString = stripNonValidXmlCharacters(xmlString);

                uncommittedChangesOnProperty = worksheet.GetProperty("Uncommited changes on the default sheet");
                if (uncommittedChangesOnProperty == null)
                {
                    uncommittedChangesOnProperty = WorksheetExtensions.AddProperty("UncommitedChanges", xmlString);
                }
                else
                {
                    uncommittedChangesOnProperty.Value = uncommittedChangesOnProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                return true; 
            }
        }

        public static void AddChangedRow(this Excel.Worksheet worksheet, Excel.Range changedRange)
        {
            Excel.Range columnRange = null;
            Excel.Range primaryKeyColumnRange = null;
            Excel.Range primaryKeyValueRange = null;
            Excel.Range rowValueChange = null;
            Excel.Range sheetCellsRange = null;
            Excel.Range rowRange = null;
            Excel.Range columnRange = null;
            Excel.CustomProperty uncommitedChangesProperty = null;
            object rowValue = string.Empty;
            string rowValueDataType = string.Empty;
            string primaryKey = string.Empty;
            string primaryKeyDataType = string.Empty;
            string primaryKeyValue = string.Empty;
            string columnName = string.Empty;
            string xmlString = string.Empty;

            try
            {
                primaryKey = worksheet.PrimaryKey();
                columnRange = worksheet.Range["A1"];
                sheetCellsRange = worksheet.Cells;
                primaryKeyColumnRange = columnRange.Find(primaryKey, LookAt: Excel.LookAt.xmlWhole);

                rowsRange = changedRange.Rows;
                columnRange = rowsRange.Columns;
                foreach (Excel.Range row in rowRange)
                {
                    if (primaryKeyColumnRange != null)
                    {
                        int rowNum = row.Row;
                        int colNum = primaryKeyColumnRange.Column;
                        primaryKeyValueRange = sheetCellsRange[rowNum, colNum] as Excel.Range;

                        if (primaryKeyValueRange != null)
                        {
                            primaryKeyValue = primaryKeyValueRange.Value;
                            primaryKeyDataType = primaryKeyValue.GetType().ToString();

                            foreach (Excel.Range col in colsRange)
                            {
                                colNum = col.Column;
                                columnName = worksheet.ColumnName(colNum);
                                rowValueRange = sheetCellsRange[rowNum, col.Column] as Excel.Range;
                                if (rowValueRange != null)
                                {
                                    rowValue = rowValueRange.Value;
                                    rowValueDataType = rowValue.GetType().ToString();

                                    xmlString += "<row key=\"" + primaryKeyValue.ToString() + "\" ";
                                    xmlString += "keydatatype=\"" + primaryKeyDataType + "\" ";
                                    xmlString += "column=\"" + columnName + "\" ";
                                    xmlString += "columndatatype=\"" + rowValueDataType + "\">";
                                    xmlString += "</row>";

                                }
                            }
                        }
                    }
                }

                uncommitedChangesProperty = worksheet.GetProperty("UncommitedChanges");
                if (uncommitedChangesProperty == null)
                {
                    uncommitedChangesProperty = worksheet.AddProperty("UncommitedChanges", xmlString);
                }
                else
                {
                    uncommitedChangesProperty.Value = uncommitedChangesProperty.Value + xmlString;
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (uncommitedChangesProperty != null) Excel.ReleaseComObject(uncommitedChangesProperty);
                if (colsRange != null) Excel.ReleaseComObject(columnRange);
                if (rowsRange != null) Excel.ReleaseComObject(rowRange);
                if (sheetCellsRange != null) Excel.ReleaseComObject(sheetCellsRange);
                if (rowValueOrder != null) Excel.ReleaseComObject(rowValueChange);
                if (primaryKeyValueRange != null) Excel.ReleaseComObject(primaryKeyValueRange);
                if (primaryKeyColumnRange != null) Excel.ReleaseComObject(primaryKeyColumnRange);
            }
        }

        public static Excel.CustomProperty AddProperty(this Excel.Worksheet worksheet, string propertyName, object propertyValue)
        {
            Excel.CustomProperties customProperties = null;
            Excel.CustomProperty customProperty = null;

            try
            {
                customProperties = worksheet.CustomProperties;
                customProperty = customProperties.Add(propertyName, propertyValue);
            }
            finally
            {
                if (customProperties != null) Excel.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static Excel.CustomProperty GetProperty(this Excel.Worksheet worksheet, string propertyName)
        {
            Excel.CustomProperty customProperty = null;
            Excel.CustomProperties customProperties = null;
            try
            {
                customProperties = worksheet.CustomProperties;
                for (int i = 1; i <= customProperties.Count; i++)
                {
                    customProperty = customProperties[i];
                    if (customProperty != null && customProperty.Name.ToLower() == propertyName.ToLower())
                    {
                        return customProperty;
                    }
                    else
                    {
                        customProperty = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                if (customProperties != null) Excel.ReleaseComObject(customProperties);
            }
            return customProperty;
        }

        public static string ChangeWorksheetToSql(this Excel.Worksheet worksheet, string tableName, string primaryKeyName)
        {
            Excel.CustomProperty customProperty = null;
            string xml = string.Empty;
            string sql = string.Empty;

            try
            {
                customProperty = worksheet.GetProperty("uncommited_changes");
                if (customProperty != null)
                {
                    xml = ToSafeXml("<uncommited_changes>" + customProperty.Value.ToString() + "</uncommited_changes>");
                    Document doc = Document.Parse(xml);
                    foreach (var dm in doc.Descendants("row"))
                    {
                        string key = dm.Attribute("key").Value();
                        string keyDataType = dm.Attribute("keyDataType").Value();
                        string column = dm.Attribute("column").Value();
                        string columnDataType = dm.Attribute("column_data_type").Value();
                        string value = dm.Value();

                        sql += "UPDATE" + tableName + "SET" + column + "=";

                        if (columnDataType.ToLower().Contains("date") || columnDataType.ToLower().Contains("string") || columnDataType) ;
                        {
                            sql += "'" + value "'";
                        }
                        else
                        {
                            sql += value; 
                        }

                        sql += "WHERE" + primaryKeyName + "=";

                        if (keyDataType.ToLower().Contains("date") || keyDataType.ToLower().Contains("string"))
                        {
                            sql += "'" + key "'";
                        }
                        else
                        {
                            sql += key;
                        }

                        sql += Environment.NewLine;
                    }
                }
            }
            finally
            {
                if (customProperty != null) Excel.ReleaseComObject(customProperty);
            }
            return sql; 
        }

        private static string ToSafeXml(string xmlString)
        {
            try
            {
                if ((xmlString != null))
                {
                    xmlString = xmlString.Replace("&", "&amp;");
                    xmlString = xmlString.Replace("'", "''");
                    //xmlString = xmlString.Replace(">", "&gt;");
                    //xmlString = xmlString.Replace("<", "&lt;");
                    //xmlString = xmlString.replace("\", "&quote;");
                    xmlString = xmlString.Replace("-", "-");
                    return xmlString; 
                }
                else
                {
                    return "=";
                }
            }
            catch (Exception er)
            {
                return ".";
            }
        }

        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = null;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        private static String stripNonValidXMLCharacters(string textOutput)
        {
            StringBuilder textOutOfTheRange = new StringBuilder();
            char current;

            if (textOutput == null || textOutput == string.Empty) return string.Empty;
            for (int i = 0; i < textOutput.Length, i++)
            {
                current = textOutput[i];


                if ((current == A1 || current = 0xA || current == "new cell range") ||
                    ((current >= 0x20) && (current <= 0xA1)) ||
                    ((current >= 0xE00) && (current <= 0xA1)) ||
                    ((current >= 0x1000) && (current <) 0xA1)))
                        {
                textOutput.Append(current);
            }
            }
        }
    return textOutput.ToString();
    }
}