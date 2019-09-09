using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;

namespace KDRS_Metadata
{
    class DataConverter
    {
        public int totalTableCount;
        public int totalSchemaCount;
        public List<Schema> schemaNames = new List<Schema>();
        public string excelFileName;

        public delegate void ProgressUpdate(int count, int totalCount);
        public event ProgressUpdate OnProgressUpdate;

        public void Convert(string filename, bool includeTables)
        {
            schemaNames.Clear();

            Application xlApp1 = new Application
            {
                DecimalSeparator = ".",
                UseSystemSeparators = false
            };
            Workbooks xlWorkbooks = xlApp1.Workbooks;

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(filename);
            XmlNode root = xmldoc.DocumentElement;
            var nsmgr = new XmlNamespaceManager(xmldoc.NameTable);
            var nameSpace = xmldoc.DocumentElement.NamespaceURI;

            nsmgr.AddNamespace("siard", nameSpace);

            Workbook xlWorkBook;
            Sheets xlWorkSheets;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlWorkbooks.Add(misValue);
            xlWorkSheets = xlWorkBook.Sheets;

            Worksheet DBWorkSheet = xlWorkSheets.get_Item(1);
            AddDBInfo(DBWorkSheet, root, nsmgr);
            Marshal.ReleaseComObject(DBWorkSheet);

            XmlNodeList schemas = root.SelectNodes("descendant::siard:schema", nsmgr);
            totalSchemaCount = schemas.Count;
            XmlNodeList allTables = root.SelectNodes("//siard:tables/siard:table", nsmgr);

            Console.WriteLine("Schemas read");

            int tableCount = 0;
            totalTableCount = allTables.Count;

            Worksheet tableOverviewWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);
            AddTableOverview(tableOverviewWorksheet, schemas, nsmgr, includeTables);
            Marshal.ReleaseComObject(tableOverviewWorksheet);

            Console.WriteLine("Added tableoverview");

            foreach (XmlNode schema in schemas)
            {
                schemaNames.Add(new Schema(getInnerText(schema["name"]),getInnerText(schema["folder"])));
                XmlNode tables = schema.SelectSingleNode("descendant::siard:tables", nsmgr);
                Console.WriteLine("Enter schema");

                if (includeTables)
                {
                    foreach (XmlNode table in tables.ChildNodes)
                    {
                        Worksheet tableWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);

                        AddTable(tableWorksheet, table, nsmgr);
                        tableCount++;
                        Console.WriteLine("Added table");

                        OnProgressUpdate?.Invoke(tableCount, totalTableCount);
                    }
                }
            }

            xlWorkBook.Sheets[1].Select();

            int fileCounter = 1;

           if (includeTables)
            {
                excelFileName = Path.ChangeExtension(Path.GetFullPath(filename), ".xlsx");
            }
            else
            {
                string origName = Path.GetFileNameWithoutExtension(filename);
                string folder = Directory.GetParent(Path.GetFullPath(filename)).ToString();
                excelFileName = Path.Combine(folder, origName + "_tablelist.xlsx");
                Console.WriteLine(excelFileName);
            }

            string checkedFileNAme = checkFileName(excelFileName, fileCounter);
            //Console.WriteLine("Outfile: " + checkedFileNAme);
                       
            xlWorkBook.SaveAs(checkedFileNAme);

            xlApp1.UseSystemSeparators = true;

            xlWorkBook.Close();
            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlApp1);

        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a worksheet with information about the database.
        private void AddDBInfo(Worksheet DBWorkSheet, XmlNode table, XmlNamespaceManager nsmgr)
        {
            DBWorkSheet.Name = "db";

            List<string> fieldNames = new List<string>()
            {
                "toolName",
                "toolVersion",
                "systemSupplier",
                "systemId",
                "systemName",
                "systemVersion",
                "systemInstance",
                "tableCount",
                "",
                "SIARD",
                "version",
                "dbname",
                "description",
                "archiver",
                "archiverContact",
                "dataOwner",
                "dataOriginTimespan",
                "lobFolder",
                "producerApplication",
                "archivalDate",
                "digestType",
                "digest",
                "clientMachine",
                "databaseProduct",
                "connection",
                "databaseUser",
                "schemas"
            };

            int cnt = 1;

            // toolname
            DBWorkSheet.Cells[cnt, 1] = fieldNames[0];
            DBWorkSheet.Cells[cnt, 2] = Globals.toolName;
            cnt++;

            // toolVersion
            DBWorkSheet.Cells[cnt, 1] = fieldNames[1];
            DBWorkSheet.Cells[cnt, 2] = Globals.toolVersion;
            cnt++;

            for (int i=2; i<7; i++)
            {
                DBWorkSheet.Cells[cnt, 1] = fieldNames[i];
                DBWorkSheet.Cells[cnt, 2] = "";
                cnt++;
            }

            Range redColorRng = DBWorkSheet.Range["A3", "C6"];
            redColorRng.Characters.Font.Color = Color.Red;

            Range orangeColorRng = DBWorkSheet.Range["A7", "C7"];
            orangeColorRng.Characters.Font.Color = Color.Orange;

            //tableCount
            DBWorkSheet.Cells[cnt, 1] = fieldNames[7];
            XmlNodeList  tableCount = table.SelectNodes("//siard:table", nsmgr);
            DBWorkSheet.Cells[cnt, 2] = tableCount.Count;
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = fieldNames[8];
            DBWorkSheet.Cells[cnt, 2] = "";
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = fieldNames[9];
            DBWorkSheet.Cells[cnt, 2] = "metadata.xml";
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = fieldNames[10];
            DBWorkSheet.Cells[cnt, 2] = table.Attributes["version"].Value;
            cnt++;

            for (int i = 11; i < 20; i++)
            {
                string field = fieldNames[i];
                DBWorkSheet.Cells[cnt, 1] = field;
                DBWorkSheet.Cells[cnt, 2] = GetNodeText(table, "//siard:" + field, nsmgr);
                cnt++;
            }

            //digestType
            DBWorkSheet.Cells[cnt, 1] = fieldNames[20];
            DBWorkSheet.Cells[cnt, 2] = "";
            cnt++;

            //digest
            DBWorkSheet.Cells[cnt, 1] = fieldNames[21];
            DBWorkSheet.Cells[cnt, 2] = "";
            cnt++;

            //clientMachine
            DBWorkSheet.Cells[cnt, 1] = fieldNames[22];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[22], nsmgr));
            cnt++;

            //databaseProduct
            DBWorkSheet.Cells[cnt, 1] = fieldNames[23];
            DBWorkSheet.Cells[cnt, 2] = GetNodeText(table, "//siard:" + fieldNames[23], nsmgr);
            cnt++;

            //connection
            DBWorkSheet.Cells[cnt, 1] = fieldNames[24];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[24], nsmgr));
            cnt++;

            //databaseUser
            DBWorkSheet.Cells[cnt, 1] = fieldNames[25];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[25], nsmgr));
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "schemas";
            XmlNodeList schemas = table.SelectNodes("//siard:schemas/siard:schema", nsmgr);

            string schemasList = GetNodeText(schemas[0], "descendant::siard:folder", nsmgr);
            for (int i=1; i<schemas.Count; i++)
            {
                schemasList += ", " + GetNodeText(schemas[i], "descendant::siard:folder", nsmgr);
            }
            DBWorkSheet.Cells[cnt, 2] = schemasList;
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "users";
            XmlNode users = table.SelectSingleNode("//siard:users", nsmgr);
            DBWorkSheet.Cells[cnt, 2] = getChildCount(users);
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "roles";
            XmlNode roles = table.SelectSingleNode("//siard:roles", nsmgr);
            DBWorkSheet.Cells[cnt, 2] = getChildCount(roles);
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "privileges";
            XmlNode privileges = table.SelectSingleNode("//siard:privileges", nsmgr);
            DBWorkSheet.Cells[cnt, 2] = getChildCount(privileges);

            DBWorkSheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            DBWorkSheet.Columns.AutoFit();
            Marshal.ReleaseComObject(DBWorkSheet);
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Creates a worksheet with table overview
        private void AddTableOverview(Worksheet tableOverviewWorksheet, XmlNodeList schemas, XmlNamespaceManager nsmgr, bool includeTables)
        {
            tableOverviewWorksheet.Name = "tables";

            List<string> columnNames = new List<string>()
            {
                "table",
                "folder",
                "schema",
                "rows",
                "columns",
                "priority",
                "pri-sort",
                "entity",
                "description",
                "note"
            };

            foreach (string name in columnNames)
            {
                tableOverviewWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
            }

            List<string> objectNames = new List<string>()
            {
                "name",
                "folder",
                "schema",
                "rows"
            };

            int count = 2;
            foreach (XmlNode schema in schemas)
            {
                XmlNode tables = schema.SelectSingleNode("descendant::siard:tables", nsmgr);
                string schemaNumber = GetNumbers(schema["folder"].InnerText);

                foreach (XmlNode table in tables.ChildNodes)
                {
                    string name = table["name"].InnerText;
                    string folder = GetNumbers(table["folder"].InnerText);
                    if (includeTables)
                    {
                        Range c1 = tableOverviewWorksheet.Cells[count, 1];
                        Range c2 = tableOverviewWorksheet.Cells[count, 1];
                        Range linkCell = tableOverviewWorksheet.get_Range(c1, c2);

                        Hyperlinks links = tableOverviewWorksheet.Hyperlinks;

                        if (totalSchemaCount < 2)
                            links.Add(linkCell, "", folder + "!A1", "", name);
                        else
                            links.Add(linkCell, "", schemaNumber + "." + folder + "!A1", "", name);

                        Marshal.ReleaseComObject(c1);
                        Marshal.ReleaseComObject(c2);
                        Marshal.ReleaseComObject(linkCell);
                        Marshal.ReleaseComObject(links);
                    }
                    else
                    {
                        tableOverviewWorksheet.Cells[count, 1] = name;
                    }

                    tableOverviewWorksheet.Cells[count, 2] = getInnerText(table["folder"]);
                    tableOverviewWorksheet.Cells[count, 3] = table.ParentNode.ParentNode["folder"].InnerText;

                    string tableRows = getInnerText(table["rows"]);
                    tableOverviewWorksheet.Cells[count, 4] = tableRows;

                    string tableColumns = getChildCount(table["columns"]);
                    tableOverviewWorksheet.Cells[count, 5] = tableColumns;

                    string tablePriority = "";
                    if (tableRows == "0")
                        tablePriority = "EMPTY";
                    else
                        tablePriority = GetNodeTxtEmpty(table, "descendant::siard:priority", nsmgr);

                    tableOverviewWorksheet.Cells[count, 6] = tablePriority;

                    int tablePriSort = Globals.PriSort(tablePriority);
                    tableOverviewWorksheet.Cells[count, 7] = tablePriSort;
                    count++;
                }
            }

            Range range = tableOverviewWorksheet.Cells[2, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            tableOverviewWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableOverviewWorksheet.Columns.AutoFit();

            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Creates a worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, XmlNode table, XmlNamespaceManager nsmgr)
        {
            string schemaNumber = GetNumbers(table.ParentNode.ParentNode["folder"].InnerText);

            if (totalSchemaCount < 2)
                tableWorksheet.Name = GetNumbers(table["folder"].InnerText);
            else 
                tableWorksheet.Name = schemaNumber + "." + GetNumbers(table["folder"].InnerText);

            Range c1 = tableWorksheet.Cells[1, 1];
            Range c2 = tableWorksheet.Cells[1, 1];
            Range linkCell = tableWorksheet.get_Range(c1, c2);

            Hyperlinks links = tableWorksheet.Hyperlinks;

            links.Add(linkCell, "", "tables!A1", "", "column <<< tables");

            int cellCount = 2;

            List<string> columnNames = new List<string>()
            {
                "column",
                "name",
                "type",
                "typeOriginal",
                "nullable",
                "defaultValue",
                "lobFolder",
                "entity",
                "description",
                "note"
            };

            foreach (string name in columnNames.Skip(1))
            {
                tableWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
            }
            //------------------------------------------------------------------------
            // Finds the metadata for each table and prints to Excel.

            string table_description = GetNodeTxtEmpty(table, "descendant::siard:description", nsmgr);

            string primaryKey_name = GetNodeText(table["primaryKey"], "descendant::siard:name", nsmgr);
            string primaryKey_column = GetNodeText(table["primaryKey"], "descendant::siard:column", nsmgr);

            string tableRows = getInnerText(table["rows"]);

            string tablePriority = GetNodeText(table, "descendant::siard:priority", nsmgr);
            if (tableRows == "0")
                tablePriority = "EMPTY";

            string tableEntity = GetNodeTxtEmpty(table, "descendant::siard:entity", nsmgr);

            string[][] rowNamesArray = new string[12][] {
                new string[2] { "schemaName", table.ParentNode.ParentNode["name"].InnerText.ToString() },
                new string[2] { "schemaFolder", table.ParentNode.ParentNode["folder"].InnerText.ToString()},
                new string[2] { "tableName", table["name"].InnerText.ToString() },
                new string[2] { "tableFolder", getInnerText(table["folder"]) },
                new string[2] { "tablePriority", tablePriority },
                new string[2] { "tableEntity", tableEntity },
                new string[2] { "tableDescription", table_description },
                new string[2] { "rows", tableRows },
                new string[2] { "columns", getChildCount(table["columns"]) },
                new string[2] { "pkName", primaryKey_name },
                new string[2] { "pkColumn", primaryKey_column },
                new string[2] { "pkDescription", GetNodeText(table["primaryKey"], "descendant::siard:description", nsmgr) }
            };

            foreach (string[] rn in rowNamesArray)
            {
                tableWorksheet.Cells[cellCount, 1] = rn;
                tableWorksheet.Cells[cellCount, 2] = rn[1];

                cellCount++;
            }
            //-------------------------------------------------------------------------------------
            // Finds all foreign keys in table and prints to Excel.
            XmlNode foreignKeys = table.SelectSingleNode("descendant::siard:foreignKeys", nsmgr);

            if (foreignKeys != null)
            {
                foreach (XmlNode fKey in foreignKeys.ChildNodes)
                {
                    string foreignKeys_name = GetNodeText(fKey, "descendant::siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkName";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_name;
                    cellCount++;

                    string foreignKeys_ref_schema = GetNodeText(fKey, "descendant::siard:referencedSchema", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkRefSchema";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_ref_schema;
                    cellCount++;

                    string foreignKeys_table = GetNodeText(fKey, "descendant::siard:referencedTable", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkRefTable";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_table;
                    cellCount++;
                    
                    XmlNodeList reference = fKey.SelectNodes("descendant::siard:reference", nsmgr);
                    if (reference != null)
                    {
                        foreach (XmlNode refer in reference)
                        {
                            string foreignKeys_column = GetNodeText(refer, "descendant::siard:column", nsmgr);
                            tableWorksheet.Cells[cellCount, 1] = "fkColumn";
                            tableWorksheet.Cells[cellCount, 2] = foreignKeys_column;
                            cellCount++;

                            string foreignKeys_ref_col = GetNodeText(refer, "descendant::siard:referenced", nsmgr);
                            tableWorksheet.Cells[cellCount, 1] = "referenced";
                            tableWorksheet.Cells[cellCount, 2] = foreignKeys_ref_col;
                            cellCount++;
                        }
                    }

                    string foreignKeys_description = GetNodeTxtEmpty(fKey, "descendant::siard:description", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkDescription";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_description;
                    cellCount++;

                    string foreignKeys_delete_action = GetNodeText(fKey, "descendant::siard:deleteAction", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkDeleteAction";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_delete_action;
                    cellCount++;

                    string foreignKeys_update_action = GetNodeText(fKey, "descendant::siard:updateAction", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkUpdateAction";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_update_action;
                    cellCount++;
                }
            }
            //-------------------------------------------------------------------------------------
            // Finds all candidate keys in table and prints to Excel.
            XmlNode candidateKeys = table.SelectSingleNode("descendant::siard:candidateKeys", nsmgr);

            if (candidateKeys != null)
            {
                foreach (XmlNode cKey in candidateKeys.ChildNodes)
                {
                    string candidateKeys_name = GetNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckName ";
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_name;
                    cellCount++;

                    string candidateKeys_description = GetNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:description", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckDescription ";
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_description;
                    cellCount++;

                    for (int i=1; i<cKey.ChildNodes.Count; i++)
                    {
                        string candidateKeys_column1 = GetNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:column[" + i + "]", nsmgr);
                        tableWorksheet.Cells[cellCount, 1] = "ckColumn";
                        tableWorksheet.Cells[cellCount, 2] = candidateKeys_column1;
                        cellCount++;
                    }
                }
            }
            //-------------------------------------------------------------------------------------
            // Finds all columns in table and prints info to Excel.
            XmlNode tableColumns = table.SelectSingleNode("descendant::siard:columns", nsmgr);

            int column_count = 1;
            if (tableColumns != null)
            {
                foreach (XmlNode column in tableColumns.ChildNodes)
                {
                    tableWorksheet.Cells[cellCount, 1] = column_count;
                    column_count++;

                    string col_name = GetNodeText(column, "descendant::siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 2] = col_name;

                    string col_type = GetNodeText(column, "descendant::siard:type", nsmgr);
                    tableWorksheet.Cells[cellCount, 3] = col_type;

                    string col_type_original = GetNodeText(column, "descendant::siard:typeOriginal", nsmgr);
                    tableWorksheet.Cells[cellCount, 4] = col_type_original;

                    string col_nullable = GetNodeText(column, "descendant::siard:nullable", nsmgr);
                    tableWorksheet.Cells[cellCount, 5] = col_nullable;

                    string col_defaultValue = GetNodeText(column, "descendant::siard:defaultValue", nsmgr);
                    tableWorksheet.Cells[cellCount, 6] = col_defaultValue;

                    string col_lobFolder = GetNodeText(column, "descendant::siard:lobFolder", nsmgr);
                    tableWorksheet.Cells[cellCount, 7] = col_lobFolder;

                    string col_entity = GetNodeText(column, "descendant::siard:entity", nsmgr);
                    tableWorksheet.Cells[cellCount, 8] = col_lobFolder;

                    string col_description = GetNodeTxtEmpty(column, "descendant::siard:description", nsmgr);
                    tableWorksheet.Cells[cellCount, 9] = col_description;

                    cellCount++;
                }
            }

            Range range = tableWorksheet.Cells[5, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            tableWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableWorksheet.Columns.AutoFit();

            Marshal.ReleaseComObject(tableWorksheet);
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns Innertext of node found in table with query.
        private string GetNodeText(XmlNode table, string query, XmlNamespaceManager nsmgr)
        {
            string varName = "[NA]";
            if (table != null)
            {
                XmlNode node = table.SelectSingleNode(query, nsmgr);
                if (node != null)
                {
                    varName = node.InnerText;
                    if (varName == "" && query != "descendant::siard:deleteAction" && query != "descendant::siard:updateAction")
                        varName = "[EMPTY]";
                }
            }
            return varName;
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns Innertext of node found in table with query. If no text return empty string.
        private string GetNodeTxtEmpty(XmlNode table, string query, XmlNamespaceManager nsmgr)
        {
            string text = "";
            if (table != null)
            {
                XmlNode node = table.SelectSingleNode(query, nsmgr);
                if (node != null)
                {
                    text = node.InnerText;
                }
            }

                    return text;
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns Innertext of node.
        private string getInnerText(XmlNode table)
        {
            string varName = "[NA]";
            if (table != null)
            {
                varName = table.InnerText;
                if (varName == "")
                    varName = "[EMPTY]";
            }
            return varName;
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns the children count of table.
        private string getChildCount(XmlNode table)
        {
            string varName = "0";
            if (table != null)
            {
                varName = table.ChildNodes.Count.ToString();
            }
            return varName;
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private static string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        private static string SensitiveString(string input)
        {
            if (input == "[NA]")
                return input;
            else
                return "*****";
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private string checkFileName(string fileName, int fileCounter)
        {

            string origName = Path.GetFileNameWithoutExtension(fileName);
            string folder = Directory.GetParent(Path.GetFullPath(fileName)).ToString();
            
            string extension = Path.GetExtension(fileName);
            while (File.Exists(fileName))
            {
                fileName = Path.Combine(folder, origName + "_" + fileCounter + extension);
                fileCounter++;
               // checkFileName(fileName, fileCounter);
            }

            return fileName;
        }
    }
    //==========================================================================================================
}
