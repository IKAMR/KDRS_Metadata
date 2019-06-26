using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace KDRS_Metadata
{
    class DataConverter
    {
        public int totalTableCount;
        public string schemaName;


        public void Convert(string filename, bool includeTables)
        {


            Application xlApp1 = new Application();
            Workbooks xlWorkbooks = xlApp1.Workbooks;

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(filename);
            XmlNode root = xmldoc.DocumentElement;
            var nsmgr = new XmlNamespaceManager(xmldoc.NameTable);
            var nameSpace = xmldoc.DocumentElement.NamespaceURI;

            nsmgr.AddNamespace("siard", nameSpace);
            //nsmgr.AddNamespace("siard", "http://www.bar.admin.ch/xmlns/siard/2.0/metadata.xsd");

            Workbook xlWorkBook;

            Sheets xlWorkSheets;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlWorkbooks.Add(misValue);

            xlWorkSheets = xlWorkBook.Sheets;

            Worksheet DBWorkSheet = xlWorkSheets.get_Item(1);
            AddDBInfo(DBWorkSheet, root, nsmgr);
            Marshal.ReleaseComObject(DBWorkSheet);

            XmlNode schemas = root.SelectSingleNode("//siard:schemas", nsmgr);
            XmlNode tables = root.SelectSingleNode("//siard:tables", nsmgr);

            totalTableCount = 0;

            foreach (XmlNode schema in schemas.ChildNodes)
            {
                schemaName = root.SelectSingleNode("//siard:name", nsmgr).InnerText;

                Worksheet tableOverviewWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);
                AddTableOverview(tableOverviewWorksheet, tables);
                Marshal.ReleaseComObject(tableOverviewWorksheet);

                if (includeTables)
                {
                    foreach (XmlNode table in tables.ChildNodes)
                    {
                        //Console.WriteLine("Adding table: " + table.SelectSingleNode("siard:foreignKeys/siard:foreignKey/siard:name", nsmgr).InnerText);
                        //Console.WriteLine("Adding table: " + table["name"].InnerText);

                        Worksheet tableWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);

                        AddTable(tableWorksheet, table, nsmgr);
                        totalTableCount++;
                    }
                }
            }

            //antTables = "Number of tables " + tables.ChildNodes.Count.ToString();

            xlWorkBook.Sheets[1].Select();

            string excelFileName;
            if (includeTables)
            {
                string origName = Path.GetFileNameWithoutExtension(filename);
                string folder = Directory.GetParent(Path.GetFullPath(filename)).ToString();
                excelFileName = Path.Combine(folder, origName + "_tables.xlsx");
                Console.WriteLine(excelFileName);
            }
            else
            {
                excelFileName = Path.ChangeExtension(Path.GetFullPath(filename), ".xlsx");
            }

            xlWorkBook.SaveAs(excelFileName);

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
            DBWorkSheet.Name = "DB";

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
            /*
            foreach (string s in fieldNames)
            {
                DBWorkSheet.Cells[cnt, 1] = s;
                DBWorkSheet.Cells[cnt, 2] = getNodeText(table, "//siard:" + s, nsmgr);
                cnt++;
            }*/

            // tooolname
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

            //tableCount
            DBWorkSheet.Cells[cnt, 1] = fieldNames[7];
            DBWorkSheet.Cells[cnt, 2] = table.SelectSingleNode("//siard:tables", nsmgr).ChildNodes.Count;
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
                DBWorkSheet.Cells[cnt, 2] = getNodeText(table, "//siard:" + field, nsmgr);
                cnt++;
            }

            //digestType
            DBWorkSheet.Cells[cnt, 1] = fieldNames[20];
            DBWorkSheet.Cells[cnt, 2] = table.Attributes["version"].Value;
            cnt++;

            //digest
            DBWorkSheet.Cells[cnt, 1] = fieldNames[21];
            DBWorkSheet.Cells[cnt, 2] = table.Attributes["version"].Value;
            cnt++;

            //clientMachine
            DBWorkSheet.Cells[cnt, 1] = fieldNames[22];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(getNodeText(table, "//siard:" + fieldNames[22], nsmgr));
            cnt++;

            //databaseProduct
            DBWorkSheet.Cells[cnt, 1] = fieldNames[23];
            DBWorkSheet.Cells[cnt, 2] = getNodeText(table, "//siard:" + fieldNames[23], nsmgr);
            cnt++;

            //connection
            DBWorkSheet.Cells[cnt, 1] = fieldNames[24];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(getNodeText(table, "//siard:" + fieldNames[24], nsmgr));
            cnt++;

            //databaseUser
            DBWorkSheet.Cells[cnt, 1] = fieldNames[25];
            DBWorkSheet.Cells[cnt, 2] = SensitiveString(getNodeText(table, "//siard:" + fieldNames[25], nsmgr));
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "schemas";
            XmlNodeList schemas = table.SelectNodes("//siard:schemas", nsmgr);

            string schemasList = getNodeText(schemas[0], "descendant::siard:folder", nsmgr);
            for (int i=1; i<schemas.Count; i++)
            {
                schemasList += ", " + getNodeText(schemas[i], "descendant::siard:folder", nsmgr);
                
            }
            DBWorkSheet.Cells[cnt, 2] = schemasList;
            cnt++;

            DBWorkSheet.Cells[cnt, 1] = "users";
            XmlNode users = table.SelectSingleNode("//siard:users", nsmgr);

            foreach (XmlNode user in users.ChildNodes)
            {
                DBWorkSheet.Cells[cnt, 2] = getNodeText(user, "descendant::siard:name", nsmgr);
                cnt++;
            }

            DBWorkSheet.Columns.AutoFit();
            Marshal.ReleaseComObject(DBWorkSheet);
        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, XmlNode table, XmlNamespaceManager nsmgr)
        {
            tableWorksheet.Name = GetNumbers(table["folder"].InnerText);

            Range c1 = tableWorksheet.Cells[1, 1];
            Range c2 = tableWorksheet.Cells[1, 1];
            Range linkCell = tableWorksheet.get_Range(c1, c2);

            Hyperlinks links = tableWorksheet.Hyperlinks;

            links.Add(linkCell, "", "DB!A1", "", "<<< home");

            int cellCount = 3;

            List<string> columnNames = new List<string>()
            {
                "column",
                "name",
                "type",
                "type original",
                "nullable",
                "default value",
                "description",
                "note"
            };

            foreach (string name in columnNames)
            {
                tableWorksheet.Cells[2, columnNames.IndexOf(name) + 1] = name;
            }
            //------------------------------------------------------------------------

            // Finds the metadata for each table and prints to Excel.

            string table_description = getNodeText(table, "descendant::siard:description", nsmgr);

            string primaryKey_name = getNodeText(table["primaryKey"], "descendant::siard:name", nsmgr);
            string primaryKey_column = getNodeText(table["primaryKey"], "descendant::siard:column", nsmgr);

            string[][] rowNamesArray = new string[9][] {
                new string[2] { "schemaName", table.ParentNode.ParentNode["name"].InnerText.ToString() },
                new string[2] { "schemaFolder", table.ParentNode.ParentNode["folder"].InnerText.ToString()},
                new string[2] { "tableName", table["name"].InnerText.ToString() },
                new string[2] { "tableFolder", table["folder"].InnerText.ToString() },
                new string[2] { "tableDescription", table_description },
                new string[2] { "rows", table["rows"].InnerText.ToString() },
                new string[2] { "column", table["columns"].ChildNodes.Count.ToString() },
                new string[2] { "pkName", primaryKey_name },
                new string[2] { "pkColumn", primaryKey_column },
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
                int foreignKeys_count = 0;
                foreach (XmlNode fKey in foreignKeys.ChildNodes)
                {
                    string foreignKeys_name = getNodeText(fKey, "descendant::siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkName " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_name;
                    cellCount++;

                    string foreignKeys_description = getNodeText(fKey, "descendant::siard:description", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkDescription " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_description;
                    cellCount++;

                    string foreignKeys_column = getNodeText(fKey, "descendant::siard:reference/siard:colum", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkColumn " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_column;
                    cellCount++;

                    string foreignKeys_ref_schema = getNodeText(fKey, "descendant::siard:referencedSchema", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkRefSchema " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_ref_schema;
                    cellCount++;

                    string foreignKeys_table = getNodeText(fKey, "descendant::siard:referencedTable", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkRefTable " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_table;
                    cellCount++;

                    string foreignKeys_ref_col = getNodeText(fKey, "descendant::siard:reference/siard:referenced", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkRefColumn0 " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_ref_col;
                    cellCount++;

                    string foreignKeys_delete_action = getNodeText(fKey, "descendant::siard:deleteAction", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkDeleteAction " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_delete_action;
                    cellCount++;

                    string foreignKeys_update_action = getNodeText(fKey, "descendant::siard:updateAction", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkUpdateAction " + foreignKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_update_action;
                    cellCount++;

                    foreignKeys_count++;
                }
            }

            //-------------------------------------------------------------------------------------
            // Finds all candidate keys in table and prints to Excel.
            XmlNode candidateKeys = table.SelectSingleNode("descendant::siard:candidateKeys", nsmgr);

            if (candidateKeys != null)
            {
                int candidateKeys_count = 0;
                foreach (XmlNode cKey in candidateKeys.ChildNodes)
                {
                    string candidateKeys_name = getNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckName " + candidateKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_name;
                    cellCount++;

                    string candidateKeys_description = getNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:description", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckDescription " + candidateKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_description;
                    cellCount++;

                    string candidateKeys_column1 = getNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:column[1]", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckColumn0 " + candidateKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_column1;
                    cellCount++;

                    string candidateKeys_column2 = getNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:column[2]", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckColumn1 " + candidateKeys_count;
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_column2;
                    cellCount++;

                    candidateKeys_count++;
                }
            }

            //-------------------------------------------------------------------------------------
            // Finds all columns in table and prints info to Excel.
            XmlNode tableColumns = table.SelectSingleNode("descendant::siard:columns", nsmgr);

            int column_count = 0;
            foreach (XmlNode column in tableColumns.ChildNodes)
            {
                tableWorksheet.Cells[cellCount, 1] = "Column" + column_count;
                column_count++;

                string col_name = getNodeText(column, "descendant::siard:name", nsmgr);
                tableWorksheet.Cells[cellCount, 2] = col_name;

                string col_type = getNodeText(column, "descendant::siard:type", nsmgr);
                tableWorksheet.Cells[cellCount, 3] = col_type;

                string col_type_original = getNodeText(column, "descendant::siard:typeOriginal", nsmgr);
                tableWorksheet.Cells[cellCount, 4] = col_type_original;

                string col_nullable = getNodeText(column, "descendant::siard:nullable", nsmgr);
                tableWorksheet.Cells[cellCount, 5] = col_nullable;

                string col_defaultValue = getNodeText(column, "descendant::siard:defaultValue", nsmgr);
                tableWorksheet.Cells[cellCount, 6] = col_defaultValue;

                string col_description = getNodeText(column, "descendant::siard:description", nsmgr);
                tableWorksheet.Cells[cellCount, 7] = col_description;

                cellCount++;
            }

            tableWorksheet.Columns.AutoFit();
            Marshal.ReleaseComObject(tableWorksheet);
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns Innertext of node found in table with query.
        private string getNodeText(XmlNode table, string query, XmlNamespaceManager nsmgr)
        {
            string varName = "[NA]";
            if (table != null)
            {
                XmlNode node = table.SelectSingleNode(query, nsmgr);
                if (node != null)
                {
                    varName = node.InnerText;
                }
            }
            return varName;
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a worksheet with table overview
        private void AddTableOverview(Worksheet tableOverviewWorksheet, XmlNode tables)
        {
            tableOverviewWorksheet.Name = "Tables";

            List<string> columnNames = new List<string>()
            {
                "table",
                "folder",
                "schema",
                "rows",
                "priority",
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
            foreach (XmlNode table in tables.ChildNodes)
            {

                int countObj = 1;
                foreach (string on in objectNames)
                {
                    try
                    {
                        
                     //   tableOverviewWorksheet.Cells[count, countObj] = table[on].InnerText;
                    }
                    catch (Exception ex)
                    {
                       // tableOverviewWorksheet.Cells[count, countObj] = "-";
                    }
                    countObj++;
                }

                string name = table["name"].InnerText;
                string folder = GetNumbers(table["folder"].InnerText);

                Range c1 = tableOverviewWorksheet.Cells[count, 1];
                Range c2 = tableOverviewWorksheet.Cells[count, 1];
                Range linkCell = tableOverviewWorksheet.get_Range(c1, c2);

                Hyperlinks links = tableOverviewWorksheet.Hyperlinks;

                links.Add(linkCell, "", folder + "!A1", "", name);

                tableOverviewWorksheet.Cells[count, 2] = table["folder"].InnerText;
                tableOverviewWorksheet.Cells[count, 3] = table.ParentNode.ParentNode["folder"].InnerText;
                tableOverviewWorksheet.Cells[count, 4] = table["rows"].InnerText;
                count++;

                Marshal.ReleaseComObject(c1);
                Marshal.ReleaseComObject(c2);
                Marshal.ReleaseComObject(linkCell);
                Marshal.ReleaseComObject(links);
            }

            tableOverviewWorksheet.Columns.AutoFit();
            Marshal.ReleaseComObject(tableOverviewWorksheet);
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
    }
    //==========================================================================================================
}
