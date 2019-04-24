using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Metadata_XLS
{
    class DataConverter
    {
        public string antTables;
        public string schemaName;

        public void Convert(string filename)
        {

            Application xlApp1 = new Application();

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(filename);
            XmlNode root = xmldoc.DocumentElement;
            var nsmgr = new XmlNamespaceManager(xmldoc.NameTable);
            var nameSpace = xmldoc.DocumentElement.NamespaceURI;

            nsmgr.AddNamespace("siard", nameSpace);
            //nsmgr.AddNamespace("siard", "http://www.bar.admin.ch/xmlns/siard/2.0/metadata.xsd");

            Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp1.Workbooks.Add(misValue);

            AddDBInfo(xlWorkBook, root, nsmgr);


            XmlNode schemas = root.SelectSingleNode("//siard:schemas", nsmgr);
            XmlNode tables = root.SelectSingleNode("//siard:tables", nsmgr);

            foreach (XmlNode schema in schemas.ChildNodes)
            {
                schemaName = root.SelectSingleNode("//siard:name", nsmgr).InnerText;
                //Console.WriteLine("Adding table overview: " + tables);
                AddTableOverview(xlApp1, xlWorkBook, tables);

                foreach (XmlNode table in tables.ChildNodes)
                {
                    //Console.WriteLine("Adding table: " + table.SelectSingleNode("siard:foreignKeys/siard:foreignKey/siard:name", nsmgr).InnerText);
                    //Console.WriteLine("Adding table: " + table["name"].InnerText);
                    AddTable(xlWorkBook, table, nsmgr);
                }

            }

            antTables = "Number of tables " + tables.ChildNodes.Count.ToString();

            xlWorkBook.SaveAs(Path.ChangeExtension(Path.GetFullPath(filename), ".xls"), XlFileFormat.xlWorkbookNormal);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp1.Quit();


            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp1);

        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a worksheet with information about the database.
        private void AddDBInfo(Workbook workbook, XmlNode table, XmlNamespaceManager nsmgr)
        {
            Worksheet DBWorkSheet = (Worksheet)workbook.Worksheets.get_Item(1);
            DBWorkSheet.Name = "DB";

            List<string> fieldNames = new List<string>()
            {
                "dbname",
                "description",
                "archiver",
                "archiverContact",
                "dataOwner",
                "dataOriginTimespan",
                "producerApplication",
                "archivalDate",
                "clientMachine",
                "databaseProduct"
            };

            int cnt = 1;

            foreach (string s in fieldNames)
            {
                DBWorkSheet.Cells[cnt, 1] = s;
                string verdi = DBWorkSheet.Cells[cnt, 1].Text;
                DBWorkSheet.Cells[cnt, 2] = getNodeText(table, "//siard:" + s, nsmgr);
                cnt++;
            }

            DBWorkSheet.Cells[cnt, 1] = "schemas";
            XmlNode schemas = table.SelectSingleNode("//siard:schemas", nsmgr);

            foreach (XmlNode schema in schemas.ChildNodes)
            {
                DBWorkSheet.Cells[cnt, 2] = getNodeText(schema, "descendant::siard:folder", nsmgr);
                cnt++;
            }

            DBWorkSheet.Cells[cnt, 1] = "users";
            XmlNode users = table.SelectSingleNode("//siard:users", nsmgr);

            foreach (XmlNode user in users.ChildNodes)
            {
                DBWorkSheet.Cells[cnt, 2] = getNodeText(user, "descendant::siard:name", nsmgr);
                cnt++;
            }

            Marshal.ReleaseComObject(DBWorkSheet);

        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a worksheet with information for each table
        private void AddTable(Workbook workbook, XmlNode table, XmlNamespaceManager nsmgr)
        {
            Worksheet tableWorksheet;
            tableWorksheet = (Worksheet)workbook.Application.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            tableWorksheet.Name = table["name"].InnerText;

            int cellCount = 2;

            List<string> columnNames = new List<string>()
            {
                "Column",
                "Name",
                "Type",
                "Type original",
                "Nullable",
                "Default value",
                "Description",
                "Note"
            };

            foreach (string name in columnNames)
            {
                tableWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
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
               // Console.WriteLine(rn[1]);
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

            Marshal.ReleaseComObject(tableWorksheet);

        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Returns Innertext of node found in table with query.
        private string getNodeText(XmlNode table, string query, XmlNamespaceManager nsmgr)
        {
            string varName = "";
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
        private void AddTableOverview(Application excelApp, Workbook workbook, XmlNode tables)
        {
            Worksheet tableOverviewWorksheet = (Worksheet)workbook.Application.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            tableOverviewWorksheet.Name = "Tables";

            List<string> columnNames = new List<string>()
            {
                "Table",
                "Folder",
                "Schema",
                "Rows"
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
                Range c1 = tableOverviewWorksheet.Cells[count, 1];
                Range c2 = tableOverviewWorksheet.Cells[count, 1];
                Range linkCell = excelApp.get_Range(c1, c2);

                tableOverviewWorksheet.Hyperlinks.Add(linkCell, "", name + "!A1", "", name);

                tableOverviewWorksheet.Cells[count, 2] = table["folder"].InnerText;
                tableOverviewWorksheet.Cells[count, 3] = table.ParentNode.ParentNode["folder"].InnerText;
                tableOverviewWorksheet.Cells[count, 4] = table["rows"].InnerText;
                count++;
            }

            Marshal.ReleaseComObject(tableOverviewWorksheet);

        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    }
    //==========================================================================================================
}
