using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

namespace KDRS_Metadata
{
    class DataConverter
    {
        public int totalTableCount;
        public int totalSchemaCount;
        public List<Schema> schemaNames = new List<Schema>();
        public string excelFileName;
        public string siardVersion;

        // ToDo: Replace hardcoded array 30 schemas with actual number of schemas
        public int thisSchemaNo;

        // #1: Schema number [0..n]
        // #2: Table counters number [0..2]
        //     Total rows, Max rows one table, Max columns one table
        public int[,] arrayTableCounters = new int[30, 3];

        // #1: Schema number [0..n]
        // #2: Counter number [0..7]
        //     PKs, FKs, CKs, noPKs, noFKs, noCKs, yesFKs, yesCKs
        
        public int[,] arrayKeysCounters = new int[30, 8];

        public delegate void ProgressUpdate(int count, int totalCount, string progressPostfix);
        public event ProgressUpdate OnProgressUpdate;

        public void Convert(string filename, bool includeTables)
        {
            int tempInt;

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
            Sheets xlWorksheets;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlWorkbooks.Add(misValue);
            xlWorksheets = xlWorkBook.Sheets;

            Worksheet DBWorksheet = xlWorksheets.get_Item(1);
            AddDBInfo(DBWorksheet, root, nsmgr);
            Marshal.ReleaseComObject(DBWorksheet);

            XmlNodeList schemas = root.SelectNodes("descendant::siard:schema", nsmgr);
            totalSchemaCount = schemas.Count;
            XmlNodeList allTables = root.SelectNodes("//siard:tables/siard:table", nsmgr);

            Console.WriteLine("Schemas read");

            int tableCount = 0;
            totalTableCount = allTables.Count;

            Worksheet tableOverviewWorksheet = xlWorksheets.Add(After: xlWorksheets[xlWorksheets.Count]);
            AddTableOverview(tableOverviewWorksheet, schemas, nsmgr, includeTables);

            Marshal.ReleaseComObject(tableOverviewWorksheet);

            Console.WriteLine("Added tableoverview");

            thisSchemaNo = 0;
            foreach (XmlNode schema in schemas)
            {
                string schemaName = getInnerText(schema["name"]);
                string schemaFolder = getInnerText(schema["folder"]);
                schemaNames.Add(new Schema(getInnerText(schema["name"]), getInnerText(schema["folder"])));
                XmlNode tables = schema.SelectSingleNode("descendant::siard:tables", nsmgr);
                Console.WriteLine("Enter schema");

                for (int n = 0; n < 3; n++)
                {
                    arrayTableCounters[thisSchemaNo, n] = 0;
                }
                for (int n = 0; n < 8; n++)
                {
                    arrayKeysCounters[thisSchemaNo, n] = 0;
                }

                if (includeTables)
                {
                    foreach (XmlNode table in tables.ChildNodes)
                    {
                        tempInt = Int32.Parse(getInnerText(table["rows"]));
                        arrayTableCounters[thisSchemaNo, 0] += tempInt;
                        if (tempInt > arrayTableCounters[thisSchemaNo, 1])
                            arrayTableCounters[thisSchemaNo, 1] = tempInt;
                        Console.WriteLine("tableRows = " + tempInt);

                        string tableName = getInnerText(table["name"]);
                        Worksheet tableWorksheet = xlWorksheets.Add(After: xlWorksheets[xlWorksheets.Count]);

                        AddTable(tableWorksheet, table, nsmgr);
                        tableCount++;
                        Console.WriteLine("Added table");

                        OnProgressUpdate?.Invoke(tableCount, totalTableCount, schemaFolder + ": " + schemaName +"  |  "+ tableName);
                    }
                }

                for (int n = 0; n < 8; n++)
                {
                    Console.WriteLine(thisSchemaNo + "." + n + " = " + arrayKeysCounters[thisSchemaNo, n].ToString());
                }
                thisSchemaNo++;
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

            excelFileName = checkFileName(excelFileName, fileCounter);
            //Console.WriteLine("Outfile: " + checkedFileNAme);
                       
            xlWorkBook.SaveAs(excelFileName);

            xlApp1.UseSystemSeparators = true;

            xlWorkBook.Close();
            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlApp1);

        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        // Creates a Worksheet with information about the database.
        private void AddDBInfo(Worksheet DBWorksheet, XmlNode table, XmlNamespaceManager nsmgr)
        {
            Range tempRng;

            DBWorksheet.Name = "db";
            DBWorksheet.Columns.AutoFit();

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
            DBWorksheet.Cells[cnt, 1] = fieldNames[0];
            DBWorksheet.Cells[cnt, 2] = Globals.toolName;
            cnt++;

            // toolVersion
            DBWorksheet.Cells[cnt, 1] = fieldNames[1];
            DBWorksheet.Cells[cnt, 2] = Globals.toolVersion;
            DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
            cnt++;

            for (int i=2; i<7; i++)
            {
                DBWorksheet.Cells[cnt, 1] = fieldNames[i];
                DBWorksheet.Cells[cnt, 2] = "";
                cnt++;
            }            

            //tableCount
            DBWorksheet.Cells[cnt, 1] = fieldNames[7];
            XmlNodeList  tableCount = table.SelectNodes("//siard:table", nsmgr);
            DBWorksheet.Cells[cnt, 2] = tableCount.Count;
            cnt++;

            // Blank row
            DBWorksheet.Cells[cnt, 1] = fieldNames[8];
            DBWorksheet.Cells[cnt, 2] = "";
            cnt++;

            // SIARD metadata.xml
            DBWorksheet.Cells[cnt, 1] = fieldNames[9];
            DBWorksheet.Cells[cnt, 2] = "metadata.xml";
            cnt++;

            // SIARD version
            DBWorksheet.Cells[cnt, 1] = fieldNames[10];
            DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
            siardVersion = table.Attributes["version"].Value;
            DBWorksheet.Cells[cnt, 2] = siardVersion;
            cnt++;

            for (int i = 11; i < 20; i++)
            {
                string field = fieldNames[i];
                DBWorksheet.Cells[cnt, 1] = field;
                if (i == 11)
                {
                    DBWorksheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + field, nsmgr));
                }
                else
                {
                    DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:" + field, nsmgr);
                }
                    cnt++;
            }

            if ("2.1" == siardVersion)
            {
                //digestType
                DBWorksheet.Cells[cnt, 1] = fieldNames[20];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:messageDigest/digestType", nsmgr);
                cnt++;

                //digest
                DBWorksheet.Cells[cnt, 1] = fieldNames[21];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:messageDigest/digest", nsmgr);
                cnt++;
            }
            else if ("2.0" == siardVersion)
            {
                //digestType
                DBWorksheet.Cells[cnt, 1] = fieldNames[20];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = "";
                cnt++;

                //digest
                DBWorksheet.Cells[cnt, 1] = fieldNames[21];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:messageDigest", nsmgr);
                cnt++;
            }
            else if ("1.0" == siardVersion)
            {
                //digestType
                DBWorksheet.Cells[cnt, 1] = fieldNames[20];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = "";
                cnt++;

                //digest
                DBWorksheet.Cells[cnt, 1] = fieldNames[21];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:messageDigest", nsmgr);
                cnt++;
            }
            else
            {
                //digestType
                DBWorksheet.Cells[cnt, 1] = fieldNames[20];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = "SIARD version";
                cnt++;

                //digest
                DBWorksheet.Cells[cnt, 1] = fieldNames[21];
                DBWorksheet.Cells[cnt, 2].NumberFormat = "@";
                DBWorksheet.Cells[cnt, 2] = "Unknown";
                cnt++;
            }

            //clientMachine
            DBWorksheet.Cells[cnt, 1] = fieldNames[22];
            DBWorksheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[22], nsmgr));
            cnt++;

            //databaseProduct
            DBWorksheet.Cells[cnt, 1] = fieldNames[23];
            DBWorksheet.Cells[cnt, 2] = GetNodeText(table, "//siard:" + fieldNames[23], nsmgr);
            cnt++;

            //connection
            DBWorksheet.Cells[cnt, 1] = fieldNames[24];
            DBWorksheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[24], nsmgr));
            cnt++;

            //databaseUser
            DBWorksheet.Cells[cnt, 1] = fieldNames[25];
            DBWorksheet.Cells[cnt, 2] = SensitiveString(GetNodeText(table, "//siard:" + fieldNames[25], nsmgr));
            cnt++;

            DBWorksheet.Cells[cnt, 1] = "schemas";
            XmlNodeList schemas = table.SelectNodes("//siard:schemas/siard:schema", nsmgr);

            string schemasList = GetNodeText(schemas[0], "descendant::siard:folder", nsmgr);
            for (int i=1; i<schemas.Count; i++)
            {
                schemasList += ", " + GetNodeText(schemas[i], "descendant::siard:folder", nsmgr);
            }
            DBWorksheet.Cells[cnt, 2] = schemasList;
            cnt++;

            DBWorksheet.Cells[cnt, 1] = "users";
            XmlNode users = table.SelectSingleNode("//siard:users", nsmgr);
            DBWorksheet.Cells[cnt, 2] = getChildCount(users);
            cnt++;

            DBWorksheet.Cells[cnt, 1] = "roles";
            XmlNode roles = table.SelectSingleNode("//siard:roles", nsmgr);
            DBWorksheet.Cells[cnt, 2] = getChildCount(roles);
            cnt++;

            DBWorksheet.Cells[cnt, 1] = "privileges";
            XmlNode privileges = table.SelectSingleNode("//siard:privileges", nsmgr);
            DBWorksheet.Cells[cnt, 2] = getChildCount(privileges);

            // Freeze Panes
            tempRng = DBWorksheet.Cells[10, 1];
            tempRng.Activate();
            tempRng.Application.ActiveWindow.FreezePanes = true;

            // Header rows bold, volor & background color
            tempRng = DBWorksheet.Range["A1", "B1"];
            tempRng.Characters.Font.Bold = true;
            tempRng.Interior.Color = Color.LightGray;

            // tempRng = DBWorksheet.Range["A3", "B6"];
            // tempRng.Characters.Font.Color = Color.Red;

            tempRng = DBWorksheet.Range["A3", "B6"];
            tempRng.Interior.Color = Color.LightYellow;

            // tempRng = DBWorksheet.Range["A7", "B7"];
            // tempRng.Characters.Font.Color = Color.Orange;

            tempRng = DBWorksheet.Range["A7", "B7"];
            tempRng.Interior.Color = Color.LightSkyBlue;

            tempRng = DBWorksheet.Range["A10", "B10"];
            tempRng.Characters.Font.Bold = true;
            tempRng.Interior.Color = Color.LightGray;

            // Border lines
            for (int m = 1; m < 9; m++)
            {
                for (int n = 1; n < 3; n++)
                {
                    tempRng = DBWorksheet.Cells[m, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }
            }

            for (int m = 10; m < 31; m++)
            {
                for (int n = 1; n < 3; n++)
                {
                    tempRng = DBWorksheet.Cells[m, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }
            }

            // Column widths
            DBWorksheet.Columns["A:A"].ColumnWidth = 20;
            DBWorksheet.Columns["B:B"].ColumnWidth = 120;
            DBWorksheet.Columns["B:B"].WrapText = true;

            // Alignment
            DBWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            DBWorksheet.Columns.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Marshal.ReleaseComObject(DBWorksheet);
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Creates a Worksheet with table overview
        private void AddTableOverview(Worksheet tableOverviewWorksheet, XmlNodeList schemas, XmlNamespaceManager nsmgr, bool includeTables)
        {
            Range tempRng;

            tableOverviewWorksheet.Name = "tables";
            tableOverviewWorksheet.Columns.AutoFit();

            List<string> columnNames = new List<string>()
            {
                "table",
                "folder",
                "schema",
                "rows",
                "columns",
                "priority",
                // "pri-sort",
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

            thisSchemaNo = 0;
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

                    string tablePriority = GetNodeTxtEmpty(table, "descendant::siard:priority", nsmgr);
                    if (string.IsNullOrEmpty(tablePriority) && tableRows == "0")
                        tablePriority = "EMPTY";
                    tableOverviewWorksheet.Cells[count, 6] = tablePriority;
                    
                    /* int tablePriSort = Globals.PriSort(tablePriority);
                     tableOverviewWorksheet.Cells[count, 7] = tablePriSort;
                     */

                    string table_entity = GetNodeTxtEmpty(table, "descendant::siard:description", nsmgr);
                    tableOverviewWorksheet.Cells[count, 7] = ExtractEntity(table_entity, "entity");
                    
                    string table_description = GetNodeTxtEmpty(table, "descendant::siard:description", nsmgr);
                    tableOverviewWorksheet.Cells[count, 8] = ExtractEntity(table_description, "description");

                    count++;
                }
                thisSchemaNo++;
            }            

            // Freeze Panes
            tempRng = tableOverviewWorksheet.Cells[2, 1];
            tempRng.Activate();
            tempRng.Application.ActiveWindow.FreezePanes = true;

            tempRng = tableOverviewWorksheet.Range["A1", "I1"];
            tempRng.Characters.Font.Bold = true;

            // Border lines
            for (int n = 1; n < 10; n++)                
            {
                if (n < 6)
                {
                    tempRng = tableOverviewWorksheet.Cells[1, n];
                    tempRng.Interior.Color = Color.LightGray;
                }

                for (int m = 1; m < count; m++)
                {
                    tempRng = tableOverviewWorksheet.Cells[m, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;                    
                }
            }

            // Cell background color
            for (int m = 1; m < count; m++)
            {
                tempRng = tableOverviewWorksheet.Cells[m, 6];
                tempRng.Interior.Color = Color.LightYellow;

                tempRng = tableOverviewWorksheet.Cells[m, 7];
                tempRng.Interior.Color = Color.LightGreen;

                tempRng = tableOverviewWorksheet.Cells[m, 8];
                tempRng.Interior.Color = Color.LightSkyBlue;

                tempRng = tableOverviewWorksheet.Cells[m, 9];
                tempRng.Interior.Color = Color.LightGray;
            }

            // Alignment
            tableOverviewWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableOverviewWorksheet.Columns.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Column widths
            tableOverviewWorksheet.Columns["A:A"].AutoFit();
            tableOverviewWorksheet.Columns["B:B"].AutoFit();  // .ColumnWidth = 8;
            tableOverviewWorksheet.Columns["C:C"].AutoFit();  // .ColumnWidth = 8;

            tableOverviewWorksheet.Columns["D:D"].AutoFit();
            tableOverviewWorksheet.Columns["D:D"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["E:E"].AutoFit();
            tableOverviewWorksheet.Columns["E:E"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["F:F"].ColumnWidth = 10;
            tableOverviewWorksheet.Columns["F:F"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            tableOverviewWorksheet.Columns["G:G"].ColumnWidth = 20;
            tableOverviewWorksheet.Columns["G:G"].WrapText = true;

            tableOverviewWorksheet.Columns["H:H"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["H:H"].WrapText = true;

            tableOverviewWorksheet.Columns["I:I"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["I:I"].WrapText = true;

            // Column sorting
            tableOverviewWorksheet.Sort.SortFields.Clear();

            tableOverviewWorksheet.Sort.SortFields.Add(tableOverviewWorksheet.Range["F:F"], XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, "HIGH, MEDIUM, LOW, SYSTEM, STATS, EMPTY, DUMMY", XlSortDataOption.xlSortNormal);
            tableOverviewWorksheet.Sort.SetRange(tableOverviewWorksheet.UsedRange);
            tableOverviewWorksheet.Sort.Header = XlYesNoGuess.xlYes;
            tableOverviewWorksheet.Sort.Apply();
            
            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Creates a Worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, XmlNode table, XmlNamespaceManager nsmgr)
        {
            Range tempRng;

            string schemaNumber = GetNumbers(table.ParentNode.ParentNode["folder"].InnerText);

            if (totalSchemaCount < 2)
                tableWorksheet.Name = GetNumbers(table["folder"].InnerText);
            else 
                tableWorksheet.Name = schemaNumber + "." + GetNumbers(table["folder"].InnerText);
            tableWorksheet.Columns.AutoFit();

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
            // thisRowCount += Int32.Parse(tableRows);

            string tablePriority = GetNodeTxtEmpty(table, "descendant::siard:priority", nsmgr);
            if (tableRows == "0")
                tablePriority = "EMPTY";

            // string tableEntity = GetNodeTxtEmpty(table, "descendant::siard:entity", nsmgr);

            // Table header
            string[][] rowNamesArray = new string[9][] 
            {
                new string[2] { "schemaName", table.ParentNode.ParentNode["name"].InnerText.ToString() },
                new string[2] { "schemaFolder", table.ParentNode.ParentNode["folder"].InnerText.ToString()},
                new string[2] { "tableName", table["name"].InnerText.ToString() },
                new string[2] { "tableFolder", getInnerText(table["folder"]) },
                new string[2] { "tablePriority", tablePriority },
                new string[2] { "tableEntity", ExtractEntity(table_description, "entity")},
                new string[2] { "tableDescription", ExtractEntity(table_description, "description")},
                new string[2] { "rows", tableRows },
                new string[2] { "columns", getChildCount(table["columns"]) }
            };

            foreach (string[] rn in rowNamesArray)
            {
                tableWorksheet.Cells[cellCount, 1] = rn;
                tableWorksheet.Cells[cellCount, 2] = rn[1];
                cellCount++;
            }

            // Primary Key
            if ("[NA]" == primaryKey_name)
            {
                arrayKeysCounters[thisSchemaNo, 3]++;  // noPKs
            }
            else
            {
                arrayKeysCounters[thisSchemaNo, 0]++;  // PKs == (yesPKs)
                tempRng = tableWorksheet.Cells[cellCount, 1];
                tempRng.Interior.Color = Color.LightGray;

                tempRng = tableWorksheet.Cells[cellCount, 2];
                tempRng.Interior.Color = Color.LightGray;

                string pk_decription = GetNodeTxtEmpty(table["primaryKey"], "descendant::siard:description", nsmgr);
                string pk_extr_entity = ExtractEntity(pk_decription, "entity");
                string pk_extr_description = ExtractEntity(pk_decription, "description");
                rowNamesArray = new string[3][] 
                {
                    new string[2] { "pkName", primaryKey_name },
                    new string[2] { "pkEntity", pk_extr_entity},
                    new string[2] { "pkDescription", pk_extr_description }
                    // new string[2] { "pkDescription", GetNodeText(table["primaryKey"], "descendant::siard:description", nsmgr) }
                };

                foreach (string[] rn in rowNamesArray)
                {
                    tableWorksheet.Cells[cellCount, 1] = rn;
                    tableWorksheet.Cells[cellCount, 2] = rn[1];
                    cellCount++;
                }

                for (int n = 1; n < 8; n++)
                {
                    tempRng = tableWorksheet.Cells[cellCount - 2, n];
                    tempRng.Interior.Color = Color.LightGreen;

                    tempRng = tableWorksheet.Cells[cellCount - 1, n];
                    tempRng.Interior.Color = Color.LightSkyBlue;
                }
                
                XmlNode pKey = table.SelectSingleNode("descendant::siard:primaryKey", nsmgr);

                for (int i = 1; i < pKey.ChildNodes.Count; i++)
                {
                    string primaryKey_column1 = GetNodeText(table["primaryKey"], "descendant::siard:column[" + i + "]", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "pkColumn";
                    tableWorksheet.Cells[cellCount, 2] = primaryKey_column1;
                    cellCount++;
                }
            }

            //-------------------------------------------------------------------------------------
            // Finds all foreign keys in table and prints to Excel.
            XmlNode foreignKeys = table.SelectSingleNode("descendant::siard:foreignKeys", nsmgr);

            if (foreignKeys == null)
            {
                arrayKeysCounters[thisSchemaNo, 4]++;  // noFKs
            }
            else
            {
                arrayKeysCounters[thisSchemaNo, 6]++;  // yesFKs
                foreach (XmlNode fKey in foreignKeys.ChildNodes)
                {
                    arrayKeysCounters[thisSchemaNo, 1]++;  // FKs
                    tempRng = tableWorksheet.Cells[cellCount, 1];
                    tempRng.Interior.Color = Color.LightPink;

                    tempRng = tableWorksheet.Cells[cellCount, 2];
                    tempRng.Interior.Color = Color.LightPink;

                    string foreignKeys_name = GetNodeText(fKey, "descendant::siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "fkName";
                    tableWorksheet.Cells[cellCount, 2] = foreignKeys_name;
                    cellCount++;

                    string foreignKeys_ref_schema = GetNodeTxtEmpty(fKey, "descendant::siard:referencedSchema", nsmgr);
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

                    string fk_description = GetNodeTxtEmpty(fKey, "descendant::siard:description", nsmgr);
                    string fk_extr_entity = ExtractEntity(fk_description, "entity");
                    string fk_extr_description = ExtractEntity(fk_description, "description");
                    // string foreignKeys_description = GetNodeTxtEmpty(fKey, "descendant::siard:description", nsmgr);

                    tableWorksheet.Cells[cellCount, 1] = "fkEntity";
                    tableWorksheet.Cells[cellCount, 2] = fk_extr_entity;
                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightGreen;
                    }
                    cellCount++;

                    tableWorksheet.Cells[cellCount, 1] = "fkDescription";
                    tableWorksheet.Cells[cellCount, 2] = fk_extr_description;
                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightSkyBlue;
                    }
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

            if (candidateKeys == null)
            {                
                arrayKeysCounters[thisSchemaNo, 5]++;  // noCKs
            }
            else
            {                
                arrayKeysCounters[thisSchemaNo, 7]++;  // yesCKs
                foreach (XmlNode cKey in candidateKeys.ChildNodes)
                {
                    arrayKeysCounters[thisSchemaNo, 2]++;  // CKs
                    tempRng = tableWorksheet.Cells[cellCount, 1];
                    tempRng.Interior.Color = Color.PaleTurquoise;

                    tempRng = tableWorksheet.Cells[cellCount, 2];
                    tempRng.Interior.Color = Color.PaleTurquoise;

                    string candidateKeys_name = GetNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:name", nsmgr);
                    tableWorksheet.Cells[cellCount, 1] = "ckName ";
                    tableWorksheet.Cells[cellCount, 2] = candidateKeys_name;
                    cellCount++;

                    string ck_description = GetNodeTxtEmpty(table["candidateKeys"], "descendant::siard:candidateKey/siard:description", nsmgr);
                    string ck_extr_entity = ExtractEntity(ck_description, "entity");
                    string ck_extr_description = ExtractEntity(ck_description, "description");
                    // string candidateKeys_description = GetNodeText(table["candidateKeys"], "descendant::siard:candidateKey/siard:description", nsmgr);

                    tableWorksheet.Cells[cellCount, 1] = "ckEntity";
                    tableWorksheet.Cells[cellCount, 2] = ck_extr_entity;
                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightGreen;
                    }
                    cellCount++;

                    tableWorksheet.Cells[cellCount, 1] = "ckDescription";
                    tableWorksheet.Cells[cellCount, 2] = ck_extr_description;
                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightSkyBlue;
                    }
                    cellCount++;

                    for (int n = 1; n < 9; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightSkyBlue;
                    }
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

            // Repeat header row for visibility at top of first Table Column row
            foreach (string name in columnNames.Skip(1))
            {
                tableWorksheet.Cells[cellCount, columnNames.IndexOf(name) + 1] = name;
            }

            // Repeat header row bold & border lines
            for (int n = 1; n < 11; n++)
            {
                tempRng = tableWorksheet.Cells[cellCount, n];
                tempRng.Characters.Font.Bold = true;
                tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                if (n < 8)
                {
                    tempRng.Interior.Color = Color.LightGray;
                }
            }

            // Cell background color
            tempRng = tableWorksheet.Cells[cellCount, 8];
            tempRng.Interior.Color = Color.LightGreen;

            tempRng = tableWorksheet.Cells[cellCount, 9];
            tempRng.Interior.Color = Color.LightSkyBlue;

            tempRng = tableWorksheet.Cells[cellCount, 10];
            tempRng.Interior.Color = Color.LightGray;

            cellCount++;

            // Insert Table Column rows
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

                    string col_lobFolder = GetNodeTxtEmpty(column, "descendant::siard:lobFolder", nsmgr);
                    tableWorksheet.Cells[cellCount, 7] = col_lobFolder;

                    string col_entity = GetNodeTxtEmpty(column, "descendant::siard:description", nsmgr);
                    string col_entity_extr = ExtractEntity(col_entity, "entity");
                    tableWorksheet.Cells[cellCount, 8] = col_entity_extr;

                    string col_description = GetNodeTxtEmpty(column, "descendant::siard:description", nsmgr);
                    string col_description_extr = ExtractEntity(col_description, "description");
                    tableWorksheet.Cells[cellCount, 9] = col_description_extr;

                    string col_note = GetNodeTxtEmpty(column, "descendant::siard:note", nsmgr);
                    tableWorksheet.Cells[cellCount, 10] = col_note;

                    // Border line
                    for (int n = 1; n < 11; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;                        
                    }

                    if ("" != col_lobFolder)
                    {
                        for (int n = 1; n < 8; n++)
                        {
                            tempRng = tableWorksheet.Cells[cellCount, n];
                            tempRng.Interior.Color = Color.LightYellow;
                        }
                    }

                    // Background color
                    tempRng = tableWorksheet.Cells[cellCount, 8];
                    tempRng.Interior.Color = Color.LightGreen;

                    tempRng = tableWorksheet.Cells[cellCount, 9];
                    tempRng.Interior.Color = Color.LightSkyBlue;

                    tempRng = tableWorksheet.Cells[cellCount, 10];
                    tempRng.Interior.Color = Color.LightGray;

                    cellCount++;
                }
            }

            // Range range = tableWorksheet.Cells[5, 1];
            tempRng = tableWorksheet.Cells[9, 3];
            tempRng.Activate();
            tempRng.Application.ActiveWindow.FreezePanes = true;

            // First row bold
            tempRng = tableWorksheet.Range["A1", "J1"];
            tempRng.Characters.Font.Bold = true;

            // Border lines            
            for (int n = 1; n < 11; n++)
            {
                tempRng = tableWorksheet.Cells[1, n];
                tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                if (n < 8)
                {
                    tempRng.Interior.Color = Color.LightGray;
                }
            }

            for (int m = 2; m < 8; m++)
            {
                for (int n = 1; n < 3; n++)
                {
                    tempRng = tableWorksheet.Cells[m, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }
            }

            // Cell background color            
            tempRng = tableWorksheet.Cells[1, 8];
            tempRng.Interior.Color = Color.LightGreen;

            tempRng = tableWorksheet.Cells[1, 9];
            tempRng.Interior.Color = Color.LightSkyBlue;

            tempRng = tableWorksheet.Cells[1, 10];
            tempRng.Interior.Color = Color.LightGray;

            tempRng = tableWorksheet.Range["A6", "B6"];
            tempRng.Interior.Color = Color.LightYellow;

            tempRng = tableWorksheet.Range["A7", "J7"];
            tempRng.Interior.Color = Color.LightGreen;            
            // ToDo: Make Wrap Text expand to selected range, not only the single cell
            // tempRng.WrapText = true;
            // tempRng.Style.WrapText = true;
            tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            tempRng = tableWorksheet.Range["A8", "J8"];
            tempRng.Interior.Color = Color.LightSkyBlue;
            // ToDo: Make Wrap Text expand to selected range, not only the single cell
            // tempRng.WrapText = true;
            // tempRng.Style.WrapText = true;
            tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            // Column widths
            tableWorksheet.Columns["A:A"].AutoFit();

            tableWorksheet.Columns["B:B"].ColumnWidth = 30;

            tableWorksheet.Columns["C:C"].AutoFit();
            tableWorksheet.Columns["D:D"].AutoFit();
            tableWorksheet.Columns["E:E"].AutoFit();
            tableWorksheet.Columns["F:F"].AutoFit();

            tableWorksheet.Columns["G:G"].AutoFit();  // .ColumnWidth = 26;
            tableWorksheet.Columns["G:G"].WrapText = true;

            tableWorksheet.Columns["H:H"].ColumnWidth = 30;
            tableWorksheet.Columns["H:H"].WrapText = true;

            tableWorksheet.Columns["I:I"].ColumnWidth = 60;
            tableWorksheet.Columns["I:I"].WrapText = true;

            tableWorksheet.Columns["J:J"].ColumnWidth = 60;
            tableWorksheet.Columns["J:J"].WrapText = true;

            // Alignment
            tableWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;            
            tableWorksheet.Columns.VerticalAlignment = XlVAlign.xlVAlignCenter;
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
                    if (string.IsNullOrEmpty(varName) && query != "descendant::siard:deleteAction" && query != "descendant::siard:updateAction")
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
                if (string.IsNullOrEmpty(varName))
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
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Extracts entities from the description, returns the entities or description depending on target type
        private string ExtractEntity(string description, string targetType)
        {
            Regex regex = new Regex(@"(?<entity1>(?<=\{)[^\{\}]+(?=\}))|(?<entity2>(?<=\[)[^\[\]]+(?=\]))|(?<desc>\w+.*)");

            string entities = "";
            string cleanDescription = "";

            if (description != null)
            {
                foreach (Match m in regex.Matches(description))
                {
                    if (m.Groups["entity1"].Value != "")
                    {
                        entities += "{" + m.Groups["entity1"].Value + "}";
                    }

                    if (m.Groups["entity2"].Value != "")
                    {
                        entities += "[" + m.Groups["entity2"].Value + "]";
                    }

                    if (m.Groups["desc"].Value != "")
                    {
                        cleanDescription += m.Groups["desc"].Value;
                    }
                }
            }

            if (targetType == "entity")
                return entities;

            return cleanDescription;
        }
    }
    //==========================================================================================================
}
