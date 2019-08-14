using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KDRS_Metadata
{
    class JsonReader

    {
        public int totalTableCount;
        public string excelFileName;
        public int tableCount;

        public bool includeTables;

        public delegate void ProgressUpdate(int count, int totalCount);
        public event ProgressUpdate OnProgressUpdate;

        public void ParseJson(string filename, List<string> priorities, bool includeTables)
        {

            this.includeTables = includeTables;

            Microsoft.Office.Interop.Excel.Application xlApp1 = new Microsoft.Office.Interop.Excel.Application();

            //xlApp1.Visible = true;

            Workbooks xlWorkBooks;
            Workbook xlWorkBook;

            Sheets xlWorkSheets;

            xlWorkBooks = xlApp1.Workbooks;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlWorkBooks.Add(misValue);

            xlWorkSheets = xlWorkBook.Sheets;

            string json;
            using (StreamReader r = new StreamReader(filename))
            {
                json = r.ReadToEnd();
            }

            Template template = JsonConvert.DeserializeObject<Template>(json);

            Worksheet templateSheet = xlWorkSheets.get_Item(1);
            AddTemplateInfo(templateSheet, template);
            Marshal.ReleaseComObject(templateSheet);

            Worksheet tableOverviewWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);
            AddTableOverview(tableOverviewWorksheet, template.TemplateSchema, priorities);
            Marshal.ReleaseComObject(tableOverviewWorksheet);

            Console.WriteLine("After");
            foreach (string l in priorities)
            {
                Console.WriteLine(l);
            }

            tableCount = 0;
            totalTableCount = template.TemplateSchema.Tables.Count;

            if (includeTables)
            {
                foreach (Table table in template.TemplateSchema.Tables)
                {
                    if (priorities.Contains(table.TablePriority))
                    {

                        Worksheet tableWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);

                        AddTable(tableWorksheet, template.TemplateSchema, table);

                        Marshal.ReleaseComObject(tableWorksheet);

                    }
                    tableCount++;

                    OnProgressUpdate?.Invoke(tableCount, totalTableCount);
                }
            }
            xlWorkBook.Sheets[1].Select();


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

            xlWorkBook.SaveAs(excelFileName);

            Marshal.ReleaseComObject(xlWorkSheets);

            xlWorkBook.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkBook);

            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp1);

            Console.WriteLine("App: " + xlApp1);
        }

        //*************************************************************************

        // Creates a worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, Schema schema, Table table)
        {

            tableWorksheet.Name = GetNumbers(table.Folder);

            Range c1 = tableWorksheet.Cells[1, 1];
            Range c2 = tableWorksheet.Cells[1, 1];
            Range linkCell = tableWorksheet.get_Range(c1, c2);

            Hyperlinks links = tableWorksheet.Hyperlinks;

            links.Add(linkCell, "", "tables!A1", "", "column <<< tables");

            //tableWorksheet.Name = table.Name;

            List<string> columnNames = new List<string>()
            {
                "column",
                "name",
                "type",
                "folder",
                "entity",
                "description",
                "note"
            };

            foreach (string name in columnNames.Skip(1))
            {
                tableWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
            }
            //-------------------------------------------------------------------
            string[][] rowNamesArray = new string[9][] {
                new string[2] { "schemaName", schema.Name },
                new string[2] { "schemaFolder", schema.Folder},
                new string[2] { "tableName", table.Name },
                new string[2] { "tableFolder", table.Folder },
                new string[2] { "tablePriority", table.TablePriority },
                new string[2] { "tableEntity", table.TableEntity },
                new string[2] { "tableDescription", table.Description },
                new string[2] { "rows", table.Rows.ToString() },
                new string[2] { "columns", table.Columns.Count().ToString() },
            };

            int count = 2;

            foreach (string[] rn in rowNamesArray)
            {
                /*if (rn[0] == "tableDescription")
                {
                    tableWorksheet.Cells[count, 1] = rn;
                    tableWorksheet.Cells[count, 3] = rn[1];
                }
                else
                {*/
                    tableWorksheet.Cells[count, 1] = rn;
                    tableWorksheet.Cells[count, 2] = rn[1];
               // }

                count++;
            }

            if (table.PrimaryKey != null)
            {
                tableWorksheet.Cells[count, 1] = "pkName";
                tableWorksheet.Cells[count, 2] = table.PrimaryKey.Name;
                tableWorksheet.Cells[count, 2] = table.PrimaryKey.Description;
                count++;
            }

            int columnCount = 0;
            foreach (Column column in table.Columns)
            {
                GetEntity(column.Description, null, column);

                tableWorksheet.Cells[count, 1] = "Column " + columnCount;
                tableWorksheet.Cells[count, 2] = column.Name;
                tableWorksheet.Cells[count, 3] = column.Datatype;
                tableWorksheet.Cells[count, 4] = column.Folder;
                tableWorksheet.Cells[count, 5] = column.Entity;
                tableWorksheet.Cells[count, 6] = column.Description;
                tableWorksheet.Cells[count, 7] = "";
                count++;

                columnCount++;
            }

            Range range = tableWorksheet.Cells[5, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            tableWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableWorksheet.Columns.AutoFit();

            Marshal.ReleaseComObject(tableWorksheet);
        }

        //*************************************************************************
        
        // Creates a worksheet with table overview
        private void AddTableOverview(Worksheet tableOverviewWorksheet, Schema schema, List<string> priorities)
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
            //-------------------------------------------------------------------
            int count = 2;
            foreach (Table table in schema.Tables)
            {
                Console.WriteLine("Table: " + table.Name + ", Description: " + table.Description);
                GetEntity(table.Description, table);
                Console.WriteLine("Table: " + table.Name + ", Description: " + table.Description);


                if (priorities.Contains(table.TablePriority) && includeTables)
                {
                    Range c1 = tableOverviewWorksheet.Cells[count, 1];
                    Range c2 = tableOverviewWorksheet.Cells[count, 1];
                    Range linkCell = tableOverviewWorksheet.get_Range(c1, c2);

                    Hyperlinks links = tableOverviewWorksheet.Hyperlinks;
                    links.Add(linkCell, "", GetNumbers(table.Folder) + "!A1", "", table.Name);

                    //tableOverviewWorksheet.Cells[count, 1] = table.Name;
                    tableOverviewWorksheet.Cells[count, 2] = table.Folder;
                    tableOverviewWorksheet.Cells[count, 3] = schema.Name;
                    tableOverviewWorksheet.Cells[count, 4] = table.Rows;
                    tableOverviewWorksheet.Cells[count, 5] = table.Columns.Count;
                    tableOverviewWorksheet.Cells[count, 6] = table.TablePriority;
                    tableOverviewWorksheet.Cells[count, 7] = "";
                    tableOverviewWorksheet.Cells[count, 8] = table.TableEntity;
                    tableOverviewWorksheet.Cells[count, 9] = table.Description;
                    tableOverviewWorksheet.Cells[count, 10] = "";

                    count++;

                    Marshal.ReleaseComObject(c1);
                    Marshal.ReleaseComObject(c2);
                    Marshal.ReleaseComObject(linkCell);
                    Marshal.ReleaseComObject(links);
                }
            }

            tableOverviewWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableOverviewWorksheet.Columns.AutoFit();

            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //*************************************************************************

        // Creates a worksheet with information about the template.
        private void AddTemplateInfo(Worksheet templateSheet, Template template)
        {
            templateSheet.Name = "db";

            /*
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
                */
            List<string> fieldNames = new List<string>()
            {
                "toolVersion",
                "systemSupplier",
                "systemId",
                "systemName",
                "systemVersion",
                "systemInstance",
                "tableCount",
                "",
                "Decom JSON",
                "modelVersion",
                "uuid",
                "name",
                "description",
                "systemName",
                "systemVersion",
                "creator",
                "organizations",
                "creationDate",
                "templateVisibility"
            };

            var prop = template.GetType().GetProperties();

            int count = 1;
            foreach (string s in fieldNames)
            {
                templateSheet.Cells[count, 1] = s;
                count++;

                if (s == "organizations" && template.Organizations != null && template.Organizations.Count > 1)
                {
                    for (int i = 1; i < template.Organizations.Count; i++)
                    {
                        templateSheet.Cells[count, 1] = null;
                        count++;
                    }
                }
            }

            double creationDate = template.CreationDate;

            var date = new DateTime(1970, 1, 1, 0, 0, 0).AddMilliseconds(creationDate).ToLocalTime();

            // toolname
            templateSheet.Cells[1, 2] = Globals.toolName;

            // toolVersion
            templateSheet.Cells[2, 2] = Globals.toolVersion;

            templateSheet.Cells[3, 2] = "";
            templateSheet.Cells[4, 2] = "";
            templateSheet.Cells[5, 2] = "";
            templateSheet.Cells[6, 2] = "";
            templateSheet.Cells[7, 2] = "";

            //tableCount
            templateSheet.Cells[7, 2] = template.TemplateSchema.Tables.Count.ToString();

            templateSheet.Cells[8, 2] = "";

            templateSheet.Cells[9, 2] = "";

            templateSheet.Cells[10, 2] = template.ModelVersion;
            templateSheet.Cells[11, 2] = template.Uuid;
            templateSheet.Cells[12, 2] = template.Name;
            templateSheet.Cells[13, 2] = template.Description;
            templateSheet.Cells[14, 2] = template.SystemName;
            templateSheet.Cells[15, 2] = template.SystemVersion;
            templateSheet.Cells[16, 2] = template.Creator;

            int count2 = 17;
            if (template.Organizations != null)
            {
                foreach (string org in template.Organizations)
                {
                    templateSheet.Cells[count2, 2] = org;
                    count2++;
                }
            }
            else
            {
                templateSheet.Cells[count2, 2] = null;
                count2++;
            }

            templateSheet.Cells[count2, 2] = date;
            templateSheet.Cells[count2 + 1, 2] = template.TemplateVisibility;

            Range range = templateSheet.Cells[2, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            templateSheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            templateSheet.Columns.AutoFit();

            Marshal.ReleaseComObject(templateSheet);
        }

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private static string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private void GetEntity(string description, Table table = null, Column column = null)
        {
            Regex regex = new Regex(@"(?<entity1>(?<=\{)[^\{\}]+(?=\}))|(?<entity2>(?<=\[)[^\[\]]+(?=\]))|(?<desc>\w+)");

            if (description != null)
            {
                string entities = "";
                string cleanDescription = "";

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

                if (table != null)
                    table.TableEntity = entities;
                else if (column != null) 
                    column.Entity = entities;

                if (entities != "")
                {
                    if (table != null)
                        table.Description = cleanDescription;
                    else if (column != null)
                        column.Description = cleanDescription;
                }
            }
        }
    }

    //====================================================================================

    public class Template
    {
        public string ModelVersion { get; set; }
        public string Uuid { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string SystemName { get; set; }
        public string SystemVersion { get; set; }
        public string Creator { get; set; }
        public List<string> Organizations { get; set; }
        public double CreationDate { get; set; }
        public string TemplateVisibility { get; set; }
        public Schema TemplateSchema { get; set; }
    }

    public class Schema
    {
        public Schema(string name, string folder)
        {
            Name = name;
            Folder = folder;
        }

        public List<Table> Tables { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Folder { get; set; }
    }

    public class Table
    {
        public string Name { get; set; }
        public string TablePriority { get; set; }
        public string TableEntity { get; set; }
        public string Folder { get; set; }
        public int Rows { get; set; }
        public string Description { get; set; }
        public List<Column> Columns { get; set; }
        public PrimaryKey PrimaryKey { get; set; }
    }

    public class Column
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Folder { get; set; }
        public string Datatype { get; set; }
        public string Entity { get; set; }
        public string Note { get; set; }
    }

    public class PrimaryKey
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string Description { get; set; }
    }
}