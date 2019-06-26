using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace KDRS_Metadata
{
    class JsonReader

    {

        public void ParseJson(string filename, List<string> priorities)
        {

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

            foreach (Table table in template.TemplateSchema.Tables)
            {
                if (priorities.Contains(table.TablePriority))
                {

                    Worksheet tableWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);

                    AddTable(tableWorksheet, template.TemplateSchema, table);

                    Marshal.ReleaseComObject(tableWorksheet);

                }
            }

            xlWorkBook.Sheets[1].Select();

            xlWorkBook.SaveAs(Path.ChangeExtension(Path.GetFullPath(filename), ".xlsx"));

            Marshal.ReleaseComObject(xlWorkSheets);

            xlWorkBook.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkBook);

            xlWorkBook = null;

            xlApp1.Quit();
            
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp1);

            xlApp1 = null;
            Console.WriteLine("App: " + xlApp1);
        }

        //*************************************************************************

        // Creates a worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, Schema schema, Table table)
        {

            tableWorksheet.Name = table.Name;

            List<string> columnNames = new List<string>()
            {
                "Column",
                "Name",
                "Type",
                "Folder",
                "Description",
                "Note"
            };

            foreach (string name in columnNames)
            {
                tableWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
            }
            //-------------------------------------------------------------------
            string[][] rowNamesArray = new string[8][] {
                new string[2] { "schemaName", schema.Name },
                new string[2] { "schemaFolder", schema.Folder},
                new string[2] { "tableName", table.Name },
                new string[2] { "tableFolder", table.Folder },
                new string[2] { "tablePriority", table.TablePriority },
                new string[2] { "tableDescription", table.Description },
                new string[2] { "rows", table.Rows.ToString() },
                new string[2] { "columns", table.Columns.Count().ToString() },
            };

            int count = 2;

            foreach (string[] rn in rowNamesArray)
            {
                if (rn[0] == "tableDescription")
                {
                    tableWorksheet.Cells[count, 1] = rn;
                    tableWorksheet.Cells[count, 3] = rn[1];
                }
                else
                {
                    tableWorksheet.Cells[count, 1] = rn;
                    tableWorksheet.Cells[count, 2] = rn[1];
                }

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
                tableWorksheet.Cells[count, 1] = "Column " + columnCount;
                tableWorksheet.Cells[count, 2] = column.Name;
                tableWorksheet.Cells[count, 3] = column.Datatype;
                tableWorksheet.Cells[count, 4] = column.Folder;
                tableWorksheet.Cells[count, 5] = column.Description;
                count++;

                columnCount++;
            }

            tableWorksheet.Columns.AutoFit();
            Marshal.ReleaseComObject(tableWorksheet);
        }

        //*************************************************************************

        private void AddTableOverview(Worksheet tableOverviewWorksheet, Schema schema, List<string> priorities)
        {
            tableOverviewWorksheet.Name = "Tables";

            List<string> columnNames = new List<string>()
            {
                "Table",
                "Folder",
                "Schema",
                "Rows",
                "Priorities"
            };

            foreach (string name in columnNames)
            {
                tableOverviewWorksheet.Cells[1, columnNames.IndexOf(name) + 1] = name;
            }
            //-------------------------------------------------------------------
            int count = 2;
            foreach (Table table in schema.Tables)
            {
                if (priorities.Contains(table.TablePriority))
                {
                    Range c1 = tableOverviewWorksheet.Cells[count, 1];
                    Range c2 = tableOverviewWorksheet.Cells[count, 1];
                    Range linkCell = tableOverviewWorksheet.get_Range(c1, c2);

                    Hyperlinks links = tableOverviewWorksheet.Hyperlinks;
                    links.Add(linkCell, "", table.Name + "!A1", "", table.Name);

                    //tableOverviewWorksheet.Cells[count, 1] = table.Name;
                    tableOverviewWorksheet.Cells[count, 2] = table.Folder;
                    tableOverviewWorksheet.Cells[count, 3] = schema.Name;
                    tableOverviewWorksheet.Cells[count, 4] = table.Rows;
                    tableOverviewWorksheet.Cells[count, 5] = table.TablePriority;

                    count++;

                    Marshal.ReleaseComObject(c1);
                    Marshal.ReleaseComObject(c2);
                    Marshal.ReleaseComObject(linkCell);
                    Marshal.ReleaseComObject(links);
                }
            }

            tableOverviewWorksheet.Columns.AutoFit();
            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //*************************************************************************

        // Creates a worksheet with information about the template.
        private void AddTemplateInfo(Worksheet templateSheet, Template template)
        {
            templateSheet.Name = "Template";

            List<string> fieldNames = new List<string>()
            {
                "modelVersion",
                "uuid",
                "name",
                "description",
                "systemName",
                "systemVersion",
                "creator",
                "organizations",
                "creationDate",
                "templateVisibility",
                "Table count"
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

            templateSheet.Cells[1, 2] = template.ModelVersion;
            templateSheet.Cells[2, 2] = template.Uuid;
            templateSheet.Cells[3, 2] = template.Name;
            templateSheet.Cells[4, 2] = template.Description;
            templateSheet.Cells[5, 2] = template.SystemName;
            templateSheet.Cells[6, 2] = template.SystemVersion;
            templateSheet.Cells[7, 2] = template.Creator;

            int count2 = 8;
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
            templateSheet.Cells[count2 + 2, 2] = template.TemplateSchema.Tables.Count.ToString();

            templateSheet.Columns.AutoFit();

            Marshal.ReleaseComObject(templateSheet);
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
        public List<Table> Tables { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Folder { get; set; }
    }

    public class Table
    {
        public string Name { get; set; }
        public string TablePriority { get; set; }
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
    }

    public class PrimaryKey
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string Description { get; set; }
    }
}
