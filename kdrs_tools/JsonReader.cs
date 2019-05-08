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


namespace Metadata_XLS
{
    class JsonReader

    {
        public void ParseJson(string filename, List<string> priorities)
        {
            Microsoft.Office.Interop.Excel.Application xlApp1 = new Microsoft.Office.Interop.Excel.Application();

            Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp1.Workbooks.Add(misValue);

            string json;
            using (StreamReader r = new StreamReader(filename))
            {
                json = r.ReadToEnd();
            }

            Template template = JsonConvert.DeserializeObject<Template>(json);
          
            AddTemplateInfo(xlWorkBook, template);

            AddTableOverview(xlApp1, xlWorkBook, template.TemplateSchema, priorities);

            Console.WriteLine("After");
            foreach (string l in priorities)
            {
                Console.WriteLine(l);
            }

            foreach (Table table in template.TemplateSchema.Tables)
            {
                if (priorities.Contains(table.TablePriority))
                {
                    //Console.WriteLine("Tabell: {0} , prioritet {1}" ,table.Name, table.TablePriority);
                    AddTable(xlWorkBook, template.TemplateSchema, table);

                }
            }

            xlWorkBook.Sheets[1].Select();

            xlWorkBook.SaveAs(Path.ChangeExtension(Path.GetFullPath(filename), ".xlsx"));

            xlWorkBook.Close(true, misValue, misValue);
            xlApp1.Quit();


            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp1);

        }

        //*************************************************************************

        // Creates a worksheet with information for each table
        private void AddTable(Workbook workbook, Schema schema, Table table)
        {
            Worksheet tableWorksheet;
            tableWorksheet = (Worksheet)workbook.Application.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
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

            Marshal.ReleaseComObject(tableWorksheet);
        }

        //*************************************************************************

        private void AddTableOverview(Application excelApp, Workbook workbook, Schema schema, List<string> priorities)
        {
            Worksheet tableOverviewWorksheet = (Worksheet)workbook.Application.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
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
                    Range linkCell = excelApp.get_Range(c1, c2);
                    tableOverviewWorksheet.Hyperlinks.Add(linkCell, "", table.Name + "!A1", "", table.Name);

                    //tableOverviewWorksheet.Cells[count, 1] = table.Name;
                    tableOverviewWorksheet.Cells[count, 2] = table.Folder;
                    tableOverviewWorksheet.Cells[count, 3] = schema.Name;
                    tableOverviewWorksheet.Cells[count, 4] = table.Rows;
                    tableOverviewWorksheet.Cells[count, 5] = table.TablePriority;

                    count++;
                }
            }

            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //*************************************************************************

        // Creates a worksheet with information about the template.
        private void AddTemplateInfo(Workbook workbook, Template template)
        {
            Worksheet templateSheet = (Worksheet)workbook.Worksheets.get_Item(1);
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
            }

            templateSheet.Cells[1, 2] = template.ModelVersion;
            templateSheet.Cells[2, 2] = template.Uuid;
            templateSheet.Cells[3, 2] = template.Name;
            templateSheet.Cells[4, 2] = template.Description;
            templateSheet.Cells[5, 2] = template.SystemName;
            templateSheet.Cells[6, 2] = template.SystemVersion;
            templateSheet.Cells[7, 2] = template.Creator;
            templateSheet.Cells[8, 2] = template.Organizations;
            templateSheet.Cells[9, 2] = template.CreationDate;
            templateSheet.Cells[10, 2] = template.TemplateVisibility;
            templateSheet.Cells[11, 2] = template.TemplateSchema.Tables.Count.ToString();

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
        public string Organizations { get; set; }
        public string CreationDate { get; set; }
        public string TemplateVisibility { get; set; }
        public Schema TemplateSchema { get; set; }
        // public string TablePriority { get; set; }
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
        //public Columns columns { get; set; }
    }

    public class Column
    {

        public string Name { get; set; }
        public string Description { get; set; }
        public string Folder { get; set; }
        public string Datatype { get; set; }
        // public string TablePriority { get; set; }
    }

    public class PrimaryKey
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string Description { get; set; }
    }
}
