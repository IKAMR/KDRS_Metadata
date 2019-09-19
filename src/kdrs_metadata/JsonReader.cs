using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
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

        public List<Schema> schemaNames = new List<Schema>();

        public bool includeTables;

        public delegate void ProgressUpdate(int count, int totalCount);
        public event ProgressUpdate OnProgressUpdate;

        public void ParseJson(string filename, List<string> priorities, bool includeTables)
        {
            schemaNames.Clear();

            this.includeTables = includeTables;

            Application xlApp1 = new Application
            {
                DecimalSeparator = ".",
                UseSystemSeparators = false
            };

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
            AddDBInfo(templateSheet, template);
            Marshal.ReleaseComObject(templateSheet);

            //Sort tables by priority
            List<string> sortOrder = new List<string> { "HIGH", "MEDIUM", "LOW", "SYSTEM", "STATS", "EMPTY", "DUMMY", null };
            template.TemplateSchema.Tables.Sort((a, b) => sortOrder.IndexOf(a.TablePriority) - sortOrder.IndexOf(b.TablePriority));

            Worksheet tableOverviewWorksheet = xlWorkSheets.Add(After: xlWorkSheets[xlWorkSheets.Count]);
            AddTableOverview(tableOverviewWorksheet, template.TemplateSchema, priorities);
            Marshal.ReleaseComObject(tableOverviewWorksheet);

            schemaNames.Add(template.TemplateSchema);

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

            int fileCounter = 1;

            excelFileName = checkFileName(excelFileName, fileCounter);
            xlWorkBook.SaveAs(excelFileName);

            Marshal.ReleaseComObject(xlWorkSheets);

            xlWorkBook.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkBook);

            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp1);
        }
        //*************************************************************************
        // Creates a worksheet with information about the template.
        private void AddDBInfo(Worksheet DBWorkSheet, Template template)
        {
            DBWorkSheet.Name = "db";

            List<string> fieldNames = new List<string>()
            {
                "toolname",
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
                DBWorkSheet.Cells[count, 1] = s;
                count++;

                if (s == "organizations" && template.Organizations != null && template.Organizations.Count > 1)
                {
                    for (int i = 1; i < template.Organizations.Count; i++)
                    {
                        DBWorkSheet.Cells[count, 1] = null;
                        count++;
                    }
                }
            }

            double creationDate = template.CreationDate;

            var date = new DateTime(1970, 1, 1, 0, 0, 0).AddMilliseconds(creationDate).ToLocalTime();

            // toolname
            DBWorkSheet.Cells[1, 2] = Globals.toolName;

            // toolVersion
            DBWorkSheet.Cells[2, 2] = Globals.toolVersion;

            DBWorkSheet.Cells[3, 2] = "";
            DBWorkSheet.Cells[4, 2] = "";
            DBWorkSheet.Cells[5, 2] = template.SystemName;
            DBWorkSheet.Cells[6, 2] = template.SystemVersion;
            DBWorkSheet.Cells[7, 2] = "";

            Range tempRng = DBWorkSheet.Range["A1", "B1"];
            tempRng.Characters.Font.Bold = true;

            Range redColorRng = DBWorkSheet.Range["A3", "C6"];
            redColorRng.Characters.Font.Color = Color.Red;

            Range orangeColorRng = DBWorkSheet.Range["A7", "C7"];
            orangeColorRng.Characters.Font.Color = Color.Orange;

            //tableCount
            DBWorkSheet.Cells[8, 2] = template.TemplateSchema.Tables.Count.ToString();

            DBWorkSheet.Cells[9, 2] = "";

            DBWorkSheet.Cells[10, 2] = "";

            DBWorkSheet.Cells[11, 2] = template.ModelVersion;
            DBWorkSheet.Cells[12, 2] = template.Uuid;
            DBWorkSheet.Cells[13, 2] = template.Name;
            DBWorkSheet.Cells[14, 2] = template.Description;
            DBWorkSheet.Cells[15, 2] = template.SystemName;
            DBWorkSheet.Cells[16, 2] = template.SystemVersion;
            DBWorkSheet.Cells[17, 2] = template.Creator;

            int count2 = 18;
            if (template.Organizations != null)
            {
                foreach (string org in template.Organizations)
                {
                    DBWorkSheet.Cells[count2, 2] = org;
                    count2++;
                }
            }
            else
            {
                DBWorkSheet.Cells[count2, 2] = null;
                count2++;
            }

            DBWorkSheet.Cells[count2, 2] = date;
            DBWorkSheet.Cells[count2 + 1, 2] = template.TemplateVisibility;


            DBWorkSheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            DBWorkSheet.Columns.AutoFit();

            Marshal.ReleaseComObject(DBWorkSheet);
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
                // "pri-sort",
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

                if (priorities.Contains(table.TablePriority))
                {
                    if (includeTables)
                    {
                        Range c1 = tableOverviewWorksheet.Cells[count, 1];
                        Range c2 = tableOverviewWorksheet.Cells[count, 1];
                        Range linkCell = tableOverviewWorksheet.get_Range(c1, c2);

                        Hyperlinks links = tableOverviewWorksheet.Hyperlinks;
                        links.Add(linkCell, "", GetNumbers(table.Folder) + "!A1", "", table.Name);

                        Marshal.ReleaseComObject(c1);
                        Marshal.ReleaseComObject(c2);
                        Marshal.ReleaseComObject(linkCell);
                        Marshal.ReleaseComObject(links);
                    }
                    else
                    {
                        tableOverviewWorksheet.Cells[count, 1] = table.Name;
                    }

                    tableOverviewWorksheet.Cells[count, 2] = table.Folder;
                    tableOverviewWorksheet.Cells[count, 3] = schema.Name;
                    tableOverviewWorksheet.Cells[count, 4] = table.Rows;
                    tableOverviewWorksheet.Cells[count, 5] = table.Columns.Count;
                    tableOverviewWorksheet.Cells[count, 6] = table.TablePriority;

                    // Pri sort
                    //tableOverviewWorksheet.Cells[count, 7] = Globals.PriSort(table.TablePriority);
                    tableOverviewWorksheet.Cells[count, 7] = table.TableEntity;
                    tableOverviewWorksheet.Cells[count, 8] = table.Description;

                    // Note
                    tableOverviewWorksheet.Cells[count, 9] = "";

                    count++;
                }
            }

            Range range = tableOverviewWorksheet.Cells[2, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            Range tempRng = tableOverviewWorksheet.Range["A1", "I1"];
            tempRng.Characters.Font.Bold = true;

            tableOverviewWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableOverviewWorksheet.Columns.AutoFit();

            tableOverviewWorksheet.Columns["B:B"].ColumnWidth = 8;
            tableOverviewWorksheet.Columns["C:C"].ColumnWidth = 8;
            tableOverviewWorksheet.Columns["F:F"].ColumnWidth = 8;

            tableOverviewWorksheet.Columns["G:G"].ColumnWidth = 14;

            tableOverviewWorksheet.Columns["H:H"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["H:H"].WrapText = true;

            tableOverviewWorksheet.Columns["I:I"].ColumnWidth = 60;
            tableOverviewWorksheet.Columns["I:I"].WrapText = true;
            /*
            tableOverviewWorksheet.Sort.SortFields.Clear();

            tableOverviewWorksheet.Sort.SortFields.Add(tableOverviewWorksheet.Range["F:F"] , XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, "HIGH, MEDIUM, LOW, SYSTEM, STATS, EMPTY, DUMMY", XlSortDataOption.xlSortNormal);
            tableOverviewWorksheet.Sort.SetRange(tableOverviewWorksheet.UsedRange);
            tableOverviewWorksheet.Sort.Header = XlYesNoGuess.xlYes;
           // tableOverviewWorksheet.Sort.Apply();
           */

            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //*************************************************************************
        // Creates a worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, Schema schema, Table table)
        {
            Console.WriteLine("Table: " + table.Name);
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
                tableWorksheet.Cells[count, 1] = rn;
                tableWorksheet.Cells[count, 2] = rn[1];
                count++;
            }
            // Primary keys
            if (table.PrimaryKey != null)
            {
                tableWorksheet.Cells[count, 1] = "pkName";
                tableWorksheet.Cells[count, 2] = table.PrimaryKey.Name;
                count++;

                if (table.PrimaryKey.Columns != null)
                {
                    foreach (string column in table.PrimaryKey.Columns)
                    {
                        tableWorksheet.Cells[count, 1] = "pkColumn";
                        tableWorksheet.Cells[count, 2] = column;
                        count++;
                    }
                }

                tableWorksheet.Cells[count, 1] = "pkDescription";
                tableWorksheet.Cells[count, 2] = table.PrimaryKey.Description;
                count++;
            }
            Console.WriteLine("fKEYS");

            // Foreign keys
            if (table.ForeignKeys != null)
            {
                foreach (ForeignKey fkey in table.ForeignKeys)
                {
                    tableWorksheet.Cells[count, 1] = "fkName";
                    tableWorksheet.Cells[count, 2] = fkey.Name;
                    count++;

                    if (fkey.Columns != null)
                    {
                        foreach (string column in fkey.Columns)
                        {
                            tableWorksheet.Cells[count, 1] = "fkColumn";
                            tableWorksheet.Cells[count, 2] = column;
                            count++;
                        }
                    }
                    
                    tableWorksheet.Cells[count, 1] = "fkRefSchema";
                    tableWorksheet.Cells[count, 2] = fkey.ReferencedSchema;
                    count++;
                    
                    tableWorksheet.Cells[count, 1] = "fkRefTable";
                    tableWorksheet.Cells[count, 2] = fkey.ReferencedTable;
                    count++;

                    if (fkey.ReferencedColumns != null)
                    {
                        foreach (string column in fkey.ReferencedColumns)
                        {
                            tableWorksheet.Cells[count, 1] = "fkReferencedColumns";
                            tableWorksheet.Cells[count, 2] = column;
                            count++;
                        }
                    }
                    
                    tableWorksheet.Cells[count, 1] = "fkDescription";
                    tableWorksheet.Cells[count, 2] = fkey.Description;
                    count++;
                    
                    tableWorksheet.Cells[count, 1] = "fkDeleteAction";
                    tableWorksheet.Cells[count, 2] = fkey.DeleteAction;
                    count++;

                    tableWorksheet.Cells[count, 1] = "fkUpdateAction";
                    tableWorksheet.Cells[count, 2] = fkey.UpdateAction;
                    count++;
                }
            }
            Console.WriteLine("cKEYS");

            // Candidate keys
            if (table.CandidateKeys != null)
            {
                foreach (CandidateKey ckey in table.CandidateKeys)
                {
                    tableWorksheet.Cells[count, 1] = "ckName";
                    tableWorksheet.Cells[count, 2] = ckey.Name;
                    count++;

                    if (ckey.Columns != null)
                    {
                        foreach (string column in ckey.Columns)
                        {
                            tableWorksheet.Cells[count, 1] = "ckColumn";
                            tableWorksheet.Cells[count, 2] = column;
                            count++;
                        }
                    }

                    tableWorksheet.Cells[count, 1] = "ckDescription";
                    tableWorksheet.Cells[count, 2] = ckey.Description;
                    count++;
                }
            }

            // Columns
            int columnCount = 1;
            foreach (Column column in table.Columns)
            {
                GetEntity(column.Description, null, column);

                tableWorksheet.Cells[count, 1] = columnCount;
                tableWorksheet.Cells[count, 2] = column.Name;
                tableWorksheet.Cells[count, 3] = column.Datatype;

                //typeOriginal
                tableWorksheet.Cells[count, 4] = "";

                tableWorksheet.Cells[count, 5] = column.Nullable;
                
                //defaultValue
                tableWorksheet.Cells[count, 6] = "";

                tableWorksheet.Cells[count, 7] = column.Folder;
                tableWorksheet.Cells[count, 8] = column.Entity;
                tableWorksheet.Cells[count, 9] = column.Description;

                //note
                tableWorksheet.Cells[count, 10] = "";
                count++;

                columnCount++;
            }

            // Range range = tableWorksheet.Cells[5, 1];
            Range range = tableWorksheet.Cells[9, 3];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            Range tempRng = tableWorksheet.Range["A1", "J1"];
            tempRng.Characters.Font.Bold = true;

            tableWorksheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            tableWorksheet.Columns.AutoFit();

            tableWorksheet.Columns["B:B"].ColumnWidth = 30;

            tableWorksheet.Columns["I:I"].ColumnWidth = 60;
            tableWorksheet.Columns["I:I"].WrapText = true;

            tableWorksheet.Columns["J:J"].ColumnWidth = 60;
            tableWorksheet.Columns["J:J"].WrapText = true;

            Marshal.ReleaseComObject(tableWorksheet);
        }
        

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private static string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Extracts entities from table dn column description.
        private void GetEntity(string description, Table table = null, Column column = null)
        {
            Regex regex = new Regex(@"(?<entity1>(?<=\{)[^\{\}]+(?=\}))|(?<entity2>(?<=\[)[^\[\]]+(?=\]))|(?<desc>\w+.*)");

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
        public List<ForeignKey> ForeignKeys { get; set; }
        public List<CandidateKey> CandidateKeys { get; set; }
    }

    public class Column
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Nullable { get; set; }
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

    public class ForeignKey
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string Description { get; set; }
        public string ReferencedSchema { get; set; }
        public string ReferencedTable { get; set; }
        public List<string> ReferencedColumns { get; set; }
        public string DeleteAction { get; set; }
        public string UpdateAction { get; set; }
    }

    public class CandidateKey
    {
        public string Name { get; set; }
        public List<string> Columns { get; set; }
        public string Description { get; set; }
    }
}