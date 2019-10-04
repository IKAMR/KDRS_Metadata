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

        public delegate void ProgressUpdate(int count, int totalCount, string progressPostfix);
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

            Sheets xlWorksheets;

            xlWorkBooks = xlApp1.Workbooks;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlWorkBooks.Add(misValue);

            xlWorksheets = xlWorkBook.Sheets;

            string json;
            using (StreamReader r = new StreamReader(filename))
            {
                json = r.ReadToEnd();
            }

            Template template = JsonConvert.DeserializeObject<Template>(json);

            Worksheet templateSheet = xlWorksheets.get_Item(1);
            AddDBInfo(templateSheet, template);
            Marshal.ReleaseComObject(templateSheet);

            //Sort tables by priority
            List<string> sortOrder = new List<string> { "HIGH", "MEDIUM", "LOW", "SYSTEM", "STATS", "EMPTY", "DUMMY", null };
            // template.TemplateSchema.Tables.Sort((a, b) => sortOrder.IndexOf(a.TablePriority) - sortOrder.IndexOf(b.TablePriority));

            Worksheet tableOverviewWorksheet = xlWorksheets.Add(After: xlWorksheets[xlWorksheets.Count]);
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

                        Worksheet tableWorksheet = xlWorksheets.Add(After: xlWorksheets[xlWorksheets.Count]);

                        AddTable(tableWorksheet, template.TemplateSchema, table);

                        Marshal.ReleaseComObject(tableWorksheet);
                    }
                    tableCount++;

                    OnProgressUpdate?.Invoke(tableCount, totalTableCount, template.TemplateSchema.Folder +": "+ template.TemplateSchema.Name +"  |  "+ table.Name);
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

            Marshal.ReleaseComObject(xlWorksheets);

            xlWorkBook.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkBook);

            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp1);
        }
        //*************************************************************************
        // Creates a Worksheet with information about the template.
        private void AddDBInfo(Worksheet DBWorksheet, Template template)
        {
            Range tempRng;

            DBWorksheet.Name = "db";
            DBWorksheet.Columns.AutoFit();

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
                DBWorksheet.Cells[count, 1] = s;
                count++;

                if (s == "organizations" && template.Organizations != null && template.Organizations.Count > 1)
                {
                    for (int i = 1; i < template.Organizations.Count; i++)
                    {
                        DBWorksheet.Cells[count, 1] = null;
                        count++;
                    }
                }
            }

            double creationDate = template.CreationDate;

            var date = new DateTime(1970, 1, 1, 0, 0, 0).AddMilliseconds(creationDate).ToLocalTime();

            // toolname
            DBWorksheet.Cells[1, 2] = Globals.toolName;

            // toolVersion
            DBWorksheet.Cells[2, 2] = Globals.toolVersion;

            DBWorksheet.Cells[3, 2] = "";
            DBWorksheet.Cells[4, 2] = "";
            DBWorksheet.Cells[5, 2] = template.SystemName;
            DBWorksheet.Cells[6, 2] = template.SystemVersion;
            DBWorksheet.Cells[7, 2] = "";            

            //tableCount
            DBWorksheet.Cells[8, 2] = template.TemplateSchema.Tables.Count.ToString();

            DBWorksheet.Cells[9, 2] = "";

            DBWorksheet.Cells[10, 2] = "";

            DBWorksheet.Cells[11, 2] = template.ModelVersion;
            DBWorksheet.Cells[12, 2] = template.Uuid;
            DBWorksheet.Cells[13, 2] = template.Name;
            DBWorksheet.Cells[14, 2] = template.Description;
            DBWorksheet.Cells[15, 2] = template.SystemName;
            DBWorksheet.Cells[16, 2] = template.SystemVersion;
            DBWorksheet.Cells[17, 2] = template.Creator;

            int count2 = 18;
            if (template.Organizations != null)
            {
                foreach (string org in template.Organizations)
                {
                    DBWorksheet.Cells[count2, 2] = org;
                    count2++;
                }
            }
            else
            {
                DBWorksheet.Cells[count2, 2] = null;
                count2++;
            }

            DBWorksheet.Cells[count2, 2] = date;
            DBWorksheet.Cells[count2 + 1, 2] = template.TemplateVisibility;

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

            for (int m = 10; m < 21; m++)
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
        //*************************************************************************
        // Creates a Worksheet with table overview
        private void AddTableOverview(Worksheet tableOverviewWorksheet, Schema schema, List<string> priorities)
        {
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
            //-------------------------------------------------------------------
            int count = 2;

            foreach (Table table in schema.Tables)
            {
                Console.WriteLine("Table: " + table.Name + ", Description: " + table.Description);
                GetEntity(table.Description, table);
                Console.WriteLine("Table: " + table.Name + ", Description2: " + table.Description);

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

                    Console.WriteLine("Table priority: " + table.TablePriority + ", Rows: " + table.Rows);
                    if (string.IsNullOrEmpty(table.TablePriority) && table.Rows == 0)
                        table.TablePriority = "EMPTY";
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

            // Freeze Panes
            Range range = tableOverviewWorksheet.Cells[2, 1];
            range.Activate();
            range.Application.ActiveWindow.FreezePanes = true;

            Range tempRng = tableOverviewWorksheet.Range["A1", "I1"];
            tempRng.Characters.Font.Bold = true;

            // Border lines
            for (int m = 1; m < count; m++)
            {
                for (int n = 1; n < 10; n++)
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

            tableOverviewWorksheet.Sort.SortFields.Add(tableOverviewWorksheet.Range["F:F"] , XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, "HIGH, MEDIUM, LOW, SYSTEM, STATS, EMPTY, DUMMY", XlSortDataOption.xlSortNormal);
            tableOverviewWorksheet.Sort.SetRange(tableOverviewWorksheet.UsedRange);
            tableOverviewWorksheet.Sort.Header = XlYesNoGuess.xlYes;
            tableOverviewWorksheet.Sort.Apply();

            Marshal.ReleaseComObject(tableOverviewWorksheet);
        }
        //*************************************************************************
        // Creates a Worksheet with information for each table
        private void AddTable(Worksheet tableWorksheet, Schema schema, Table table)
        {
            Range tempRng;

            Console.WriteLine("Table: " + table.Name);
            tableWorksheet.Name = GetNumbers(table.Folder);
            tableWorksheet.Columns.AutoFit();

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

            int cellCount = 2;

            foreach (string[] rn in rowNamesArray)
            {
                tableWorksheet.Cells[cellCount, 1] = rn;
                tableWorksheet.Cells[cellCount, 2] = rn[1];
                cellCount++;
            }
            // Primary keys
            if (table.PrimaryKey != null)
            {
                tempRng = tableWorksheet.Cells[cellCount, 1];
                tempRng.Interior.Color = Color.LightGray;

                tempRng = tableWorksheet.Cells[cellCount, 2];
                tempRng.Interior.Color = Color.LightGray;

                tableWorksheet.Cells[cellCount, 1] = "pkName";
                tableWorksheet.Cells[cellCount, 2] = table.PrimaryKey.Name;
                cellCount++;

                if (table.PrimaryKey.Columns != null)
                {
                    foreach (string column in table.PrimaryKey.Columns)
                    {
                        tableWorksheet.Cells[cellCount, 1] = "pkColumn";
                        tableWorksheet.Cells[cellCount, 2] = column;
                        cellCount++;
                    }
                }

                // Entity
                tableWorksheet.Cells[cellCount, 1] = "pkEntity";
                string pk_extr_entity = ExtractEntity(table.PrimaryKey.Description, "entity");
                tableWorksheet.Cells[cellCount, 2] = pk_extr_entity;

                for (int n = 1; n < 8; n++)
                {
                    tempRng = tableWorksheet.Cells[cellCount, n];
                    tempRng.Interior.Color = Color.LightGreen;
                }
                cellCount++;

                // Description
                tableWorksheet.Cells[cellCount, 1] = "pkDescription";
                string pk_extr_description = ExtractEntity(table.PrimaryKey.Description, "description");
                tableWorksheet.Cells[cellCount, 2] = pk_extr_description;

                for (int n = 1; n < 8; n++)
                {
                    tempRng = tableWorksheet.Cells[cellCount, n];
                    tempRng.Interior.Color = Color.LightSkyBlue;
                }
                cellCount++;
            }
            Console.WriteLine("fKEYS");

            // Foreign keys
            if (table.ForeignKeys != null)
            {
                foreach (ForeignKey fkey in table.ForeignKeys)
                {
                    tempRng = tableWorksheet.Cells[cellCount, 1];
                    tempRng.Interior.Color = Color.LightPink;

                    tempRng = tableWorksheet.Cells[cellCount, 2];
                    tempRng.Interior.Color = Color.LightPink;

                    tableWorksheet.Cells[cellCount, 1] = "fkName";
                    tableWorksheet.Cells[cellCount, 2] = fkey.Name;
                    cellCount++;

                    if (fkey.Columns != null)
                    {
                        foreach (string column in fkey.Columns)
                        {
                            tableWorksheet.Cells[cellCount, 1] = "fkColumn";
                            tableWorksheet.Cells[cellCount, 2] = column;
                            cellCount++;
                        }
                    }
                    
                    tableWorksheet.Cells[cellCount, 1] = "fkRefSchema";
                    tableWorksheet.Cells[cellCount, 2] = fkey.ReferencedSchema;
                    cellCount++;
                    
                    tableWorksheet.Cells[cellCount, 1] = "fkRefTable";
                    tableWorksheet.Cells[cellCount, 2] = fkey.ReferencedTable;
                    cellCount++;

                    if (fkey.ReferencedColumns != null)
                    {
                        foreach (string column in fkey.ReferencedColumns)
                        {
                            tableWorksheet.Cells[cellCount, 1] = "fkReferencedColumns";
                            tableWorksheet.Cells[cellCount, 2] = column;
                            cellCount++;
                        }
                    }

                    // Entity
                    tableWorksheet.Cells[cellCount, 1] = "fkEntity";
                    string fk_extr_entity = ExtractEntity(fkey.Description, "entity");
                    tableWorksheet.Cells[cellCount, 2] = fk_extr_entity;

                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightGreen;
                    }
                    cellCount++;

                    // Description
                    tableWorksheet.Cells[cellCount, 1] = "fkDescription";
                    string fk_extr_description = ExtractEntity(fkey.Description, "description");
                    tableWorksheet.Cells[cellCount, 2] = fk_extr_description;

                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightSkyBlue;
                    }
                    cellCount++;
                    
                    tableWorksheet.Cells[cellCount, 1] = "fkDeleteAction";
                    tableWorksheet.Cells[cellCount, 2] = fkey.DeleteAction;
                    cellCount++;

                    tableWorksheet.Cells[cellCount, 1] = "fkUpdateAction";
                    tableWorksheet.Cells[cellCount, 2] = fkey.UpdateAction;
                    cellCount++;
                }
            }
            Console.WriteLine("cKEYS");

            // Candidate keys
            if (table.CandidateKeys != null)
            {
                foreach (CandidateKey ckey in table.CandidateKeys)
                {
                    tempRng = tableWorksheet.Cells[cellCount, 1];
                    tempRng.Interior.Color = Color.PaleTurquoise;

                    tempRng = tableWorksheet.Cells[cellCount, 2];
                    tempRng.Interior.Color = Color.PaleTurquoise;

                    tableWorksheet.Cells[cellCount, 1] = "ckName";
                    tableWorksheet.Cells[cellCount, 2] = ckey.Name;
                    cellCount++;

                    if (ckey.Columns != null)
                    {
                        foreach (string column in ckey.Columns)
                        {
                            tableWorksheet.Cells[cellCount, 1] = "ckColumn";
                            tableWorksheet.Cells[cellCount, 2] = column;
                            cellCount++;
                        }
                    }

                    // Entity
                    tableWorksheet.Cells[cellCount, 1] = "ckEntity";
                    string ck_extr_entity = ExtractEntity(ckey.Description, "entity");
                    tableWorksheet.Cells[cellCount, 2] = ck_extr_entity;

                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightGreen;
                    }
                    cellCount++;

                    // Description
                    tableWorksheet.Cells[cellCount, 1] = "ckDescription";
                    string ck_extr_description = ExtractEntity(ckey.Description, "description");
                    tableWorksheet.Cells[cellCount, 2] = ck_extr_description;

                    for (int n = 1; n < 8; n++)
                    {
                        tempRng = tableWorksheet.Cells[cellCount, n];
                        tempRng.Interior.Color = Color.LightSkyBlue;
                    }
                    cellCount++;
                }
            }

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
            int columnCount = 1;
            foreach (Column column in table.Columns)
            {
                GetEntity(column.Description, null, column);

                tableWorksheet.Cells[cellCount, 1] = columnCount;
                tableWorksheet.Cells[cellCount, 2] = column.Name;
                tableWorksheet.Cells[cellCount, 3] = column.Datatype;

                //typeOriginal
                tableWorksheet.Cells[cellCount, 4] = "";

                tableWorksheet.Cells[cellCount, 5] = column.Nullable;
                
                //defaultValue
                tableWorksheet.Cells[cellCount, 6] = "";

                tableWorksheet.Cells[cellCount, 7] = column.Folder;
                tableWorksheet.Cells[cellCount, 8] = column.Entity;
                tableWorksheet.Cells[cellCount, 9] = column.Description;

                //note
                tableWorksheet.Cells[cellCount, 10] = "";                

                // Border line
                for (int n = 1; n < 11; n++)
                {
                    tempRng = tableWorksheet.Cells[cellCount, n];
                    tempRng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    tempRng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }

                // Console.WriteLine("lobFolder = '" + column.Folder + "'");
                if (null != column.Folder && "" != column.Folder )
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
                columnCount++;
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
        public int rowsCount()
        {
            int rows = 0;
            foreach(Table table in Tables)
            {
                rows += table.Rows;
            }
            return rows;
        }

        // ToDo 2019-09-24 TFA TAA: Restructure Table as class for XML (DataConverter.cs)
        // Now QuickFix parse add counters in the PK/FK/CK loops JSON & XML

        /* public int countPK()
        {
            int countPK = 0;
            foreach (Table table in Tables)
            {
                if (table.PrimaryKey != null)
                    countPK++;
            }
            return countPK;
        }
        public int countFK()
        {
            int countFK = 0;
            foreach (Table table in Tables)
            {
                if (table.ForeignKeys != null)
                    countFK += table.ForeignKeys.Count;
            }
            return countFK;
        }
        public int countCK()
        {
            int countCK = 0;
            foreach (Table table in Tables)
            {
                if (table.CandidateKeys != null)
                    countCK += table.CandidateKeys.Count;
            }
            return countCK;
        } */

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