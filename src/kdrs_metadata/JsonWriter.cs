using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDRS_Metadata
{
    class JsonTemplateWriter
    {
       // Template newTemplate = new Template
       // {
       //     Name = "templateName",
       //     //TemplateSchema = new Schema {
       //         Name = "schemaName"
       //     }
       // };

        public void ReadXlsx(string XlsFileName)
        {
            Console.WriteLine("Reading xlsx");

            Application xlApp1 = new Application();
            Workbooks xlWorkBooks = xlApp1.Workbooks;
            Workbook xlWorkBook = xlWorkBooks.Open(XlsFileName);

            Sheets xlWorksheets = xlWorkBook.Worksheets;

            Worksheet outputSheet = xlWorksheets["output"];

            Template inputTemplate = new Template();

            inputTemplate.ModelVersion = outputSheet.Cells[2, 2].Text;
            inputTemplate.Uuid = outputSheet.Cells[3, 2].Text;
            inputTemplate.Name = outputSheet.Cells[4, 2].Text;
            inputTemplate.Description = outputSheet.Cells[5, 2].Text;
            inputTemplate.SystemName = outputSheet.Cells[6, 2].Text;
            inputTemplate.SystemVersion = outputSheet.Cells[7, 2].Text;
            inputTemplate.Creator = outputSheet.Cells[8, 2].Text;
            inputTemplate.Organizations = new List<string>();

            int counter = 9;
            while ("creationDate" != outputSheet.Cells[counter, 1].Text)
            {
                inputTemplate.Organizations.Add(outputSheet.Cells[counter, 2].Text);
                counter++;
            }
            //inputTemplate.CreationDate = outputSheet.Cells[10, 2].Value;
            counter++;

            inputTemplate.TemplateVisibility = outputSheet.Cells[counter, 2].Text;
            counter++;

            Worksheet tablesSheet = xlWorksheets["tables"];

            Range column = tablesSheet.UsedRange.Columns["C:C", Type.Missing].Cells;
            int tableCount = column.Count - 1;

            // Get distinct schemaNames
            HashSet<string> schemaNames = new HashSet<string>();
            bool firstRow = true;
            foreach (Range row in column)
            {
                if (firstRow)
                    firstRow = false;
                else
                    schemaNames.Add(row.Text);
            }

            // Make list of schemas
            inputTemplate.TemplateSchemaList = new List<Schema>();
            foreach (string name in schemaNames)
            {
                Schema tableSchema = new Schema(name, "")
                {
                    Tables = new List<Table>()
                };

                for (int i = 4; i <= tableCount; i++)
                {
                    Worksheet tableSheet = xlWorksheets[i];

                    if (tableSheet.Cells[2, 2].Text != tableSchema.Name)
                        break;
                    else
                    {
                        tableSchema.Folder = tableSheet.Cells[3, 2].Text;
                        tableSchema.Tables.Add(new Table
                        {
                            Name = tableSheet.Cells[4, 2].Text,
                            Folder = tableSheet.Cells[5, 2].Text,
                            TablePriority = tableSheet.Cells[6, 2].Text,
                            TableEntity = tableSheet.Cells[7, 2].Text,
                            Description = tableSheet.Cells[8, 2].Text,
                            //Rows = tableSheet.Cells[9, 2].Text,
                            PrimaryKey = new PrimaryKey
                            {
                                Name = tableSheet.Cells[11, 2].Text,
                                //Columns = "",
                                Description = tableSheet.Cells[3, 2].Text

                            }
                        });

                    }

                }

                inputTemplate.TemplateSchemaList.Add(tableSchema);
            }



            Console.WriteLine("Input tamplate neme: " + inputTemplate.Name);

            JsonSerializer serializer = new JsonSerializer();

            using (StreamWriter sw = new StreamWriter(@"Y:\developer\debug\KDRS_Metadata\v0.9.5-rc1\outputTemplate.json"))
                using (JsonWriter writer = new JsonTextWriter(sw))
            {
                writer.Formatting = Formatting.Indented;
                serializer.Serialize(writer, inputTemplate);
            }
        }
    }
}
