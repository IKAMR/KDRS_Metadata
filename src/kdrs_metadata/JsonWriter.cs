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
            inputTemplate.Organizations.Add(outputSheet.Cells[9, 2].Text);
            //inputTemplate.CreationDate = Convert.ToDouble(outputSheet.Cells[2, 10].Text);
            inputTemplate.TemplateVisibility = outputSheet.Cells[11, 2].Text;

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
