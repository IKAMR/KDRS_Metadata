﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;

namespace KDRS_Metadata
{
    public static class Globals
    {
        public static readonly String toolName = "KDRS-Metadata";
        public static readonly String toolVersion = "0.5";
    }

    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application xlApp;

        DataConverter converter = new DataConverter();
        JsonReader jsonReader = new JsonReader();

        List<string> priorities = new List<string> { };

        Hashtable myHashtable;

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.DragEnter += new DragEventHandler(Form1_DragEnter);

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel er ikke installert!!");
                return;
            }
            else
            {
                Console.WriteLine("Excel Ok!");
            }

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            CheckExcellProcesses();
            string fileName = "No file added";

            try
            {
                CheckPrioList();
                label1.Text = "";
                label2.Text = "";
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Count() > 1)
                    label1.Text = "Vennligst bare en fil av gangen... ;D";
                else
                {
                    fileName = files[0].ToString();

                    string filType = Path.GetExtension(fileName);
                    Console.WriteLine(filType);
                    switch (filType)
                    {
                        case ".json":
                            label1.Text = "Converting " + fileName;

                            jsonReader.ParseJson(fileName, priorities);
                            label2.Text = "Number of tables: " + jsonReader.tableCount;

                            break;
                        case ".xml":
                            label1.Text = "Converting " + fileName;
                            converter.Convert(fileName, includeTables.Checked);
                            label2.Text = converter.schemaName + "\n" + converter.totalTableCount;
                            break;
                    }

                    label1.Text = "Job complete!";

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                label2.Text = "Error: " + ex.Message + "\n" + fileName;
                KillExcel();
            }
            finally
            {

                KillExcel();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        //----------------------------------------------------------------------------------------------

        private void CheckPrioList()
        {
            //"HIGH", "MEDIUM", "LOW", "SYSTEM", "EMPTY", null
            if (priorityHigh.Checked)
            {
                priorities.Add("HIGH");
                Console.WriteLine("High checked");
            }
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("HIGH");
            }

            if (priorityMedium.Checked)
                priorities.Add("MEDIUM");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("MEDIUM");
            }

            if (priorityLow.Checked)
                priorities.Add("LOW");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("LOW");
            }

            if (prioritySystem.Checked)
                priorities.Add("SYSTEM");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("SYSTEM");
            }

            if (priorityStat.Checked)
                priorities.Add("STAT");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("STAT");
            }

            if (priorityDummy.Checked)
                priorities.Add("DUMMY");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("DUMMY");
            }

            if (priorityEmpty.Checked)
                priorities.Add("EMPTY");
            else if (!priorityHigh.Checked)
            {
                priorities.Remove("EMPTY");
            }

            if (priorityNull.Checked)
                priorities.Add(null);
            else if (!priorityHigh.Checked)
            {
                priorities.Remove(null);
            }
        }
        //----------------------------------------------------------------------------------------------

        private void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");

            // check to kill the right process
            foreach (Process ExcelProcess in AllProcesses)
            {
                if (myHashtable.ContainsKey(ExcelProcess.Id) == false)
                    ExcelProcess.Kill();
            }

            AllProcesses = null;
        }
        //----------------------------------------------------------------------------------------------

        private void CheckExcellProcesses()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");
            myHashtable = new Hashtable();
            int iCount = 0;

            foreach (Process ExcelProcess in AllProcesses)
            {
                myHashtable.Add(ExcelProcess.Id, iCount);
                iCount = iCount + 1;
            }
        }
        //----------------------------------------------------------------------------------------------


        private void Form1_Load(object sender, EventArgs e)
        {

        }
        //----------------------------------------------------------------------------------------------


        /*
        public System.Data.DataTable CreateDataTableFromXml(string xmlFileName)
        {
            System.Data.DataTable Dt = new System.Data.DataTable();
            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(xmlFileName);
                Dt.Load(ds.CreateDataReader());
            }
            catch(Exception ex)
            {

            }
            return Dt;
        }

        private void ExportDataTableToExcel(System.Data.DataTable table, string xlFile)
        {

            Microsoft.Office.Interop.Excel.Application xlApp1 = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp1.Workbooks.Add(Type.Missing);

            xlWorkSheet = (Worksheet)xlWorkBook.ActiveSheet;
            xlWorkSheet.Name = table.TableName;

            for (int i=1; i<table.Columns.Count+1; i++)
            {
                xlWorkSheet.Cells[i, 1] = table.Columns[i - 1].ColumnName;
            }

            for (int j=0; j<table.Rows.Count; j++)
            {
                for (int k=0; k<table.Columns.Count; k++)
                {
                    xlWorkSheet.Cells[k + 1, j + 2] = table.Rows[j].ItemArray[k].ToString();
                }
            }

            MessageBox.Show("Saving " + xlFile);

            xlWorkBook.SaveAs(Path.ChangeExtension(Path.GetFullPath(xlFile), ".xls"), XlFileFormat.xlWorkbookNormal);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp1.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp1);
            
        }*/
    }
}