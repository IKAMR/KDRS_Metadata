using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KDRS_Metadata
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application xlApp;

        DataConverter converter = new DataConverter();
        JsonReader jsonReader = new JsonReader();

        List<string> priorities = new List<string> { };

        Hashtable myHashtable;

        List<string> resultList = new List<string>();

        string inputFileName;

        public Form1()
        {
            InitializeComponent();
            Text = Globals.toolName + " " + Globals.toolVersion;
            this.AllowDrop = true;
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.DragEnter += new DragEventHandler(Form1_DragEnter);

            //textBox1.AutoSize = true;

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

            label1.Text = "";
            textBox1.Clear();
            resultList.Clear();
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Count() > 1)
            {
                label1.Text = "Vennligst bare en fil av gangen... ;D";
            }
            else
            {
                fileName = files[0].ToString();
                Console.WriteLine(fileName);

                inputFileName = fileName;

                backgroundWorker1 = new BackgroundWorker();
                backgroundWorker1.DoWork += backgroundWorker1_DoWork;
                backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
                backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;
                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.RunWorkerAsync(fileName);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                textBox1.Text = "Error: " + e.Error.Message;
                KillExcel();
            }
            else
            {
                label1.Text = "Job complete!";
                textBox1.Text = "";
                foreach (string l in (List<string>)e.Result)
                {
                    textBox1.AppendText("\r\n" + l);
                }
                KillExcel();
            }
        }

        private void reader_OnProgressUpdate(int value, int total)
        {
            base.Invoke((System.Action)delegate
            {
                textBox1.Text = "Tables handled: " + value + " of " + total;
            });
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label1.Text = "Converting " + e.UserState.ToString();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            CheckPrioList();
            
            string fileName = e.Argument as string;

            string fileType = Path.GetExtension(fileName);
            Console.WriteLine("FileType: " + fileType);

            try
            {
                Console.WriteLine("Trying file: " + fileName + ", type: " + fileType);
                switch (fileType)
                {
                    case ".json":

                        backgroundWorker1.ReportProgress(0, fileName);
                        resultList.Add("Source: " + fileName);

                        jsonReader.OnProgressUpdate += reader_OnProgressUpdate;

                        jsonReader.ParseJson(fileName, priorities, includeTables.Checked);

                        resultList.Add("Target: " + jsonReader.excelFileName);
                        resultList.Add("Tables: " + jsonReader.tableCount);

                        foreach (Schema schema in jsonReader.schemaNames)
                        {
                            resultList.Add(schema.Folder + "   Name: " + schema.Name);
                        }

                        break;
                    case ".xml":

                        backgroundWorker1.ReportProgress(0, fileName);
                        resultList.Add("Source: " + fileName);

                        converter.OnProgressUpdate += reader_OnProgressUpdate;

                        converter.Convert(fileName, includeTables.Checked);

                        resultList.Add("Target: " + converter.excelFileName);
                        resultList.Add("Tables: " + converter.totalTableCount);

                        foreach (Schema schema in converter.schemaNames)
                        {
                            resultList.Add(schema.Folder + "   Name: " + schema.Name);
                        }

                        
                        break;
                }
                e.Result = resultList;
            }
            catch (COMException ex)
            {
                Console.WriteLine("ComExeption: " + ex);
                throw new Exception("Please close file: " + converter.excelFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exeption: " + ex);
                throw ex;
            }

        }

        //----------------------------------------------------------------------------------------------

        private void CheckPrioList()
        {
            priorities.Clear();

            //"HIGH", "MEDIUM", "LOW", "SYSTEM", "EMPTY", null
            if (priorityHigh.Checked)
            {
                priorities.Add("HIGH");
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

        private void btnCopyLog_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(textBox1.Text);
        }

        private void btnSaveLog_Click(object sender, EventArgs e)
        {
            string inputFolder = Path.GetDirectoryName(inputFileName);
            string filename  = Path.Combine(inputFolder, Path.GetFileNameWithoutExtension(inputFileName) + "_log_" + DateTime.Now.ToString("yyyy-MM-dd-HHmm") + ".txt");

            File.WriteAllText(filename, textBox1.Text);
        }
        //----------------------------------------------------------------------------------------------

    }
    public static class Globals
    {
        public static readonly String toolName = "KDRS Metadata";
        public static readonly String toolVersion = "0.9.2-rc2";

        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        public static int PriSort(string priority)
        {
            switch (priority)
            {
                case "HIGH":
                    return 1;
                case "MEDIUM":
                    return 2;
                case "LOW":
                    return 3;
                case "SYSTEM":
                    return 4;
                case "STATS":
                    return 5;
                case "EMPTY":
                    return 6;
                case "DUMMY":
                    return 7;
            }

            return 8;
        }
    }
}
