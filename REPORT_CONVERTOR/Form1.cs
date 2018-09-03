using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace REPORT_CONVERTOR
{
    public partial class Form1 : Form
    {
        string[] inputfiles;
        int filecount = 0;
        string message;
        string inputDirectory = "";
        string outputpath = "";
        string outputOption = "PortableDocFormat";
        ReportDocument cryRpt;
        int PARALLELISM = 4;
        public Form1()
        {
            InitializeComponent();
            //cryRpt = new ReportDocument();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Text File";
            theDialog.Multiselect = true;
            theDialog.Filter = "Crystal Report File|*.rpt";
            theDialog.InitialDirectory = @"C:\";
            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                inputfiles = theDialog.FileNames;
                textBox1.Text = inputfiles[0];
                inputDirectory = Path.GetDirectoryName(inputfiles[0]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = folderDialog.SelectedPath;
                    outputpath = folderDialog.SelectedPath;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (inputDirectory.Length == 0 || outputpath.Length == 0 || inputDirectory.Equals(outputpath))
            {
                MessageBox.Show("You must specify input and output path and also they cannot be same", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            button2.Enabled = false;
            button2.Text = "Processing ...";
            int count = 0;
            listBox1.Items.Clear();
            filecount = inputfiles.Length;
            label6.Text = "0 of " + filecount;
            if (!checkBox1.Checked)
            {
                foreach (string file in inputfiles)
                {
                    message = ProcessReports(outputpath, Path.GetFileNameWithoutExtension(file), file);
                    listBox1.Items.Add(message);
                    count++;
                    label6.Text = count + " of " + filecount;
                }
                label6.Text = "Completed";
                button2.Text = "Convert";
                button2.Enabled = true;
            }
            else
            {
                count = 0;
                filecount = Directory.GetFiles(inputDirectory, "*.rpt", SearchOption.TopDirectoryOnly).Length;
                label6.Text = "0 of " + filecount;
                //Parallel Foreach Implementation
                var yourForeachTask = Task.Run(() =>
                {
                    Parallel.ForEach(Directory.GetFiles(inputDirectory, "*.rpt", SearchOption.TopDirectoryOnly), new ParallelOptions { MaxDegreeOfParallelism = PARALLELISM }, (currentFile) =>
                    {
                        message = ProcessReports(outputpath, Path.GetFileNameWithoutExtension(currentFile), currentFile);
                        //listBox1.Items.Add(message);    
                        count++;
                        BeginInvoke((Action)delegate
                        {
                            label6.Text = count + " of " + filecount;
                            listBox1.Items.Add(message);
                        });
                        Console.WriteLine("Processing {0} on thread {1}", currentFile, Thread.CurrentThread.ManagedThreadId);
                    });
                });
                //label6.Text = "Completed";
                //button2.Text = "Convert";
                //button2.Enabled = true;
            }
        }
        public string ProcessReports(string outputpath, string filenameNoExtension, string file)
        {
                ReportDocument cryRpt = new ReportDocument();
                cryRpt.Load(file);
                switch (outputOption)
                {
                    case "PortableDocFormat":
                        cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, outputpath + "\\" + filenameNoExtension + ".pdf");
                        break;
                    case "CrystalReport":
                        cryRpt.ExportToDisk(ExportFormatType.CrystalReport, outputpath + "\\" + filenameNoExtension + ".rpt");
                        break;
                    case "CharacterSeparatedValues":
                        cryRpt.ExportToDisk(ExportFormatType.CharacterSeparatedValues, outputpath + "\\" + filenameNoExtension + ".csv");
                        break;
                    case "Excel":
                        cryRpt.ExportToDisk(ExportFormatType.Excel, outputpath + "\\" + filenameNoExtension + ".xls");
                        break;
                    case "ExcelRecord":
                        cryRpt.ExportToDisk(ExportFormatType.ExcelRecord, outputpath + "\\" + filenameNoExtension + ".xls");
                        break;
                    case "ExcelWorkbook":
                        cryRpt.ExportToDisk(ExportFormatType.ExcelWorkbook, outputpath + "\\" + filenameNoExtension + ".xls");
                        break;
                    case "TabSeperatedText":
                        cryRpt.ExportToDisk(ExportFormatType.TabSeperatedText, outputpath + "\\" + filenameNoExtension + ".txt");
                        break;
                    case "RichText":
                        cryRpt.ExportToDisk(ExportFormatType.RichText, outputpath + "\\" + filenameNoExtension + ".rtf");
                        break;
                    case "Text":
                        cryRpt.ExportToDisk(ExportFormatType.Text, outputpath + "\\" + filenameNoExtension + ".txt");
                        break;
                    case "Xml":
                        cryRpt.ExportToDisk(ExportFormatType.Xml, outputpath + "\\" + filenameNoExtension + ".xml");
                        break;
                    case "EditableRTF":
                        cryRpt.ExportToDisk(ExportFormatType.EditableRTF, outputpath + "\\" + filenameNoExtension + ".rtf");
                        break;
                    case "WordForWindows":
                        cryRpt.ExportToDisk(ExportFormatType.WordForWindows, outputpath + "\\" + filenameNoExtension + ".doc");
                        break;
                    case "HTML32":
                        cryRpt.ExportToDisk(ExportFormatType.HTML32, outputpath + "\\" + filenameNoExtension + ".html");
                        break;
                    case "HTML40":
                        cryRpt.ExportToDisk(ExportFormatType.HTML40, outputpath + "\\" + filenameNoExtension + ".html");
                        break;
                    case "RPTR":
                        cryRpt.ExportToDisk(ExportFormatType.RPTR, outputpath + "\\" + filenameNoExtension + ".rptr");
                        break;
                    default:
                        cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, outputpath + "\\" + filenameNoExtension + ".pdf");
                        break;
                }
                cryRpt.Close();
                cryRpt.Dispose();
                return file;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            outputOption = comboBox1.Text;
        }
        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("\nLog Entry : ");
            w.Write(" {0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
            w.Write("  :{0}", logMessage);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (AboutBox1 box = new AboutBox1())
            {
                box.ShowDialog(this);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
