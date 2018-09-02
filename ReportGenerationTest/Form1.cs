using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Windows.Forms;

namespace ReportGenerationTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConnectionInfo crconnectioninfo = new ConnectionInfo();
            ReportDocument cryrpt = new ReportDocument();
            TableLogOnInfos crtablelogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtablelogoninfo = new TableLogOnInfo();

            Tables CrTables;

            crconnectioninfo.ServerName = "AIBSIT";
            crconnectioninfo.DatabaseName = "AIBSIT";
            crconnectioninfo.UserID = "db2inst4";
            crconnectioninfo.Password = "db2inst4";

            string startupPath = System.IO.Directory.GetCurrentDirectory();
            ReportDocument cryRpt = new ReportDocument();
            cryRpt.Load(startupPath + "\\ACCOUNT_STATEMENT_PRODUCTION.rpt");

            ParameterFields paremeters = cryRpt.ParameterFields;

            foreach (ParameterField parameter in paremeters)
            {
                MessageBox.Show(parameter.Name);
            }

            CrTables = cryRpt.Database.Tables;

            foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in CrTables)
            {
                crtablelogoninfo = CrTable.LogOnInfo;
                crtablelogoninfo.ConnectionInfo = crconnectioninfo;
                CrTable.ApplyLogOnInfo(crtablelogoninfo);
            }

            cryRpt.SetParameterValue("startDate", "2016-01-01");
            cryRpt.SetParameterValue("endDate", Convert.ToDateTime("2018-04-04"));
            cryRpt.SetParameterValue("accountNumber", "01320360047701");

            crystalReportViewer1.ReportSource = cryRpt;
            crystalReportViewer1.Refresh();

            //cryRpt.SetDatabaseLogon("db2inst4", "db2inst4", "192.168.12.116", "AIBSIT");
            //

            //cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "D:\\myfile.pdf");
        }
    }
}
