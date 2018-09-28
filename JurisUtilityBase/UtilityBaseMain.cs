using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            if (rbOne.Checked && String.IsNullOrEmpty(txtPrebill.Text))
                MessageBox.Show("Please type in the prebill number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                UpdateStatus("Processing Tax 1 Update.", 0, 1);
                string RunSQL = "";
                // Enter your SQL code here
                // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);
                if (rbAll.Checked == true)
                    RunSQL = "update prebillmatter " +
                         "set pbmtax1bld = ivutax " +
                         " from (select Clicode, Clireportingname, Matcode, Matreportingname,matsysnbr, pbfprebill, cast(sum(utamount) as decimal(12,2)) as IVUEntries, cast(sum(utamount * .04) as decimal(12,2)) as IVUTax" +
                            " from matter" +
                            " inner join client on matclinbr=clisysnbr" +
                            " inner join prebillfeeitem on pbfmatter=matsysnbr" +
                            " inner join unbilledtime on pbfutrecnbr=utrecnbr and pbfutbatch=utbatch" +
                            " where utcode1='IVU'  " +
                            "group by clicode, clireportingname, matcode, matreportingname, pbfprebill, matsysnbr) PBM " +
                            "where pbmprebill=pbfprebill and pbmmatter=matsysnbr ";
                else
                    RunSQL = "update prebillmatter " +
                         "set pbmtax1bld = ivutax " +
                         " from (select Clicode, Clireportingname, Matcode, Matreportingname,matsysnbr, pbfprebill, cast(sum(utamount) as decimal(12,2)) as IVUEntries, cast(sum(utamount * .04) as decimal(12,2)) as IVUTax" +
                            " from matter" +
                            " inner join client on matclinbr=clisysnbr" +
                            " inner join prebillfeeitem on pbfmatter=matsysnbr" +
                            " inner join unbilledtime on pbfutrecnbr=utrecnbr and pbfutbatch=utbatch" +
                              " where utcode1='IVU' and pbfprebill=" + txtPrebill.Text.ToString() +
                               "group by clicode, clireportingname, matcode, matreportingname, pbfprebill, matsysnbr) PBM " +
                            "where pbmprebill=pbfprebill and pbmmatter=matsysnbr ";

                _jurisUtility.ExecuteNonQueryCommand(0, RunSQL);
                UpdateStatus("Tax 1 Updated.", 1, 1);
            }
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (!rbAll.Checked && !rbOne.Checked)
                MessageBox.Show("Please select at least one checkbox", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
           // if (string.IsNullOrEmpty(toAtty) || string.IsNullOrEmpty(fromAtty))
          //      MessageBox.Show("Please select from both Timekeeper drop downs", "Selection Error");
          //  else
          //  {
                //generates output of the report for before and after the change will be made to client
            if (!rbAll.Checked && !rbOne.Checked)
                MessageBox.Show("Please select at least one checkbox", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (rbOne.Checked && String.IsNullOrEmpty(txtPrebill.Text))
                MessageBox.Show("Please type in the prebill number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string SQLTkpr = getReportSQL();

                DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                ReportDisplay rpds = new ReportDisplay(myRSTkpr);
                rpds.Show();

            }
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (rbAll.Checked == true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,pbfprebill, cast(sum(utamount) as decimal(12,2)) as IVUEntries, cast(sum(utamount * .04) as decimal(12,2)) as IVUTax" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join prebillfeeitem on pbfmatter=matsysnbr" +
                        " inner join unbilledtime on pbfutrecnbr=utrecnbr and pbfutbatch=utbatch" +       
                        " where utcode1='IVU'  " +
                        " group by clicode, clireportingname, matcode, matreportingname, pbfprebill";
         else
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,pbfprebill, cast(sum(utamount) as decimal(12,2)) as IVUEntries, cast(sum(utamount * .04) as decimal(12,2)) as IVUTax" +
                          " from matter" +
                          " inner join client on matclinbr=clisysnbr" +
                          " inner join prebillfeeitem on pbfmatter=matsysnbr" +
                          " inner join unbilledtime on pbfutrecnbr=utrecnbr and pbfutbatch=utbatch" +
                          " where utcode1='IVU' and pbfprebill=" + txtPrebill.Text.ToString() + 
                          " group by clicode, clireportingname, matcode, matreportingname, pbfprebill";


            return reportSQL;
        }

        private void labelDescription_Click(object sender, EventArgs e)
        {

        }

        private void rbOne_CheckedChanged(object sender, EventArgs e)
        {
            if(rbOne.Checked == true)
            { txtPrebill.Visible = true; }
            else
            { txtPrebill.Visible = false; }
    
        }
    }
}
