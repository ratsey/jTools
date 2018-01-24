using Microsoft.Win32;
using System;
using System.Drawing;
using System.Windows.Forms;
using Utils.Dlg;

namespace JTools
{
    // Uses Thread Apartment Safe Open/Save File Dialogs for C#
    // https://www.codeproject.com/Articles/841702/Thread-Apartment-Safe-Open-Save-File-Dialogs-for-C
    // Christian Kleinheinz, 24 Nov 2014 

    public partial class frmSetup : Form
    {
        public frmSetup()
        {
            InitializeComponent();

            ReadSettings();
        }

        private void btnFindMLO_Click(object sender, EventArgs e)
        {
            CFileOpenDlgThreadApartmentSafe dlg = new CFileOpenDlgThreadApartmentSafe();
            dlg.DefaultExt = "EXE";
            dlg.Filter = "Executables (*.exe)|*.exe";

            Point ptStartLocation = new Point(this.Location.X, this.Location.Y);

            dlg.StartupLocation = ptStartLocation;
            DialogResult res = dlg.ShowDialog();

            if (res == DialogResult.OK)
            {
                txtMLOExe.Text = dlg.FilePath;
            }
        }

        private void btnMLODataFile_Click(object sender, EventArgs e)
        {
            CFileOpenDlgThreadApartmentSafe dlg = new CFileOpenDlgThreadApartmentSafe();
            dlg.DefaultExt = "ML";
            dlg.Filter = "MyLifeOrganized Data File (*.ml)|*.ML";
            dlg.StartupLocation = new Point(this.Location.X, this.Location.Y);

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtMLODataFile.Text = dlg.FilePath;
            }
        }

        private void ReadSettings()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Ratsey\OneNoteMLO");

            string mlexe = rk.GetValue("MLOExecutable", "").ToString();
            string dataFile = rk.GetValue("DataFile", "").ToString();
            string rootGUID = rk.GetValue("RootGUID", "").ToString();
            string isDebug = rk.GetValue("IsDebug", "False").ToString();
            string logFile = rk.GetValue("LogFile", "").ToString();

            txtMLOExe.Text = mlexe;
            txtMLODataFile.Text = dataFile;
            txtMLOGUID.Text = rootGUID;
            txtLogFile.Text = logFile;
            chkDebug.Checked = isDebug == "True";
        }

        private void WriteSettings(string EXE, string DataFile, string RootGUID, string IsDebug, string LogFile)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Ratsey\OneNoteMLO", true);

            rk.SetValue("MLOExecutable", EXE);
            rk.SetValue("DataFile", DataFile);
            rk.SetValue("RootGUID", RootGUID);
            rk.SetValue("IsDebug", IsDebug);
            rk.SetValue("LogFile", LogFile);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            WriteSettings(txtMLOExe.Text, txtMLODataFile.Text, txtMLOGUID.Text, chkDebug.Checked ? "True" : "False", txtLogFile.Text);
            this.Close();
        }

        private void btnLogFile_Click(object sender, EventArgs e)
        {
            CFileOpenDlgThreadApartmentSafe dlg = new CFileOpenDlgThreadApartmentSafe();
            dlg.DefaultExt = "TXT";
            dlg.Filter = "JTools Log File (*.txt)|*.txt";
            dlg.StartupLocation = new Point(this.Location.X, this.Location.Y);

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtLogFile.Text = dlg.FilePath;
            }
        }
    }
}
