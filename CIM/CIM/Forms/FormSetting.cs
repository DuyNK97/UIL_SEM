using CIM.Class;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CIM.Forms
{
    public partial class FormSetting : Form
    {
        private bool isLoadingForm;

        public FormSetting()
        {
            InitializeComponent();
            isLoadingForm = true;

            Dictionary<string, string> currentData = Global.ReadValueFileTxt(Global.GetFilePathSetting(), new List<string> { "Is_Check_NAS", "DiskLocal", "DiskNetwork", "Auto_Delete_CSV", "Day_Delete_CSV" });

            Global.IsCheckNAS = int.Parse(currentData["Is_Check_NAS"]);
            Global.AutoDeleteCSV = int.Parse(currentData["Auto_Delete_CSV"]);

            pathLocal.Text = currentData["DiskLocal"];
            pathNAS.Text = currentData["DiskNetwork"];
            dayDeleteCSV.Text = currentData["Day_Delete_CSV"];

            btnSaveNAS.Checked = Global.IsCheckNAS == 1;
            cbAutoDeleteCSV.Checked = Global.AutoDeleteCSV == 1;

            dayDeleteCSV.Enabled = Global.AutoDeleteCSV == 1;
        }

        private void FormSetting_Load(object sender, EventArgs e)
        {
            isLoadingForm = false;
        }

        private async void btnSaveNAS_CheckedChanged(object sender, EventArgs e)
        {
            if (isLoadingForm) return;

            bool isChecked = btnSaveNAS.Checked;

            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
            {
                { "Is_Check_NAS", isChecked ? "1" : "0" }
            });

            Global.IsCheckNAS = isChecked ? 1 : 0;

            lbResultSetting.Text = string.Empty;
            await Task.Delay(50);
            lbResultSetting.Text = "Saved successfully!";
        }

        private void cbAutoDeleteCSV_CheckedChanged(object sender, EventArgs e)
        {
            if (isLoadingForm) return;

            bool isChecked = cbAutoDeleteCSV.Checked;

            dayDeleteCSV.Enabled = isChecked;

            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
            {
                { "Auto_Delete_CSV", isChecked ? "1" : "0" }
            });

            Global.AutoDeleteCSV = isChecked ? 1 : 0;
        }

        private async void btnSaveConfigPath_Click(object sender, EventArgs e)
        {
            string diskLocal = pathLocal.Text.Trim();
            string diskNetwork = pathNAS.Text.Trim();
            string dayDelete = dayDeleteCSV.Text.Trim();

            if (string.IsNullOrWhiteSpace(diskLocal) || string.IsNullOrWhiteSpace(diskNetwork))
            {
                MessageBox.Show("Path disk can not empty!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(dayDelete) && cbAutoDeleteCSV.Checked)
            {
                MessageBox.Show("Day delete CSV can not empty!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
            {
                { "DiskLocal", diskLocal },
                { "DiskNetwork", diskNetwork },
            });

            if (cbAutoDeleteCSV.Checked)
            {
                Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
                {
                    { "Day_Delete_CSV", dayDelete }
                });
            }

            Global.CSVD = diskLocal;
            Global.CSV = diskNetwork;

            lbResultSetting.Text = string.Empty;
            await Task.Delay(50);
            lbResultSetting.Text = "Saved successfully!";
        }
    }
}
