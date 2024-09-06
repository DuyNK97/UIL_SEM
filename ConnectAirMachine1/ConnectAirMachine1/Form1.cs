using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ConnectAirMachine1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void btnGetResultAirMachine_Click(object sender, EventArgs e)
        {
            TCPClientAirMachineHandle clientAirMachineHandler = new TCPClientAirMachineHandle();

            ResultAirMachine rsAirMachine = await clientAirMachineHandler.ConnectTCP(Int32.Parse(txtClient.Text));

            if (!string.IsNullOrWhiteSpace(rsAirMachine.result))
            {
                MessageBox.Show($"Result: {rsAirMachine.result}, sccm: {rsAirMachine.sccm}");
            }
        }
    }
}
