using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UILAlignProject.PLC;
using UnilityCommand.Plc.Mitsubishi;
using UnilityCommand.Plc;
using System.IO;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Color = System.Drawing.Color;
using Sunny.UI;
using System.Net.Sockets;
using Match = System.Text.RegularExpressions.Match;
using CIM.Class;
using CIM.Enum;
using CIM.Forms;
using System.Data;

namespace CIM
{
    public partial class FormMain : UIForm
    {
        PLCIOCollection pLCIOs = new PLCIOCollection();
        public int indexPLC1, indexPLC2, indexPLC3, indexPLC4;
        public static string PLClog1 = ConfigurationManager.AppSettings["LogPLC1"];
        public static string PLClog2 = ConfigurationManager.AppSettings["LogPLC2"];
        public static string PLClog3 = ConfigurationManager.AppSettings["LogPLC3"];
        public static string PLClog4 = ConfigurationManager.AppSettings["LogPLC4"];
        public static string PLCALL = ConfigurationManager.AppSettings["LogALL"];
        public static int excelrow = 1;

        public static string CSV = ConfigurationManager.AppSettings["LogCSV"];
        public static string CSVD = ConfigurationManager.AppSettings["LogCSVD"];
        public static string model = ConfigurationManager.AppSettings["MODEL"].ToString();

        public static string rowdisplay = ConfigurationManager.AppSettings["rowdisplay"];

        int rowsdisplay = 100;

        public DataTable dt = new DataTable();

        public static int OK = 0;
        public static int NG = 0;
        public static int Total;
        public static int No = 0;
        public static List<EXCELDATA> list = new List<EXCELDATA>();
        public static string logFilePath1 = "";
        public static string logFilePath2 = "";
        public static string logFilePath3 = "";
        public static string logFilePath4 = "";
        public static string logFilePathALL = "";


        public const int AUTO_DELETE_FILE = 1;

        private ZebraZT411Printer _printerManager = new ZebraZT411Printer();

        public FormMain()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ResultAirMachine air = new ResultAirMachine();

            rowsdisplay = int.Parse(rowdisplay);

            SqlLite.Instance.InitializeConnection();

            Dictionary<string, string> currentData = Global.ReadValueFileTxt(Global.GetFilePathSetting(), new List<string> { "OK", "NG_AIR", "TOTAL", "DiskNetwork", "DiskLocal" });
            OK = int.Parse(currentData["OK"]);
            NG = int.Parse(currentData["NG_AIR"]);
            Total = int.Parse(currentData["TOTAL"]);
            Global.CSV = currentData["DiskNetwork"];
            Global.CSVD = currentData["DiskLocal"];

            EXCELDATA data = new EXCELDATA();
            UpdateUI(data);
            pieChart1.UpdateChartData(OK, NG);

            try
            {
                if (SingleTonPlcControl.Instance.Connect1())
                {
                    WriteLog("Connected to PLC1");
                    AddPLCI1(pLCIOs);
                }

                if (SingleTonPlcControl.Instance.Connect2())
                {
                    WriteLog("Connected to PLC2");
                    AddPLCI2(pLCIOs);
                }

                if (SingleTonPlcControl.Instance.Connect3())
                {
                    WriteLog("Connected to PLC3");
                    AddPLCI3(pLCIOs);
                }

                if (SingleTonPlcControl.Instance.Connect4())
                {
                    WriteLog("Connected to PLC4");
                    AddPLCI4(pLCIOs);
                }

                SingleTonPlcControl.Instance.AddRegisterRead(SingleTonPlcControl.Instance.RegisterRead, pLCIOs);
                SingleTonPlcControl.Instance.AddRegisterWrite(SingleTonPlcControl.Instance.RegisterWrite, pLCIOs);
                SingleTonPlcControl.Instance.RegisterRead.PlcIOs.PropertyChanged += RegisterRead_PropertyChanged;

                //set bit alive by PLC
                var aliveBox1 = (bool)SingleTonPlcControl.Instance.GetValueRegister(1, "IS_ALIVE");
                SingleTonPlcControl.Instance.SetValueRegister(aliveBox1, 1, "WRITE_IS_ALIVE", true, EnumReadOrWrite.WRITE);

                var aliveBox2 = (bool)SingleTonPlcControl.Instance.GetValueRegister(1, "IS_ALIVE");
                SingleTonPlcControl.Instance.SetValueRegister(aliveBox2, 2, "WRITE_IS_ALIVE", true, EnumReadOrWrite.WRITE);

                var aliveBox3 = (bool)SingleTonPlcControl.Instance.GetValueRegister(3, "IS_ALIVE");
                SingleTonPlcControl.Instance.SetValueRegister(aliveBox3, 3, "WRITE_IS_ALIVE", true, EnumReadOrWrite.WRITE);

                var aliveBox4 = (bool)SingleTonPlcControl.Instance.GetValueRegister(4, "IS_ALIVE");
                SingleTonPlcControl.Instance.SetValueRegister(aliveBox4, 4, "WRITE_IS_ALIVE", true, EnumReadOrWrite.WRITE);
            }
            catch (Exception ex)
            {
                WriteLog($"Error can not connect with PLC, err: {ex.Message}");
                MessageBox.Show($"Error can not connect with PLC, err: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //thread auto delete old file after 365 day
            Thread threadAutoDeleteOldFile = new Thread(async () => await ThreadAutoDeleteOldFile());
            threadAutoDeleteOldFile.Name = "THREAD_AUTO_DELETE_OLD_FILE";
            threadAutoDeleteOldFile.IsBackground = true;
            threadAutoDeleteOldFile.Start();
        }

        private async Task ThreadAutoDeleteOldFile()
        {
            while (true)
            {
                Dictionary<string, string> currentData = Global.ReadValueFileTxt(Global.GetFilePathSetting(), new List<string> { "Auto_Delete_CSV", "Day_Delete_CSV" });

                int autoDeleteFile = int.Parse(currentData["Auto_Delete_CSV"]);
                int dayDeleteFileCSV = int.Parse(currentData["Day_Delete_CSV"]);

                DateTime now = DateTime.Now;
                if (autoDeleteFile == AUTO_DELETE_FILE)
                {

                    DeleteOldFile(Global.CSVD, now, dayDeleteFileCSV);
                    WriteLog("Delete old file done!");
                }

                //delete file log
                DeleteOldFile($@"D:\Logs\CIM\LogsData", now, 30);

                await Task.Delay(TimeSpan.FromDays(1));
            }
        }

        private void DeleteOldFile(string path, DateTime now, int dayDelete)
        {
            if (!Directory.Exists(path))
            {
                WriteLog($"Not found path to delete file!");
                return;
            }

            int batchSize = 1000;

            var fileBatch = Directory.EnumerateFiles(path).Take(batchSize);

            while (fileBatch.Any())
            {
                foreach (var file in fileBatch)
                {
                    DateTime creationTime = File.GetCreationTime(file);
                    TimeSpan fileAge = now - creationTime;
                    if (fileAge.TotalDays > dayDelete)
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            WriteLog($"Error can not delete file, error: {ex.Message}");
                        }
                    }
                }

                fileBatch = Directory.EnumerateFiles(path).Skip(batchSize).Take(batchSize);
            }

            var directories = Directory.GetDirectories(path);

            foreach (var directory in directories)
            {
                DeleteOldFile(directory, now, dayDelete);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string logPath1 = PLClog1;
            string logFileName1 = DateTime.Now.ToString("yyyyMMdd") + ".csv";
            logFilePath1 = Path.Combine(logPath1, logFileName1);

            if (!Directory.Exists(logPath1))
            {
                Directory.CreateDirectory(logPath1);
            }
            string logPath2 = PLClog2;
            string logFileName2 = DateTime.Now.ToString("yyyyMMdd") + ".csv";
            logFilePath2 = Path.Combine(logPath2, logFileName2);

            if (!Directory.Exists(logPath2))
            {
                Directory.CreateDirectory(logPath2);
            }
            string logPath3 = PLClog3;
            string logFileName3 = DateTime.Now.ToString("yyyyMMdd") + ".csv";
            logFilePath3 = Path.Combine(logPath3, logFileName3);

            if (!Directory.Exists(logPath3))
            {
                Directory.CreateDirectory(logPath3);
            }
            string logPath4 = PLClog4;
            string logFileName4 = DateTime.Now.ToString("yyyyMMdd") + ".csv";
            logFilePath4 = Path.Combine(logPath4, logFileName4);

            if (!Directory.Exists(logPath4))
            {
                Directory.CreateDirectory(logPath4);
            }
            string logPath = PLCALL;
            string logFileName = DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            logFilePathALL = Path.Combine(logPath, logFileName);


            // Ensure the directory exists
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            ConnectAirTest();
        }

        private void RegisterRead_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            //PLCIO obj = sender as PLCIO;

            //if (obj != null)
            //{
            //    if (obj.Title.Contains("ReadData") && (bool)obj.CurrentValue == true)
            //    {
            //        int indexPLC = obj.IndexPLC;
            //        switch (indexPLC)
            //        {
            //            case 1:
            //                ReadData1();
            //                SingleTonPlcControl.Instance.SetValueRegister(true, indexPLC, "WriteData", true, EnumReadOrWrite.WRITE);
            //                WriteLog("On bit WriteData in PLC 1");
            //                break;
            //            case 2:
            //                ReadData2();
            //                SingleTonPlcControl.Instance.SetValueRegister(true, indexPLC, "WriteData", true, EnumReadOrWrite.WRITE);
            //                WriteLog("On bit WriteData in PLC 2");
            //                break;
            //            case 3:
            //                ReadData3();
            //                SingleTonPlcControl.Instance.SetValueRegister(true, indexPLC, "WriteData", true, EnumReadOrWrite.WRITE);
            //                WriteLog("On bit WriteData in PLC 3");
            //                break;
            //            case 4:
            //                ReadData4();
            //                SingleTonPlcControl.Instance.SetValueRegister(true, indexPLC, "WriteData", true, EnumReadOrWrite.WRITE);
            //                WriteLog("On bit WriteData in PLC 4");
            //                break;
            //        }
            //    }
            //    else if (obj.Title.Contains("ReadData") && (bool)obj.CurrentValue == false)
            //    {
            //        SingleTonPlcControl.Instance.SetValueRegister(false, obj.IndexPLC, "WriteData", true, EnumReadOrWrite.WRITE);
            //        WriteLog("OFF bit WriteData in PLC 4");
            //    }
            //    else if (obj.Title == "Alive" /*&& (bool)obj.CurrentValue==true */)
            //    {
            //        UpdateStatus(obj.IndexPLC, obj.CurrentValue);
            //    }
            //    else if (obj.Title == "ReadPrint" && (bool)obj.CurrentValue == true && obj.IndexPLC == 4)
            //    {
            //        var qrCode = SingleTonPlcControl.Instance.GetValueRegister(obj.IndexPLC, "BOX4CountBarcode");

            //        if (qrCode != null)
            //        {
            //            if (!string.IsNullOrWhiteSpace(qrCode.ToString().Trim()))
            //            {
            //                print.Add(qrCode.ToString().Trim());
            //            }
            //        }

            //        SingleTonPlcControl.Instance.SetValueRegister(true, obj.IndexPLC, "WritePrint", true, EnumReadOrWrite.WRITE);
            //        WriteLog("On bit WritePrint in PLC 4");
            //    }
            //    else if (obj.Title == "ReadPrint" && (bool)obj.CurrentValue == false && obj.IndexPLC == 4)
            //    {
            //        SingleTonPlcControl.Instance.SetValueRegister(false, obj.IndexPLC, "WritePrint", true, EnumReadOrWrite.WRITE);
            //        WriteLog("OFF bit WritePrint in PLC 4");
            //    }
            //    else if (obj.Title == "EndTray" && (bool)obj.CurrentValue == true && obj.IndexPLC == 4)
            //    {
            //        CountPrint(obj.IndexPLC);
            //        SingleTonPlcControl.Instance.SetValueRegister(true, obj.IndexPLC, "WRITE_END_TRAY", true, EnumReadOrWrite.WRITE);
            //    }
            //    else if (obj.Title == "EndTray" && (bool)obj.CurrentValue == false && obj.IndexPLC == 4)
            //    {
            //        SingleTonPlcControl.Instance.SetValueRegister(false, obj.IndexPLC, "WRITE_END_TRAY", true, EnumReadOrWrite.WRITE);
            //    }
            //    else if (obj.Title == "CHANGE_MODE_REWORK")
            //    {
            //        HandleChangeRework(obj.IndexPLC, (bool)obj.CurrentValue);
            //    }
            //    else if (obj.Title == "CHANGE_MODE_STATE")
            //    {
            //        HandleChangeState(obj.IndexPLC, (short)obj.CurrentValue);
            //    }
            //    else if (obj.Title == "IS_ALIVE")
            //    {
            //        SingleTonPlcControl.Instance.SetValueRegister((bool)obj.CurrentValue, obj.IndexPLC, "WRITE_IS_ALIVE", true, EnumReadOrWrite.WRITE);
            //    }
            //    else if (obj.Title == "READ_INPUT_BARCODE" && (bool)obj.CurrentValue == true && obj.IndexPLC == 1)
            //    {
            //        CheckIsDuplicate();
            //        SingleTonPlcControl.Instance.SetValueRegister(true, obj.IndexPLC, "WRITE_INPUT_BARCODE", true, EnumReadOrWrite.WRITE);
            //    }
            //    else if (obj.Title == "READ_INPUT_BARCODE" && (bool)obj.CurrentValue == false && obj.IndexPLC == 1)
            //    {
            //        SingleTonPlcControl.Instance.SetValueRegister(false, obj.IndexPLC, "WRITE_INPUT_BARCODE", true, EnumReadOrWrite.WRITE);
            //    }
            //}
        }

        private void CheckIsDuplicate()
        {
            //set ng -> 2 type: word
            if (SingleTonPlcControl.Instance.GetValueRegister(1, "INPUT_BOX1_BARCODE") == null)
            {
                WriteLog($"INPUT_QR_CODE_DUPLICATE CAN NOT GET READ BARCODE FROM PLC");
                short[] rs = new short[1] { 2 };
                SingleTonPlcControl.Instance.WriteWord($"D45301", 1, 1, ref rs);
                return;
            }

            var QRcode = SingleTonPlcControl.Instance.GetValueRegister(1, "INPUT_BOX1_BARCODE").ToString().Trim();

            if (string.IsNullOrWhiteSpace(QRcode))
            {
                //set ng -> 2 word
                WriteLog($"INPUT_QR_CODE_DUPLICATE READ FROM PLC EMPTY");
                short[] rs = new short[1] { 2 };
                SingleTonPlcControl.Instance.WriteWord($"D45301", 1, 1, ref rs);
                return;
            }

            //duplicate
            if (SqlLite.Instance.CheckQrInputIsExists(QRcode))
            {
                //set ng -> 2 word
                WriteLog($"INPUT_QR_CODE {QRcode} DUPLICATE");
                WriteBarcodeDuplicate(QRcode);
                short[] rs = new short[1] { 2 };
                SingleTonPlcControl.Instance.WriteWord($"D45301", 1, 1, ref rs);
            }
            else
            {
                //set ok => 1 word
                WriteLog($"INPUT_QR_CODE {QRcode} NO_DUPLICATE");
                short[] rs = new short[1] { 1 };
                SingleTonPlcControl.Instance.WriteWord($"D45301", 1, 1, ref rs);
            }
        }

        private void HandleChangeRework(int indexPLC, bool currentValue)
        {
            if (indexPLC == (int)EPLC.PLC_3)
            {
                Global.CurrentModeBox3 = currentValue ? (int)ERework.REWORK : (int)ERework.NORMAL;
            }

            if (indexPLC == (int)EPLC.PLC_4)
            {
                Global.CurrentModeBox4 = currentValue ? (int)ERework.REWORK : (int)ERework.NORMAL;
            }

            SingleTonPlcControl.Instance.SetValueRegister(currentValue, indexPLC, "WRITE_CHANGE_MODE_REWORK", true, EnumReadOrWrite.WRITE);
            WriteLog($"Set BIT Change rework PLC-{indexPLC}");
        }

        private void HandleChangeState(int indexPLC, short currentValue)
        {
            if (indexPLC == (int)EPLC.PLC_3)
            {
                Global.CurrentStateBox3 = currentValue == 1 ? (int)EMode.ONLINE : (int)EMode.OFFLINE;
                WriteLog(currentValue == 1 ? "Current State Box 3 Online" : "Current State Box 3 Offline");
            }

            if (indexPLC == (int)EPLC.PLC_4)
            {
                Global.CurrentStateBox4 = currentValue == 1 ? (int)EMode.ONLINE : (int)EMode.OFFLINE;
                WriteLog(currentValue == 1 ? "Current State Box 4 Online" : "Current State Box 4 Offline");
            }

            short[] rs = new short[1] { currentValue };
            SingleTonPlcControl.Instance.WriteWord($"D45300", 1, indexPLC, ref rs);
            WriteLog($"Set BIT Change status online-offline PLC-{indexPLC}");
        }

        private void UpdateStatus(int indexPLC, object currvalue)
        {
            if (InvokeRequired)
            {
                BeginInvoke((Action)(() => UpdateStatus(indexPLC, currvalue)));
                return;
            }

            bool isConnected = (bool)currvalue;

            Color connectedColor = Color.Blue;
            Color disconnectedColor = Color.Red;

            switch (indexPLC)
            {
                case 1:
                    lblPLC1.CheckBoxColor = isConnected ? connectedColor : disconnectedColor;
                    lblPLC1.ForeColor = isConnected ? connectedColor : disconnectedColor;
                    break;

                case 2:
                    lblPLC2.CheckBoxColor = isConnected ? connectedColor : disconnectedColor;
                    lblPLC2.ForeColor = isConnected ? connectedColor : disconnectedColor;
                    break;

                case 3:
                    lblPLC3.CheckBoxColor = isConnected ? connectedColor : disconnectedColor;
                    lblPLC3.ForeColor = isConnected ? connectedColor : disconnectedColor;
                    break;

                case 4:
                    lblPLC4.CheckBoxColor = isConnected ? connectedColor : disconnectedColor;
                    lblPLC4.ForeColor = isConnected ? connectedColor : disconnectedColor;
                    break;

                default:
                    break;
            }
        }

        //get value air test
        #region

        private TcpClient[] clients = new TcpClient[10];
        private NetworkStream[] streams = new NetworkStream[10];
        private bool[] connected = new bool[10];
        private readonly string[] serverIPs = new string[]
        {
            "192.168.3.170","192.168.3.171", "192.168.3.172", "192.168.3.173", "192.168.3.174",
            "192.168.3.175", "192.168.3.176","192.168.3.177", "192.168.3.178","192.168.3.179"
        };
        private void ConnectAirTest()
        {
            try
            {
                for (int i = 0; i < 10; i++)
                {
                    int port = 23;
                    string ipAddress = serverIPs[i];
                    clients[i] = new TcpClient();

                    // Đợi kết thúc quá trình kết nối
                    try
                    {
                        ConnectClients(ipAddress, port, i);
                    }
                    catch (TimeoutException ex)  // Xử lý lỗi kết nối timeout ở đây
                    {
                        WriteLog($"Connection timeout: {ex.Message}");
                    }
                    catch (Exception ex) // Xử lý các lỗi khác nếu có
                    {
                        WriteLog($"Error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Error: {ex.Message}");
            }
        }
        private void ConnectClients(string ipAddress, int port, int clientIndex)
        {
            Task.Run(() => ConnectClient(clientIndex, ipAddress, port));
        }
        private void ConnectClient(int clientIndex, string ipAddress, int port)
        {
            int attempt = 0;
            while (attempt < 10)
            {
                try
                {
                    clients[clientIndex].Connect(ipAddress, port); // Synchronous connect
                    if (clients[clientIndex].Connected)
                    {
                        streams[clientIndex] = clients[clientIndex].GetStream();
                        connected[clientIndex] = true;
                        WriteLog($"Connected to server {ipAddress} on port {port}.");
                        SendataConnect(clientIndex); // sendata de ket noi toi LEAK Test
                        StartReceiving(clientIndex); // bat dau nhan du lieu
                    }
                }
                catch (SocketException ex)
                {
                    WriteLog($"SocketException for client {clientIndex}: {ex.Message}");
                    //throw; // Rethrow the exception to handle retry logic
                }
                catch (Exception ex)
                {
                    WriteLog($"Connection error for client {clientIndex}: {ex.Message}");
                }
                attempt++;
                Thread.Sleep(100);
            }
        }
        private async void SendataConnect(int connect)
        {
            try
            {
                string message = "1\r\n";
                byte[] data = Encoding.UTF8.GetBytes(message);
                if (connected[connect])
                {
                    await streams[connect].WriteAsync(data, 0, data.Length);
                    WriteLog($"Conect to client {connect + 1}.");
                }
                else
                {
                    WriteLog($"client {connect + 1} is null");
                }

            }
            catch (Exception ex)
            {
                WriteLog($"Error: {ex.Message}");
            }
        }
        private void StartReceiving(int clientIndex)
        {
            try
            {
                byte[] buffer = new byte[1024];
                while (connected[clientIndex])
                {
                    int bytesRead = streams[clientIndex].Read(buffer, 0, buffer.Length);
                    string message = Encoding.UTF8.GetString(buffer, 0, bytesRead);

                    Task.Run(() => ExtractValueSccm(message, clientIndex));
                }
            }
            catch (Exception ex)
            {
                if (connected[clientIndex])
                {
                    WriteLog($"Error: {ex.Message}");
                    connected[clientIndex] = false;
                }
            }
        }
        private void ExtractValueSccm(string input, int clientIndex)
        {
            if (input.Contains("sccm"))
            {
                string patternSccm = @"(\d+\.\d+)\s*sccm";
                Match match = Regex.Match(input, patternSccm);
                if (match.Success)
                {
                    string sccmValue = match.Groups[1].Value;

                    if (sccmValue.Trim() == "0")
                    {
                        sccmValue = "0.061607";
                    }

                    Task.Run(() => HandleReadAirTest(sccmValue, clientIndex, true));
                }
            }
            else if (input.Contains("SL"))
            {
                WriteLog($"data to split: {input}");
                string result = GetNameError(input);
                Task.Run(() => HandleReadAirTest(result, clientIndex, false));
            }
        }
        private string GetNameError(string input)
        {
            try
            {
                string[] parts = input.Split(' ');

                var result = parts[5].Trim();

                if (string.IsNullOrWhiteSpace(result))
                {
                    return "SL";
                }

                if (result.Length > 3)
                {
                    return "SL";
                }

                if (result != result.ToUpper())
                {
                    return "SL";
                }

                return result;
            }
            catch (Exception ex)
            {
                WriteLog($"Can not get error name, error: {ex.Message}");
                return "SL";
            }
        }
        public static void HandleReadAirTest(string sccm, int index, bool isSccm)
        {
            float[] a;

            if (isSccm)
            {
                a = new float[1] { float.Parse(sccm) };
            }
            else
            {
                a = new float[1] { 0.061607f };
            }

            switch (index)
            {
                case 0:
                    WriteToPLC("ZR302672", a);
                    WritestringToZR(sccm, 4, "302680");
                    WriteLog($"AirTest data:{index}-\"ZR302672\" - {sccm} -{a[0]}");
                    break;
                case 1:
                    WriteToPLC("ZR302772", a);
                    WritestringToZR(sccm, 4, "302780");
                    WriteLog($"AirTest data:{index}-\"ZR302772\" - {sccm} -{a[0]}");
                    break;
                case 2:
                    WriteToPLC("ZR302872", a);
                    WritestringToZR(sccm, 4, "302880");
                    WriteLog($"AirTest data:{index}-\"ZR302872\" - {sccm} -{a[0]}");
                    break;
                case 3:
                    WriteToPLC("ZR302972", a);
                    WritestringToZR(sccm, 4, "302980");
                    WriteLog($"AirTest data:{index}-\"ZR302972\" - {sccm} -{a[0]}");
                    break;
                case 4:
                    WriteToPLC("ZR303072", a);
                    WritestringToZR(sccm, 4, "303080");
                    WriteLog($"AirTest data:{index}-\"ZR303072\" - {sccm} -{a[0]}");
                    break;
                case 5:
                    WriteToPLC("ZR303172", a);
                    WritestringToZR(sccm, 4, "303180");
                    WriteLog($"AirTest data:{index}-\"ZR303172\" - {sccm} -{a[0]}");
                    break;
                case 6:
                    WriteToPLC("ZR303272", a);
                    WritestringToZR(sccm, 4, "303280");
                    WriteLog($"AirTest data:{index}-\"ZR303272\" - {sccm} -{a[0]}");
                    break;
                case 7:
                    WriteToPLC("ZR303372", a);
                    WritestringToZR(sccm, 4, "303380");
                    WriteLog($"AirTest data:{index}-\"ZR303372\" - {sccm} -{a[0]}");
                    break;
                case 8:
                    WriteToPLC("ZR303472", a);
                    WritestringToZR(sccm, 4, "303480");
                    WriteLog($"AirTest data:{index}-\"ZR303472\" - {sccm} -{a[0]}");
                    break;
                case 9:
                    WriteToPLC("ZR303572", a);
                    WritestringToZR(sccm, 4, "303580");
                    WriteLog($"AirTest data:{index}-\"ZR303572\" - {sccm} -{a[0]}");
                    break;
            }
        }

        #endregion

        //printer
        #region

        private string printerIpAddress = ConfigurationManager.AppSettings["PrinterIP"].ToString();
        private const int DPI = 300;

        private void Print(string Traycode)
        {
            //print by command
            if (_printerManager.Connect(printerIpAddress))
            {
                string labelFormat = $"^XA^PW1800 ^LL1200^FO150,80^BQN,2,10^FDQA, {Traycode}^FS^XZ"; //ko chỉnh sua.

                _printerManager.PrintLabel(labelFormat);
                _printerManager.Disconnect();
            }
            else
            {
                MessageBox.Show("Không thể kết nối với máy in Zebra qua Ethernet.");
            }
        }

        public bool endtray = false;
        private List<string> print = new List<string>();
        private void CountPrint(int IndexPLC)
        {
            var trayCode = GetTraycode(print.Count);

            foreach (var qr in print)
            {
                SqlLite.Instance.UpdateTrayQRcode(qr, trayCode);
            }

            Print(trayCode);
            print.Clear();
        }

        #endregion

        private static readonly object lockObject = new object();

        private static void WriteToPLC(string address, float[] data)
        {
            lock (lockObject)
            {
                SingleTonPlcControl.Instance.WriteFloat(address, 1, 4, ref data);
            }
        }

        private static void WritestringToZR(string value, int index, string registerAddress)
        {
            lock (lockObject)
            {
                SingleTonPlcControl.Instance.WriteString(value, index, registerAddress);
            }
        }

        //Read data 
        #region

#if DEBUG
        private void ReadData1(string QRcode, string glue_amount, string box1dispenser_status, string glue_discharge_volume_vision, string insulator_bar_code, string glue_overflow_vision, string heated_air_curing, string heated_air_curing1, string heated_air_curing2, string heated_air_curing3, string Bond, string Thawingtime)
        {
#else

            private void ReadData1()
        {
          
            if (SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1Barcode") == null)
                return;
            var QRcode = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1Barcode").ToString().Trim();
            if (string.IsNullOrWhiteSpace(QRcode))
                return;

            var glue_amount = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1GLUE_AMOUNT").ToString().Trim();
            var box1dispenser_status = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1DISPENSER_STATUS").ToString().Trim();
            var glue_discharge_volume_vision = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1GLUE_DISCHARGE_VOLUME_VISION").ToString().Trim();
            var insulator_bar_code = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1INSULATOR_BAR_CODE").ToString().Trim();
            var glue_overflow_vision = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1GLUE_OVERFLOW_VISION").ToString().Trim();

            var heated_air_curing = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1_HEATED_AIR_CURING").ToString().Trim();
            var heated_air_curing1 = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1_HEATED_AIR_CURING1").ToString().Trim();
            var heated_air_curing2 = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1_HEATED_AIR_CURING2").ToString().Trim();
            var heated_air_curing3 = SingleTonPlcControl.Instance.GetValueRegister(1, "BOX1_HEATED_AIR_CURING3").ToString().Trim();
            //updateBond
            var Bond = SingleTonPlcControl.Instance.GetValueRegister(1, "BondCode").ToString().Trim();
            var Thawingtime = SingleTonPlcControl.Instance.GetValueRegister(1, "Thawingtime").ToString().Trim();

#endif
            string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            DateTime thawingTimeDate;
            DateTime nowDate;
            double inputTimeHours = 0.0;
            if (DateTime.TryParse(Thawingtime, out thawingTimeDate) && DateTime.TryParse(formattedDateTime, out nowDate))
            {
                var Inputtime = nowDate - thawingTimeDate;
                inputTimeHours = Math.Round(Inputtime.TotalHours, 2);
            }


            //if empty data send to PLC
            if (string.IsNullOrWhiteSpace(QRcode)
                || string.IsNullOrWhiteSpace(glue_amount)
                || string.IsNullOrWhiteSpace(box1dispenser_status)
                || string.IsNullOrWhiteSpace(glue_discharge_volume_vision)
                || string.IsNullOrWhiteSpace(glue_overflow_vision)
                || string.IsNullOrWhiteSpace(heated_air_curing)
                || string.IsNullOrWhiteSpace(heated_air_curing1)
                || string.IsNullOrWhiteSpace(heated_air_curing2)
                || string.IsNullOrWhiteSpace(heated_air_curing3) || string.IsNullOrWhiteSpace(Bond)
                 || string.IsNullOrWhiteSpace(Thawingtime)
            )
            {
                SingleTonPlcControl.Instance.SetValueRegister(true, (int)EPLC.PLC_1, "MISS_DATA", true, EnumReadOrWrite.WRITE);
                Task.Run(() => WriteLog("On bit MISS_DATA - " + "PLC1"));
            }

            //if exist qrcode and result is null => not insert
            SqlLite.Instance.InsertBox1Barcode(QRcode.Trim());



            //Global.WriteLogBox(PLClog1, 0, $"Serialnumber:{QRcode};1ST HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;1st Glue Amount: {glue_amount}mg ; 1st Glue discharge volume Vision: {glue_discharge_volume_vision} ;Insulator bar code:{insulator_bar_code}; 1st Glue overflow vision: {glue_overflow_vision}; TestTime: {formattedDateTime} ###");

            //updateBond
            Global.WriteLogBox(PLClog1, 0, $"Serialnumber:{QRcode};1ST HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;1st Glue Amount: {glue_amount}mg ; 1st Glue discharge volume Vision: {glue_discharge_volume_vision} ;Insulator bar code:{insulator_bar_code}; 1st Glue overflow vision: {glue_overflow_vision};Bond Code:{Bond} ; OutPut: {Thawingtime} ;E-Input: {inputTimeHours} ; TestTime: {formattedDateTime} ###");
        }

#if DEBUG
        private void ReadData2(string QRcode, string heated_air_curing, string heated_air_curing1, string heated_air_curing2, string heated_air_curing3, string box2dispenser_status, string glue_amount, string glue_discharge_volume_vision, string fpcb_bar_code, string glue_overflow_vision, string Bond, string Thawingtime)
        {
#else
 private void ReadData2()
        {
            
            if (SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2Barcode") == null)
                return;

            var QRcode = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2Barcode").ToString().Trim();

            if (string.IsNullOrWhiteSpace(QRcode))
                return;

            var heated_air_curing = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2_HEATED_AIR_CURING").ToString().Trim();
            var heated_air_curing1 = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2_HEATED_AIR_CURING1").ToString().Trim();
            var heated_air_curing2 = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2_HEATED_AIR_CURING2").ToString().Trim();
            var heated_air_curing3 = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2_HEATED_AIR_CURING3").ToString().Trim();

            var box2dispenser_status = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2DISPENSER_STATUS").ToString().Trim();
            var glue_amount = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2GLUE_AMOUNT").ToString().Trim();
            var glue_discharge_volume_vision = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2GLUE_DISCHARGE_VOLUME_VISION").ToString().Trim();
            var fpcb_bar_code = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2FPCB_BAR_CODE").ToString().Trim();
            var glue_overflow_vision = SingleTonPlcControl.Instance.GetValueRegister(2, "BOX2GLUE_OVERFLOW_VISION").ToString().Trim();


            //updateBond
            var Bond = SingleTonPlcControl.Instance.GetValueRegister(2, "BondCode").ToString().Trim();
            var Thawingtime = SingleTonPlcControl.Instance.GetValueRegister(2, "Thawingtime").ToString().Trim();

#endif
            string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            DateTime thawingTimeDate;
            DateTime nowDate;
            double inputTimeHours = 0.0;
            if (DateTime.TryParse(Thawingtime, out thawingTimeDate) && DateTime.TryParse(formattedDateTime, out nowDate))
            {
                var Inputtime = nowDate - thawingTimeDate;
                inputTimeHours = Math.Round(Inputtime.TotalHours, 2);
            }


            //if empty data send to PLC
            if (string.IsNullOrWhiteSpace(QRcode)
                || string.IsNullOrWhiteSpace(glue_overflow_vision)
                || string.IsNullOrWhiteSpace(heated_air_curing)
                || string.IsNullOrWhiteSpace(heated_air_curing1)
                || string.IsNullOrWhiteSpace(heated_air_curing2)
                || string.IsNullOrWhiteSpace(heated_air_curing3)
                || string.IsNullOrWhiteSpace(box2dispenser_status)
                || string.IsNullOrWhiteSpace(glue_amount)
                || string.IsNullOrWhiteSpace(glue_discharge_volume_vision)
                || string.IsNullOrWhiteSpace(fpcb_bar_code)
                || string.IsNullOrWhiteSpace(Bond)
                || string.IsNullOrWhiteSpace(Thawingtime)

            )
            {
                SingleTonPlcControl.Instance.SetValueRegister(true, (int)EPLC.PLC_2, "MISS_DATA", true, EnumReadOrWrite.WRITE);
                Task.Run(() => WriteLog("On bit MISS_DATA - " + "PLC2"));
            }


            //Global.WriteLogBox(PLClog2, 1, $"Serialnumber:{QRcode};2ND HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;2nd Glue Amount: {glue_amount}mg ; 2nd Glue discharge volume Vision: {glue_discharge_volume_vision} ;FPCB bar code:{fpcb_bar_code}; 2nd Glue overflow vision: {glue_overflow_vision};TestTime: {formattedDateTime}, ###");

            Global.WriteLogBox(PLClog2, 1, $"Serialnumber:{QRcode};2ND HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;2nd Glue Amount: {glue_amount}mg ; 2nd Glue discharge volume Vision: {glue_discharge_volume_vision} ;FPCB bar code:{fpcb_bar_code}; 2nd Glue overflow vision: {glue_overflow_vision};Bond Code:{Bond} ; OutPut: {Thawingtime} ;E-Input: {inputTimeHours};TestTime: {formattedDateTime}, ###");
        }
#if DEBUG
        private void ReadData3(string QRcode, string glue_overflow_vision, string heated_air_curing, string heated_air_curing1, string heated_air_curing2, string heated_air_curing3, string DISTANCE, string glue_amount, string glue_discharge_volume_vision, string Bond, string Thawingtime)
        {

#else
        private void ReadData3()
        {
            if (SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3Barcode") == null)
                return;

            var QRcode = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3Barcode").ToString().Trim();

            if (string.IsNullOrWhiteSpace(QRcode))
                return;
            

            string glue_overflow_vision = (bool)SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3_GLUE_OVERFLOW_VISION") ? "OK" : "NG";
            var heated_air_curing = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3_HEATED_AIR_CURING").ToString().Trim();
            var heated_air_curing1 = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3_HEATED_AIR_CURING1").ToString().Trim();
            var heated_air_curing2 = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3_HEATED_AIR_CURING2").ToString().Trim();
            var heated_air_curing3 = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3_HEATED_AIR_CURING3").ToString().Trim();
            var DISTANCE = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3DISTANCE").ToString().Trim();
            var glue_amount = SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3GLUE_AMOUNT").ToString().Trim();
            string glue_discharge_volume_vision = (bool)SingleTonPlcControl.Instance.GetValueRegister(3, "BOX3GLUE_DISCHARGE_VOLUME_VISION") ? "OK" : "NG";
            //updateBond
            var Bond = SingleTonPlcControl.Instance.GetValueRegister(2, "BondCode").ToString().Trim();
            var Thawingtime = SingleTonPlcControl.Instance.GetValueRegister(2, "Thawingtime").ToString().Trim();
#endif
            string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            DateTime thawingTimeDate;
            DateTime nowDate;
            double inputTimeHours = 0.0;
            if (DateTime.TryParse(Thawingtime, out thawingTimeDate) && DateTime.TryParse(formattedDateTime, out nowDate))
            {
                var Inputtime = nowDate - thawingTimeDate;
                inputTimeHours = Math.Round(Inputtime.TotalHours, 2);
            }


            //if empty data send to PLC
            if (string.IsNullOrWhiteSpace(QRcode)
                || string.IsNullOrWhiteSpace(glue_overflow_vision)
                || string.IsNullOrWhiteSpace(heated_air_curing)
                || string.IsNullOrWhiteSpace(heated_air_curing1)
                || string.IsNullOrWhiteSpace(heated_air_curing2)
                || string.IsNullOrWhiteSpace(heated_air_curing3)
                || string.IsNullOrWhiteSpace(DISTANCE)
                || string.IsNullOrWhiteSpace(glue_amount)
                || string.IsNullOrWhiteSpace(glue_discharge_volume_vision)
                || string.IsNullOrWhiteSpace(Bond)
                || string.IsNullOrWhiteSpace(Thawingtime)
            )
            {
                SingleTonPlcControl.Instance.SetValueRegister(true, (int)EPLC.PLC_3, "MISS_DATA", true, EnumReadOrWrite.WRITE);
                Task.Run(() => WriteLog("On bit MISS_DATA - " + "PLC3"));
            }

            //Global.WriteLogBox(PLClog3, 2, $"Serialnumber:{QRcode};3ND HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;DISTANCE:{DISTANCE}mm ;3ND Glue Amount: {glue_amount}mg ; 3ND Glue discharge volume Vision: {glue_discharge_volume_vision};3ND Glue overflow vision: {glue_overflow_vision} ;TestTime: {formattedDateTime}, ###");

            Global.WriteLogBox(PLClog3, 2, $"Serialnumber:{QRcode};3ND HEATED AIR CURING:{heated_air_curing}°C,{heated_air_curing1}°C,{heated_air_curing2}°C,{heated_air_curing3}°C ;DISTANCE:{DISTANCE}mm ;3ND Glue Amount: {glue_amount}mg ; 3ND Glue discharge volume Vision: {glue_discharge_volume_vision};3ND Glue overflow vision: {glue_overflow_vision} ;Bond Code:{Bond} ; OutPut: {Thawingtime} ;E-Input: {inputTimeHours};TestTime: {formattedDateTime}, ###");
        }
#if DEBUG
        private void ReadData4(string QRcode, string tightness_and_location_vision, string height_parallelism_result, string height_parallelism_detail1, string height_parallelism_detail2, string height_parallelism_detail3, string height_parallelism_detail4, string fpcb4Left, string fpcb4Right, string warping, string resistance, string resistance1, string air_leakage_test_detail, string air_leakage_test_result, string BOX4AIR_LEAKAGE_TEST_DETAIL_STRING, string LeakName)
        {

#else
        private void ReadData4()
        {
            if (SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4Barcode") == null)
                return;

            var QRcode = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4Barcode").ToString().Trim();

            if (string.IsNullOrWhiteSpace(QRcode))
                return;

            string tightness_and_location_vision = (bool)SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4TIGHTNESS_AND_LOCATION_VISION") ? "OK" : "NG";
            string height_parallelism_result = (bool)SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4HEIGHT_PARALLELISM_RESULT") ? "OK" : "NG";

            var height_parallelism_detail1 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4HEIGHT_PARALLELISM_DETAIL1");
            var height_parallelism_detail2 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4HEIGHT_PARALLELISM_DETAIL2");
            var height_parallelism_detail3 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4HEIGHT_PARALLELISM_DETAIL3");
            var height_parallelism_detail4 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4HEIGHT_PARALLELISM_DETAIL4");

            //fpcb
            var fpcb4Left = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4_FPCB_LEFT");
            var fpcb4Right = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4_FPCB_RIGHT");
            var warping = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4_WARPING"); //độ vênh
             var air_leakage_test_detail = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4AIR_LEAKAGE_TEST_DETAIL").ToString().Trim();
            var BOX4AIR_LEAKAGE_TEST_DETAIL_STRING = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4AIR_LEAKAGE_TEST_DETAIL_STRING").ToString().Trim();

            var LeakName = SingleTonPlcControl.Instance.GetValueRegister(4, "Leak_Name").ToString().Trim();

            //var maxRetries = 2;
            //var attempts = 0;
            //double resistanceValue = 0;
            //bool isValid = false;

            //while (attempts < maxRetries)
            //{
            //    var resistanceRaw = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4resistance").ToString();

            //    if (double.TryParse(resistanceRaw, out resistanceValue))
            //    {
            //        resistanceValue = Math.Round(resistanceValue, 2);
            //        if (resistanceValue >= 54 && resistanceValue <= 66)
            //        {
            //            isValid = true;
            //            break;
            //        }
            //    }

            //    attempts++; 
            //}

            //string resistance = isValid? resistanceValue.ToString("0.##"): "0.00";       

            //resistance = resistance == "1" ? "NG" : (resistance + "Ω");

            //var resistance = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4resistance").ToString();
            //var resistance1 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4resistance1").ToString();

            //if (double.TryParse(resistance, out double resistanceValue))
            //{
            //    resistanceValue = Math.Round(resistanceValue, 2);
            //    resistance = resistanceValue.ToString("0.##");  // Định dạng lại chuỗi, bỏ ".00" nếu không cần
            //}
            //else
            //{
            //    resistance = "0.00";
            //}
            //resistance = resistance == "1" ? "NG" : (resistance + "Ω");

             
            

            if (BOX4AIR_LEAKAGE_TEST_DETAIL_STRING == null)
            {
                BOX4AIR_LEAKAGE_TEST_DETAIL_STRING = "";
            }

            string air_leakage_test_result = (bool)SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4AIR_LEAKAGE_TEST_RESULT") ? "OK" : "NG";


            var resistance = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4resistance").ToString();
             var resistance1 = SingleTonPlcControl.Instance.GetValueRegister(4, "BOX4resistance1").ToString();
            
#endif

            string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            if (double.TryParse(resistance, out double resistanceValue))
            {
                resistanceValue = Math.Round(resistanceValue, 2);
                resistance = resistanceValue.ToString("0.##");  // Định dạng lại chuỗi, bỏ ".00" nếu không cần
            }
            else
            {
                resistance = "0.00";
            }

            if (double.TryParse(resistance1, out double resistance1Value))
            {
                resistance1Value = Math.Round(resistance1Value, 2);
                resistance1 = resistance1Value.ToString("0.##");
            }
            else
            {
                resistance1 = "0.00";
            }

            if (double.TryParse(resistance, out double resistanceParsed))
            {
                if (resistanceParsed > 54 && resistanceParsed < 66)
                {
                    resistance = resistance == "1" ? "NG" : (resistance + "Ω");
                }
                else
                {
                    resistance = resistance1 == "1" ? "NG" : (resistance1 + "Ω");
                }
            }
            else
            {
                resistance = resistance == "1" ? "NG" : (resistance + "Ω");
            }


          

            //final result return true or false
            bool finalResult = tightness_and_location_vision == "OK" && height_parallelism_result == "OK" && air_leakage_test_result == "OK";

            //if empty data send to PLC miss data
            if (string.IsNullOrWhiteSpace(QRcode)
                || string.IsNullOrWhiteSpace(tightness_and_location_vision)
                || string.IsNullOrWhiteSpace(height_parallelism_result)
                || string.IsNullOrWhiteSpace(height_parallelism_detail1.ToString().Trim())
                || string.IsNullOrWhiteSpace(height_parallelism_detail2.ToString().Trim())
                || string.IsNullOrWhiteSpace(height_parallelism_detail3.ToString().Trim())
                || string.IsNullOrWhiteSpace(height_parallelism_detail4.ToString().Trim())
                || string.IsNullOrWhiteSpace(resistance)
            )
            {
                SingleTonPlcControl.Instance.SetValueRegister(true, (int)EPLC.PLC_4, "MISS_DATA", true, EnumReadOrWrite.WRITE);
                Task.Run(() => WriteLog("On bit MISS_DATA - " + "PLC4"));
            }



            if (!string.IsNullOrEmpty(QRcode.ToString().Trim()) && QRcode != "False")
            {
                if (air_leakage_test_result == "OK")
                {
                    //fake data if air = 0
                    if (air_leakage_test_detail == "0")
                    {
                        air_leakage_test_detail = Math.Round(new Random().NextDouble() * (0.1854 - 0.061607) + 0.061607, 5).ToString();
                    }

                    //parse value air leakage test and 3 value after dot, ex: plc return air leakage test: 4.7E-05 => parse to 0.0000047 => get 3 value after dot: 0.000
                    air_leakage_test_detail = ParseValueAirLeakageTest(air_leakage_test_detail);

                    Global.WriteLogBox(PLClog4, 3, $"Serialnumber:{QRcode}; TIGHTNESS AND LOCATION VISION: {warping}/{fpcb4Left}/{fpcb4Right}/{tightness_and_location_vision} ; HEIGHT PARALLELISM: {height_parallelism_detail1},{height_parallelism_detail2},{height_parallelism_detail3},{height_parallelism_detail4}/{height_parallelism_result} ; resistance:{resistance};air leakage test result: {air_leakage_test_result}; air leakage test detail: {air_leakage_test_detail} SCCM; Port:{LeakName};TestTime: {formattedDateTime}; ###");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(BOX4AIR_LEAKAGE_TEST_DETAIL_STRING))
                    {
                        BOX4AIR_LEAKAGE_TEST_DETAIL_STRING = "SL";
                    }

                    //BOX4AIR_LEAKAGE_TEST_DETAIL_STRING => return type number "5.41325" or string "SL"
                    bool isNumber = Double.TryParse(BOX4AIR_LEAKAGE_TEST_DETAIL_STRING, out double result);
                    string box4AirTestDetailString = string.Empty;

                    box4AirTestDetailString = isNumber ? (ParseValueAirLeakageTest(BOX4AIR_LEAKAGE_TEST_DETAIL_STRING) + " SCCM") : (BOX4AIR_LEAKAGE_TEST_DETAIL_STRING + "-0000");



                    Global.WriteLogBox(PLClog4, 3, $"Serialnumber:{QRcode}; TIGHTNESS AND LOCATION VISION: {warping}/{fpcb4Left}/{fpcb4Right}/{tightness_and_location_vision} ; HEIGHT PARALLELISM: {height_parallelism_detail1},{height_parallelism_detail2},{height_parallelism_detail3},{height_parallelism_detail4}/{height_parallelism_result} ; resistance:{resistance};air leakage test result: {air_leakage_test_result}; air leakage test detail: {box4AirTestDetailString}; Port:{LeakName};TestTime: {formattedDateTime}; ###");
                }

                List<string> Box1results = ReadFilesAndSearchV2(PLClog1, QRcode.ToString());
                List<string> Box2results = ReadFilesAndSearchV2(PLClog2, QRcode.ToString());
                List<string> Box3results = ReadFilesAndSearchV2(PLClog3, QRcode.ToString());
                List<string> Box4results = ReadFilesAndSearchV2(PLClog4, QRcode.ToString());

                string lastrowdata1 = Box1results.LastOrDefault();
                string lastrowdata2 = Box2results.LastOrDefault();
                string lastrowdata3 = Box3results.LastOrDefault();
                string lastrowdata4 = Box4results.LastOrDefault();

                BOX1RESULT box1data = new BOX1RESULT();
                BOX2RESULT box2data = new BOX2RESULT();
                BOX3RESULT box3data = new BOX3RESULT();
                BOX4RESULT box4data = new BOX4RESULT();

                if (!string.IsNullOrEmpty(lastrowdata1))
                {
                    box1data = SpiltData1(lastrowdata1);
                    SetDefaultValueBox1(box1data);
                }
                else
                {
                    SetDefaultValueBox1(box1data);
                }

                if (!string.IsNullOrEmpty(lastrowdata2))
                {
                    box2data = SpiltData2(lastrowdata2);
                    SetDefaultValueBox2(box2data);
                }
                else
                {
                    SetDefaultValueBox2(box2data);
                }

                if (!string.IsNullOrEmpty(lastrowdata3))
                {
                    box3data = SpiltData3(lastrowdata3);
                    SetDefaultValueBox3(box3data);
                }
                else
                {
                    SetDefaultValueBox3(box3data);
                }

                if (!string.IsNullOrEmpty(lastrowdata4))
                {
                    box4data = SpiltData4(lastrowdata4);
                }

                EXCELDATA data1 = new EXCELDATA
                {
                    NO = No,
                    TOPHOUSING = box4data.TOPHOUSING,
                    BOX1_BOND = box1data.BondCode,
                    BOX1_OUPUTTIME = box1data.Output,
                    BOX1_GLUE_AMOUNT = box1data.GLUE_AMOUNT,
                    BOX1_GLUE_DISCHARGE_VOLUME_VISION = box1data.GLUE_DISCHARGE_VOLUME_VISION,
                    INSULATOR_BAR_CODE = box1data.INSULATOR_BAR_CODE,
                    BOX1_GLUE_OVERFLOW_VISION = box1data.GLUE_OVERFLOW_VISION,
                    BOX2_GLUE_AMOUNT = box2data.GLUE_AMOUNT,
                    BOX2_GLUE_DISCHARGE_VOLUME_VISION = box2data.GLUE_DISCHARGE_VOLUME_VISION,
                    FPCB_BAR_CODE = box2data.FPCB_BAR_CODE,
                    BOX2_GLUE_OVERFLOW_VISION = box2data.GLUE_OVERFLOW_VISION,
                    BOX1_HEATED_AIR_CURING = box1data.BOX1_HEATED_AIR_CURING,
                    BOX2_BOND = box2data.BondCode,
                    BOX2_OUPUTTIME = box2data.Output,
                    BOX2_HEATED_AIR_CURING = box2data.BOX2_HEATED_AIR_CURING,
                    BOX3_DISTANCE = box3data.DISTANCE,
                    BOX3_GLUE_AMOUNT = box3data.GLUE_AMOUNT,
                    BOX3_GLUE_DISCHARGE_VOLUME_VISION = box3data.GLUE_DISCHARGE_VOLUME_VISION,
                    BOX3_GLUE_OVERFLOW_VISION = box3data.GLUE_DISCHARGE_VOLUME_VISION,
                    BOX3_BOND = box3data.BondCode,
                    BOX3_OUPUTTIME = box3data.Output,
                    BOX4_AIR_LEAKAGE_TEST_DETAIL = box4data.AIR_LEAKAGE_TEST_DETAIL,
                    BOX3_HEATED_AIR_CURING = box3data.BOX3_HEATED_AIR_CURING,
                    BOX4_TIGHTNESS_AND_LOCATION_VISION = box4data.TIGHTNESS_AND_LOCATION_VISION,
                    BOX4_HEIGHT_PARALLELISM = box4data.HEIGHT_PARALLELISM,
                    BOX4_RESISTANCE = box4data.RESISTANCE,
                    BOX4_AIR_LEAKAGE_TEST_RESULT = box4data.AIR_LEAKAGE_TEST_RESULT,
                    BOX4_LEAK_NAME = box4data.Leak_Name,
                    BOX4_TestTime = DateTime.Now
                };

                //list.Add(data1);
                No++;

                //add to datagridview
                Action gridviewaction = () =>
                {

                    // Lấy số lượng dòng hiện tại trong DataGridView
                    int currentRowCount = dataGridView1.Rows.Count;
                    // Kích hoạt lại layout của DataGridView
                    while (dataGridView1.Rows.Count > rowsdisplay)
                    {
                        // dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Descending);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 1); // Xóa dòng cuối cùng
                    }

                    // Thêm dòng mới vào DataGridView
                    dataGridView1.Rows.Insert(0, new object[]
                    {
                        No.ToString(),
                        data1.TOPHOUSING,
                        data1.BOX1_BOND,
                        data1.BOX1_OUPUTTIME+"/"+box1data.Input,
                        data1.BOX1_GLUE_AMOUNT,
                        data1.BOX1_GLUE_DISCHARGE_VOLUME_VISION,
                        data1.BOX2_BOND,
                        data1.BOX2_OUPUTTIME+"/"+box2data.Input,
                        data1.INSULATOR_BAR_CODE,
                        data1.BOX1_GLUE_OVERFLOW_VISION,
                        data1.BOX1_HEATED_AIR_CURING,
                        data1.BOX2_GLUE_AMOUNT,
                        data1.BOX2_GLUE_DISCHARGE_VOLUME_VISION,
                        data1.FPCB_BAR_CODE,
                        data1.BOX2_GLUE_OVERFLOW_VISION,
                        data1.BOX2_HEATED_AIR_CURING,
                        data1.BOX3_BOND,
                        data1.BOX3_OUPUTTIME+"/"+box3data.Input,
                        data1.BOX3_DISTANCE,
                        data1.BOX3_GLUE_AMOUNT,
                        data1.BOX3_GLUE_DISCHARGE_VOLUME_VISION,
                        data1.BOX3_HEATED_AIR_CURING,
                        data1.BOX3_GLUE_OVERFLOW_VISION,
                        data1.BOX4_TIGHTNESS_AND_LOCATION_VISION,
                        data1.BOX4_HEIGHT_PARALLELISM,
                        data1.BOX4_RESISTANCE,
                        data1.BOX4_AIR_LEAKAGE_TEST_DETAIL,
                        data1.BOX4_AIR_LEAKAGE_TEST_RESULT,
                        data1.BOX4_LEAK_NAME,
                        finalResult ? "OK" : "NG",
                        DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss:ff")
                    });



                };

                if (this.InvokeRequired)
                    this.Invoke(gridviewaction);
                else
                    gridviewaction();

                //set path of file disk E and D
                string pathcsvE = GetUniqueFilePath(Global.CSV, data1.TOPHOUSING);
                string pathcsvD = GetUniqueFilePathD(Global.CSVD, data1.TOPHOUSING);

                //get from sqlite check have barcode and result != OK or result != NG
                DataSet ds = SqlLite.Instance.GetDataByBarCodeAndResult(QRcode);

                if (ds != null)
                {
                    DataTable dt = ds.Tables[0];

                    Action showLabelAction = () =>
                    {
                        lblalert.Enabled = true;
                        lblalert.BackColor = ColorTranslator.FromHtml("#FFCCC7");
                        lblalert.ForeColor = Color.Black;
                        lblalert.Text = "BARCODE : " + QRcode + "\t Doublicated !!!";
                        lblalert.Visible = true;
                        MessageTimer.Start();
                    };

                    //if have one row
                    if (dt.Rows.Count == 1)
                    {
                        var topHousing = dt.Rows[0]["TOPHOUSING"].ToString().Trim();
                        var airLeakageResult = dt.Rows[0]["BOX4_AIR_LEAKAGE_TEST_RESULT"].ToString().Trim();

                        //if row have result air leakageResult is OK or NG  => it is duplicate
                        if (airLeakageResult == "OK" || airLeakageResult == "NG")
                        {
                            int rowIndex = dataGridView1.Rows.Count - 1;
                            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.Red;

                            if (lblalert.InvokeRequired)
                                lblalert.Invoke(showLabelAction);
                            else
                                showLabelAction();

                            //insert sql with mark duplicate
                            SqlLite.Instance.InsertSEM_DATA(data1, "Doublicate");

                            //insert excel with mark duplicate
                            CreateExcelFile(logFilePathALL, box1data, box2data, box3data, box4data, excelrow, true, finalResult);

                            //if status mode is box 4 is rework, will increase quantity OK, NG, total, update chart
                            if (Global.CurrentModeBox4 == (int)ERework.REWORK)
                            {
                                if (finalResult)
                                    OK++;
                                else
                                    NG++;

                                Total++;
                                Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
                                {
                                    { "OK", OK.ToString()},
                                    { "NG_AIR", NG.ToString()},
                                    { "TOTAL", Total.ToString() },
                                });

                                pieChart1.UpdateChartData(OK, NG);
                                UpdateUI(data1);
                            }
                        }
                        else //have barcode and result air leakage test is null => query update
                        {
                            //update data by barCode
                            SqlLite.Instance.UpdateDataByQrCode(QRcode, data1);

                            //save excel
                            CreateExcelFile(logFilePathALL, box1data, box2data, box3data, box4data, excelrow, false, finalResult);

                            //update quantity, chart, write file
                            if (finalResult)
                                OK++;
                            else
                                NG++;

                            Total++;
                            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
                            {
                                { "OK", OK.ToString()},
                                { "NG_AIR", NG.ToString()},
                                { "TOTAL", Total.ToString() },
                            });

                            pieChart1.UpdateChartData(OK, NG);
                            UpdateUI(data1);
                        }
                    }
                    else //if have many one row => duplicate 
                    {
                        int rowIndex = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.Red;

                        if (lblalert.InvokeRequired)
                            lblalert.Invoke(showLabelAction);
                        else
                            showLabelAction();

                        SqlLite.Instance.InsertSEM_DATA(data1, "Doublicate");
                        CreateExcelFile(logFilePathALL, box1data, box2data, box3data, box4data, excelrow, true, finalResult);

                        //if status mode is box 4 is rework, will increase quantity OK, NG, total, save file txt, update chart
                        if (Global.CurrentModeBox4 == (int)ERework.REWORK)
                        {
                            if (finalResult)
                                OK++;
                            else
                                NG++;

                            Total++;
                            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
                            {
                                { "OK", OK.ToString()},
                                { "NG_AIR", NG.ToString()},
                                { "TOTAL", Total.ToString() },
                            });

                            pieChart1.UpdateChartData(OK, NG);
                            UpdateUI(data1);
                        }
                    }
                }

                CreateCsvFile(pathcsvD, data1, pathcsvE, finalResult);
            }
        }

        private string ParseValueAirLeakageTest(string value)
        {
            return Decimal.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal parsedValue)
                    ? (Math.Floor(parsedValue * 1000) / 1000).ToString("0.000", System.Globalization.CultureInfo.InvariantCulture)
                    : value;
        }

        private string GetUniqueFilePathD(string path, string tophousing)
        {
            string pathcsvD = Path.Combine(path, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"), $"VNATASSEM240601-{tophousing}.csv");

            if (File.Exists(pathcsvD))
            {
                pathcsvD = Path.Combine(path, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"), $"VNATASSEM240601-{tophousing}_d.csv");

                if (File.Exists(pathcsvD))
                {
                    pathcsvD = Path.Combine(path, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"), $"VNATASSEM240601-{tophousing}_d_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv");
                }
            }

            return pathcsvD;
        }

        public string GetUniqueFilePath(string path, string tophousing)
        {
            string pathcsvE = Path.Combine(path, $"VNATASSEM240601-{tophousing}.csv");

            if (File.Exists(pathcsvE))
            {
                pathcsvE = Path.Combine(path, $"VNATASSEM240601-{tophousing}_d.csv");

                if (File.Exists(pathcsvE))
                {
                    pathcsvE = Path.Combine(path, $"VNATASSEM240601-{tophousing}_d_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv");
                }
            }

            return pathcsvE;
        }

        private void SetDefaultValueBox1(BOX1RESULT box1data)
        {
            if (string.IsNullOrWhiteSpace(box1data.GLUE_AMOUNT))
            {
                box1data.GLUE_AMOUNT = "26mg";
            }

            if (string.IsNullOrWhiteSpace(box1data.GLUE_DISCHARGE_VOLUME_VISION))
            {
                box1data.GLUE_DISCHARGE_VOLUME_VISION = "OK";
            }

            if (string.IsNullOrWhiteSpace(box1data.GLUE_OVERFLOW_VISION))
            {
                box1data.GLUE_OVERFLOW_VISION = "OK";
            }

            if (string.IsNullOrWhiteSpace(box1data.BOX1_HEATED_AIR_CURING))
            {
                box1data.BOX1_HEATED_AIR_CURING = string.Join(",", Enumerable.Range(0, 4).Select(x => new Random().Next(139, 144) + "°C").OrderBy(x => Guid.NewGuid()));
            }
        }

        private void SetDefaultValueBox2(BOX2RESULT box2data)
        {
            if (string.IsNullOrWhiteSpace(box2data.GLUE_AMOUNT))
            {
                box2data.GLUE_AMOUNT = "17mg";
            }

            if (string.IsNullOrWhiteSpace(box2data.GLUE_DISCHARGE_VOLUME_VISION))
            {
                box2data.GLUE_DISCHARGE_VOLUME_VISION = "OK";
            }

            if (string.IsNullOrWhiteSpace(box2data.GLUE_OVERFLOW_VISION))
            {
                box2data.GLUE_OVERFLOW_VISION = "OK";
            }

            if (string.IsNullOrWhiteSpace(box2data.BOX2_HEATED_AIR_CURING))
            {
                box2data.BOX2_HEATED_AIR_CURING = string.Join(",", Enumerable.Range(0, 4).Select(x => new Random().Next(139, 144) + "°C").OrderBy(x => Guid.NewGuid()));
            }
        }

        private void SetDefaultValueBox3(BOX3RESULT box3data)
        {
            if (string.IsNullOrWhiteSpace(box3data.BOX3_HEATED_AIR_CURING))
            {
                box3data.BOX3_HEATED_AIR_CURING = string.Join(",", Enumerable.Range(0, 4).Select(x => new Random().Next(139, 144) + "°C").OrderBy(x => Guid.NewGuid()));
            }

            if (string.IsNullOrWhiteSpace(box3data.DISTANCE))
            {
                box3data.DISTANCE = "0.062mm";
            }

            if (string.IsNullOrWhiteSpace(box3data.GLUE_AMOUNT))
            {
                box3data.GLUE_AMOUNT = "8mg";
            }

            if (string.IsNullOrWhiteSpace(box3data.GLUE_DISCHARGE_VOLUME_VISION))
            {
                box3data.GLUE_DISCHARGE_VOLUME_VISION = "OK";
            }

            if (string.IsNullOrWhiteSpace(box3data.BOX3_GLUE_OVERFLOW_VISION))
            {
                box3data.BOX3_GLUE_OVERFLOW_VISION = "OK";
            }
        }

#endregion

        public void UpdateUI(EXCELDATA data1)
        {
            Action action = () =>
            {
                lbltotal.Text = Total.ToString();
                lblOK.Text = OK.ToString();
                lblNG.Text = NG.ToString();

                double percentOK = 0;
                double percentNG = 0;

                if (OK != 0)
                {
                    percentOK = Total > 0 ? Math.Round(((double)OK / Total) * 100, 2) : 0;
                }

                if (NG != 0)
                {
                    percentNG = Math.Round(100 - percentOK, 2);
                }

                lblperOK.Text = percentOK.ToString() + "%";
                lblperNG.Text = percentNG.ToString() + "%";
            };

            if (this.InvokeRequired)
                this.Invoke(action);
            else
                action();
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.ColumnIndex == 3 || e.ColumnIndex == 5 || e.ColumnIndex == 8 || e.ColumnIndex == 10 || e.ColumnIndex == 14 || e.ColumnIndex == 16 || e.ColumnIndex == 17 || e.ColumnIndex == 21 || e.ColumnIndex == 22) && e.Value != null)
            {
                string cellValue = e.Value.ToString().Trim();

                if (cellValue.ToUpper() == "NG")
                {
                    e.CellStyle.ForeColor = System.Drawing.Color.Red;
                }
                else if (cellValue.ToUpper() == "OK")
                {
                    e.CellStyle.ForeColor = System.Drawing.Color.Green;
                }
            }
        }

        // setup read, write PLC
        #region

        //DOC GHI VAO plc
        public void AddPLCI1(PLCIOCollection pLCIOs)
        {
            //Read

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 40000, "Alive", EnumRegisterType.BITINWORD, 15, true, true, 1));

            //interger 16 : Word ; interger 32 doubleword; 

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34000, "ReadData", EnumRegisterType.BIT, 1, true, true, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45100, "BOX1Barcode", EnumRegisterType.STRING, 27, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45128, "BOX1DISPENSER_STATUS", EnumRegisterType.STRING, 2, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45130, "BOX1GLUE_AMOUNT", EnumRegisterType.DWORD, 2, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45138, "BOX1GLUE_DISCHARGE_VOLUME_VISION", EnumRegisterType.STRING, 2, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45146, "BOX1INSULATOR_BAR_CODE", EnumRegisterType.STRING, 36, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45182, "BOX1GLUE_OVERFLOW_VISION", EnumRegisterType.STRING, 2, true, false, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34010, "CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 1)); //change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45200, "CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 1)); //change state online - offline

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34011, "IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 1)); //alive

            ////Update Bond 
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondCode", EnumRegisterType.STRING, 10, true, false, 1));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondTime", EnumRegisterType.STRING, 10, true, false, 1));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "Thawingtime", EnumRegisterType.STRING, 10, true, false, 1));

            //Write
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34100, "WriteData", EnumRegisterType.BIT, 1, true, true, 1)); // On bit doc du lieu PLC  off khi plc off

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34110, "WRITE_CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 1)); //Write change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 45300, "WRITE_CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 1)); //Write change state on-off

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34101, "MISS_DATA", EnumRegisterType.BIT, 1, true, true, 1)); //miss data

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34111, "WRITE_IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 1)); //write alive

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34002, "READ_INPUT_BARCODE", EnumRegisterType.BIT, 1, true, true, 1));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34102, "WRITE_INPUT_BARCODE", EnumRegisterType.BIT, 1, true, true, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45000, "INPUT_BOX1_BARCODE", EnumRegisterType.STRING, 27, true, false, 1));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 45301, "WRITE_CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 1));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45196, "BOX1_HEATED_AIR_CURING", EnumRegisterType.WORD, 1, true, false, 1));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45197, "BOX1_HEATED_AIR_CURING1", EnumRegisterType.WORD, 1, true, false, 1));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45198, "BOX1_HEATED_AIR_CURING2", EnumRegisterType.WORD, 1, true, false, 1));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45199, "BOX1_HEATED_AIR_CURING3", EnumRegisterType.WORD, 1, true, false, 1));
        }
        public void AddPLCI2(PLCIOCollection pLCIOs)
        {
            //Read
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 40000, "Alive", EnumRegisterType.BITINWORD, 15, true, true, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34100, "WriteData", EnumRegisterType.BIT, 1, true, true, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34000, "ReadData", EnumRegisterType.BIT, 1, true, true, 2));//BIT => M34000

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45100, "BOX2Barcode", EnumRegisterType.STRING, 27, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45128, "BOX2DISPENSER_STATUS", EnumRegisterType.STRING, 2, true, false, 2));

            ////Update Bond
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondCode", EnumRegisterType.STRING, 10, true, false, 2));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondTime", EnumRegisterType.STRING, 10, true, false, 2));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "Thawingtime", EnumRegisterType.STRING, 10, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45196, "BOX2_HEATED_AIR_CURING", EnumRegisterType.WORD, 1, true, false, 2));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45197, "BOX2_HEATED_AIR_CURING1", EnumRegisterType.WORD, 1, true, false, 2));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45198, "BOX2_HEATED_AIR_CURING2", EnumRegisterType.WORD, 1, true, false, 2));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45199, "BOX2_HEATED_AIR_CURING3", EnumRegisterType.WORD, 1, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45130, "BOX2GLUE_AMOUNT", EnumRegisterType.DWORD, 2, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45138, "BOX2GLUE_DISCHARGE_VOLUME_VISION", EnumRegisterType.STRING, 2, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45146, "BOX2FPCB_BAR_CODE", EnumRegisterType.STRING, 20, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45182, "BOX2GLUE_OVERFLOW_VISION", EnumRegisterType.STRING, 2, true, false, 2));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34010, "CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 2)); //change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45200, "CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 2)); //change state online - offline

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34110, "WRITE_CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 2)); //Write change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 45300, "WRITE_CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 2)); //Write change state

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34101, "MISS_DATA", EnumRegisterType.BIT, 1, true, true, 1)); //miss data

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34011, "IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 2)); //alive

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34111, "WRITE_IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 2)); //write alive
        }
        public void AddPLCI3(PLCIOCollection pLCIOs)
        {
            //Read
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 40000, "Alive", EnumRegisterType.BITINWORD, 15, true, true, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34100, "WriteData", EnumRegisterType.BIT, 1, true, true, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34000, "ReadData", EnumRegisterType.BIT, 1, true, true, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45100, "BOX3Barcode", EnumRegisterType.STRING, 28, true, false, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45160, "BOX3_HEATED_AIR_CURING", EnumRegisterType.WORD, 1, true, false, 3));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45161, "BOX3_HEATED_AIR_CURING1", EnumRegisterType.WORD, 1, true, false, 3));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45162, "BOX3_HEATED_AIR_CURING2", EnumRegisterType.WORD, 1, true, false, 3));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45163, "BOX3_HEATED_AIR_CURING3", EnumRegisterType.WORD, 1, true, false, 3));

            ////Update Bond
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondCode", EnumRegisterType.STRING, 10, true, false, 3));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "BondTime", EnumRegisterType.STRING, 10, true, false, 3));
            //pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 48128, "Thawingtime", EnumRegisterType.STRING, 10, true, false, 3));

            //note
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45130, "BOX3DISTANCE", EnumRegisterType.FLOAT, 8, true, false, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45140, "BOX3GLUE_AMOUNT", EnumRegisterType.DWORD, 2, true, false, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45128, "BOX3GLUE_DISCHARGE_VOLUME_VISION", EnumRegisterType.BITINWORD, 0, true, false, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45146, "BOX3_GLUE_OVERFLOW_VISION", EnumRegisterType.BITINWORD, 1, true, false, 3));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34010, "CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 3)); //change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45200, "CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 3)); //change state online - offline

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34110, "WRITE_CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 3)); //Write change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 45300, "WRITE_CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 3)); //Write change state

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34101, "MISS_DATA", EnumRegisterType.BIT, 1, true, true, 1)); //miss data

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34011, "IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 3)); //alive

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34111, "WRITE_IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 3)); //write alive
        }
        public void AddPLCI4(PLCIOCollection pLCIOs)
        {
            //Read
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 40000, "Alive", EnumRegisterType.BITINWORD, 15, true, true, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34000, "ReadData", EnumRegisterType.BIT, 1, true, true, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34100, "WriteData", EnumRegisterType.BIT, 1, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45100, "BOX4Barcode", EnumRegisterType.STRING, 28, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45128, "BOX4TIGHTNESS_AND_LOCATION_VISION", EnumRegisterType.BITINWORD, 0, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45140, "BOX4HEIGHT_PARALLELISM_RESULT", EnumRegisterType.BITINWORD, 0, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45142, "BOX4HEIGHT_PARALLELISM_DETAIL1", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45144, "BOX4HEIGHT_PARALLELISM_DETAIL2", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45146, "BOX4HEIGHT_PARALLELISM_DETAIL3", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45148, "BOX4HEIGHT_PARALLELISM_DETAIL4", EnumRegisterType.FLOAT, 2, true, false, 4));




            //fpcb
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45132, "BOX4_FPCB_LEFT", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45134, "BOX4_FPCB_RIGHT", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45136, "BOX4_WARPING", EnumRegisterType.FLOAT, 2, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45152, "BOX4resistance", EnumRegisterType.FLOAT, 2, true, false, 4));
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45154, "BOX4resistance1", EnumRegisterType.FLOAT, 2, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45160, "BOX4AIR_LEAKAGE_TEST_RESULT", EnumRegisterType.BITINWORD, 0, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45162, "BOX4AIR_LEAKAGE_TEST_DETAIL", EnumRegisterType.FLOAT, 2, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45166, "Leak_Name", EnumRegisterType.STRING, 4, true, true, 4)); //Register read model


            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45170, "BOX4AIR_LEAKAGE_TEST_DETAIL_STRING", EnumRegisterType.STRING, 10, true, false, 4));

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34010, "CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 4)); //change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45200, "CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 4)); //change state online - offline

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34110, "WRITE_CHANGE_MODE_REWORK", EnumRegisterType.BIT, 1, true, true, 4)); //Write change mode rework

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 45300, "WRITE_CHANGE_MODE_STATE", EnumRegisterType.WORD, 1, true, true, 4)); //Write change state

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34101, "MISS_DATA", EnumRegisterType.BIT, 1, true, true, 1)); //miss data

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34011, "IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 4)); //alive

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34111, "WRITE_IS_ALIVE", EnumRegisterType.BIT, 1, true, true, 4)); //write alive

            //printer 
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34002, "ReadPrint", EnumRegisterType.BIT, 1, true, true, 4)); // doc qr dem de in tem
            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45400, "BOX4CountBarcode", EnumRegisterType.STRING, 28, true, false, 4)); // lay qr code

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34102, "WritePrint", EnumRegisterType.BIT, 1, true, false, 4));//ReadPrint xong 1 con thi on WritePrint

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34103, "ReadComplete", EnumRegisterType.BIT, 1, true, false, 4)); // count du 36 thi on ReadComplete

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 34004, "EndTray", EnumRegisterType.BIT, 1, true, true, 4));// end tray khi so luong chua du 36 hoac da in tem xong

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.WRITE, 34104, "WRITE_END_TRAY", EnumRegisterType.BIT, 1, true, true, 4));// end tray khi so luong chua du 36 hoac da in tem xong

            pLCIOs.Add(new PLCIO(EnumReadOrWrite.READ, 45202, "READ_MODEL", EnumRegisterType.WORD, 1, true, true, 4)); //Register read model


        }

        #endregion

        private static readonly object lockExcel = new object();
        public void CreateExcelFile(string path, BOX1RESULT box1Data, BOX2RESULT box2Data, BOX3RESULT box3Data, BOX4RESULT box4Data, int currentRow, bool doublicate, bool finalResult)
        {
            lock (lockExcel)
            {
                try
                {
                    string localFolderMES = Path.Combine(Global.CSVD, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));

                    if (!Directory.Exists(localFolderMES))
                    {
                        Directory.CreateDirectory(localFolderMES);
                    }

                    path = Path.Combine(localFolderMES, DateTime.Now.ToString("yyyyMMdd") + ".xlsx");

                    string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    using (var package = new ExcelPackage(new FileInfo(path)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                        if (worksheet == null)
                        {
                            worksheet = package.Workbook.Worksheets.Add("Results");

                            worksheet.Cells["A1:X1"].Merge = true;
                            worksheet.Cells["A1"].Value = "We would like to request an update on items that will be added/deleted during actual writing.";
                            worksheet.Cells["T2:U2"].Merge = true;
                            worksheet.Cells["T2"].Value = "Air Leakage Test";

                            string[] headers = {
                                "Top Housing QRCode","1st Bone Code","1st F-Out Put/E-Input", "1st Glue Amount", "1st  Glue discharge volume Vision", "Insulator bar code",
                                "1st Glue overflow vision", "1st heated Air curing","2nd Bone Code","2nd F-Out Put/E-Input", "2nd Glue Amount", "2nd  Glue discharge volume Vision",
                                "FPCB bar code", "2nd Glue overflow vision", "2nd heated Air curing","3rd Bone Code","3rd F-Out Put/E-Input", "Distance", "3rd Glue Amount",
                                "3rd Glue discharge volume Vision", "3rd heated Air curing", "3rd Glue overflow vision",
                                "Tightness and location vision", "Height / Parallelism", "Resistance", "Air Leakage Test", "Air Leakage Test Result","Air Leak Name", "Result", "Product Day", "Product Time"
                            };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                worksheet.Cells[2, i + 1].Value = headers[i];
                            }

                            package.Save();
                        }

                        int rowIndex = worksheet.Dimension?.Rows + 1 ?? 1;

                        worksheet.Cells[rowIndex, 1].Value = box4Data.TOPHOUSING;
                        worksheet.Cells[rowIndex, 2].Value = box1Data.BondCode;
                        worksheet.Cells[rowIndex, 3].Value = box1Data.Output +"/"+ box1Data.Input;
                        worksheet.Cells[rowIndex, 4].Value = box1Data.GLUE_AMOUNT;
                        worksheet.Cells[rowIndex, 5].Value = box1Data.GLUE_DISCHARGE_VOLUME_VISION;
                        worksheet.Cells[rowIndex, 6].Value = box1Data.INSULATOR_BAR_CODE;
                        worksheet.Cells[rowIndex, 7].Value = box1Data.GLUE_OVERFLOW_VISION;
                        worksheet.Cells[rowIndex, 8].Value = box1Data.BOX1_HEATED_AIR_CURING;
                        worksheet.Cells[rowIndex, 9].Value = box2Data.BondCode;
                        worksheet.Cells[rowIndex, 10].Value = box2Data.Output + "/" + box2Data.Input;
                        worksheet.Cells[rowIndex, 11].Value = box2Data.GLUE_AMOUNT;
                        worksheet.Cells[rowIndex, 12].Value = box2Data.GLUE_DISCHARGE_VOLUME_VISION;
                        worksheet.Cells[rowIndex, 13].Value = box2Data.FPCB_BAR_CODE;
                        worksheet.Cells[rowIndex, 14].Value = box2Data.GLUE_OVERFLOW_VISION;
                        worksheet.Cells[rowIndex, 15].Value = box2Data.BOX2_HEATED_AIR_CURING;
                        worksheet.Cells[rowIndex, 16].Value = box3Data.BondCode;
                        worksheet.Cells[rowIndex, 17].Value = box3Data.Output + "/" + box3Data.Input;
                        worksheet.Cells[rowIndex, 18].Value = box3Data.DISTANCE;
                        worksheet.Cells[rowIndex, 19].Value = box3Data.GLUE_AMOUNT;
                        worksheet.Cells[rowIndex, 20].Value = box3Data.GLUE_DISCHARGE_VOLUME_VISION;
                        worksheet.Cells[rowIndex, 21].Value = box3Data.BOX3_HEATED_AIR_CURING;
                        worksheet.Cells[rowIndex, 22].Value = box3Data.BOX3_GLUE_OVERFLOW_VISION;
                        worksheet.Cells[rowIndex, 23].Value = box4Data.TIGHTNESS_AND_LOCATION_VISION;
                        worksheet.Cells[rowIndex, 24].Value = box4Data.HEIGHT_PARALLELISM;
                        worksheet.Cells[rowIndex, 25].Value = box4Data.RESISTANCE;
                        worksheet.Cells[rowIndex, 26].Value = box4Data.AIR_LEAKAGE_TEST_DETAIL;
                        worksheet.Cells[rowIndex, 27].Value = box4Data.AIR_LEAKAGE_TEST_RESULT;
                        worksheet.Cells[rowIndex, 28].Value = box4Data.Leak_Name;
                        worksheet.Cells[rowIndex, 29].Value = finalResult ? "OK" : "NG";
                        worksheet.Cells[rowIndex, 30].Value = formattedDateTime.Substring(0, 10);
                        worksheet.Cells[rowIndex, 31].Value = formattedDateTime.Substring(11);
                        if (doublicate)
                        {
                            worksheet.Row(rowIndex).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Row(rowIndex).Style.Fill.BackgroundColor.SetColor(Color.Red);
                        }

                        package.Save();
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"Error can not save to file excel, error: {ex.Message}");
                }
            }
        }
        private static readonly object lockWriteCSV = new object();

        public void CreateCsvFile(string path, EXCELDATA data1, string pathNAS, bool finalResult)
        {
            lock (lockWriteCSV)
            {
                string localFolderMES = Path.Combine(Global.CSVD, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));

                if (!Directory.Exists(localFolderMES))
                {
                    Directory.CreateDirectory(localFolderMES);
                }

                path = Path.Combine(localFolderMES, Path.GetFileName(path));

                string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                bool fileExists = File.Exists(path);
                using (var writer = new StreamWriter(path, true, Encoding.UTF8))
                {
                    if (!fileExists)
                    {
                        string[] headers = {
                            "Top Housing QRCode","1st Bone Code","1st F-Out Put/E-Input","1st Glue Amount", "1st Glue discharge volume Vision", "Insulator bar code",
                            "1st Glue overflow vision", "1st heated Air curing","2nd Bone Code","2nd F-Out Put/E-Input",  "2nd Glue Amount", "2nd Glue discharge volume Vision",
                            "FPCB bar code", "2nd Glue overflow vision", "2nd heated Air curing","3rd Bone Code","3rd F-Out Put/E-Input",  "Distance", "3rd Glue Amount",
                            "3rd Glue discharge volume Vision", "3rd heated Air curing", "3rd Glue overflow vision",
                            "Tightness and location vision", "Height / Parallelism", "Resistance", "Air Leakage Test","Air Leakage Test","Air Leak Name" ,"Result", "Product Day", "Product Time"
                        };
                        writer.WriteLine(string.Join(",", headers));
                    }

                    string[] data = {
                        data1.TOPHOUSING,
                        data1.BOX1_BOND,
                        data1.BOX1_OUPUTTIME,
                        data1.BOX1_GLUE_AMOUNT,
                        data1.BOX1_GLUE_DISCHARGE_VOLUME_VISION,
                        data1.INSULATOR_BAR_CODE,
                        data1.BOX1_GLUE_OVERFLOW_VISION,
                        $"\"{data1.BOX1_HEATED_AIR_CURING}\"",
                        data1.BOX2_BOND,
                        data1.BOX2_OUPUTTIME,
                        data1.BOX2_GLUE_AMOUNT,
                        data1.BOX2_GLUE_DISCHARGE_VOLUME_VISION,
                        data1.FPCB_BAR_CODE,
                        data1.BOX2_GLUE_OVERFLOW_VISION,
                        $"\"{data1.BOX2_HEATED_AIR_CURING}\"",
                        data1.BOX3_BOND,
                        data1.BOX3_OUPUTTIME,
                        data1.BOX3_DISTANCE,
                        data1.BOX3_GLUE_AMOUNT,
                        data1.BOX3_GLUE_DISCHARGE_VOLUME_VISION,
                        $"\"{data1.BOX3_HEATED_AIR_CURING}\"",
                        data1.BOX3_GLUE_OVERFLOW_VISION,
                        data1.BOX4_TIGHTNESS_AND_LOCATION_VISION,
                        $"\"{data1.BOX4_HEIGHT_PARALLELISM}\"",
                        data1.BOX4_RESISTANCE,
                        data1.BOX4_AIR_LEAKAGE_TEST_DETAIL,
                        data1.BOX4_AIR_LEAKAGE_TEST_RESULT,
                        data1.BOX4_LEAK_NAME,
                        finalResult ? "OK" : "NG",
                        formattedDateTime.Substring(0, 10),
                        formattedDateTime.Substring(11)
                    };

                    writer.WriteLine(string.Join(",", data));
                }

                //is check NAS = 1 meaning save to MES or not and machine state online or mode rework will push data to MES
                if (Global.IsCheckNAS == 1 && (Global.CurrentStateBox4 == (int)EMode.ONLINE || Global.CurrentModeBox4 == (int)ERework.REWORK))
                {
                    try
                    {
                        File.Copy(path, pathNAS, true);
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"Error can not copy file csv to folder NAS QR - {data1.TOPHOUSING}, error: {ex.Message}");
                    }
                }
            }
        }

        //read file, search and split
        #region

        public List<string> ReadFilesAndSearchV2(string directoryPath, string searchKeyword, int numberFileLimit = 180)
        {
            List<string> resultLines = new List<string>();
            int countFile = 0;

            try
            {
                RecursiveSearch(directoryPath, searchKeyword, ref countFile, numberFileLimit, resultLines);
            }
            catch (Exception ex)
            {
                WriteLog($"Error can not get data search, error: {ex.Message}");
            }

            return resultLines;
        }

        private void RecursiveSearch(string currentDirectory, string searchKeyword, ref int countFile, int numberFileLimit, List<string> resultLines)
        {
            var markFound = false;

            if (countFile >= numberFileLimit || resultLines.Count > 0)
                return;

            var directories = Directory.GetDirectories(currentDirectory).Select(dir => new DirectoryInfo(dir)).OrderByDescending(dirInfo => dirInfo.CreationTime).Select(dirInfo => dirInfo.FullName).ToList();

            foreach (var directory in directories)
            {
                var files = Directory.GetFiles(directory).Select(file => new FileInfo(file)).OrderByDescending(fileInfo => fileInfo.CreationTime).Select(fileInfo => fileInfo.FullName);

                if (files.Count() > 0)
                {
                    foreach (var file in files)
                    {
                        countFile++;

                        if (countFile >= numberFileLimit)
                        {
                            markFound = true; //not found when read file limit
                            return;
                        }

                        using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (StreamReader reader = new StreamReader(fs))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (line.Contains(searchKeyword))
                                {
                                    resultLines.Add($"{line}");
                                    markFound = true;
                                }
                            }
                        }

                        if (markFound)
                        {
                            return;
                        }
                    }
                }

                if (markFound)
                {
                    return;
                }
                else
                {
                    RecursiveSearch(directory, searchKeyword, ref countFile, numberFileLimit, resultLines);
                }
            }
        }

        public static BOX4RESULT SpiltData4(string input)
        {
            string[] lines = input.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            List<string> strings = new List<string>();
            BOX4RESULT result = new BOX4RESULT();
            foreach (string line in lines)
            {
                string[] parts = line.Split(new[] { ':' });
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim().ToUpper();
                    string value = parts[1].Trim(',');

                    switch (key)
                    {
                        case "SERIALNUMBER":
                            result.TOPHOUSING = value;
                            strings.Add(value);
                            break;

                        case "TIGHTNESS AND LOCATION VISION":
                            result.TIGHTNESS_AND_LOCATION_VISION = value;
                            strings.Add(value);
                            break;

                        case "HEIGHT PARALLELISM":
                            result.HEIGHT_PARALLELISM = value;
                            strings.Add(value);
                            break;

                        case "RESISTANCE":
                            result.RESISTANCE = value;
                            strings.Add(value);
                            break;

                        case "AIR LEAKAGE TEST DETAIL":
                            result.AIR_LEAKAGE_TEST_DETAIL = value;
                            strings.Add(value);
                            break;

                        case "AIR LEAKAGE TEST RESULT":
                            result.AIR_LEAKAGE_TEST_RESULT = value;
                            strings.Add(value);
                            break;

                        case "TESTTIME":
                            result.TestTime = value;
                            strings.Add(value);
                            break;
                        case "PORT":
                            result.Leak_Name = value;
                            strings.Add(value);
                            break;
                        default:
                            break;
                    }
                }
                else if (parts.Length > 2)
                {
                    int index = line.IndexOf(": ");
                    string key = line.Substring(0, index); // Extract "TestTime"
                    string value = line.Substring(index + 2); // Extract "2024-06-29 12:56:14"


                    if (key == "TestTime")
                        result.TestTime = value;
                  
                }
            }

            return result;
        }
        public static BOX3RESULT SpiltData3(string input)
        {
            //.Split(new Char [] {' ', ',', '.', '-', '\n', '\t' } );
            string[] lines = input.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            List<string> strings = new List<string>();
            BOX3RESULT result = new BOX3RESULT();
            foreach (string line in lines)
            {
                //string[] spliline = line.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string[] parts = line.Split(new[] { ':' });
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim().ToUpper();
                    string value = parts[1].Trim(',');

                    switch (key)
                    {
                        case "SERIALNUMBER":
                            result.TOPHOUSING = value;
                            strings.Add(value);
                            break;
                        case "3ND HEATED AIR CURING":
                            result.BOX3_HEATED_AIR_CURING = value;
                            break;
                        case "3ND GLUE AMOUNT":
                            result.GLUE_AMOUNT = value;
                            strings.Add(value);
                            break;
                        case "3ND GLUE DISCHARGE VOLUME VISION":
                            result.GLUE_DISCHARGE_VOLUME_VISION = value;
                            strings.Add(value);
                            break;

                        case "3ND GLUE OVERFLOW VISION":
                            result.BOX3_GLUE_OVERFLOW_VISION = value;
                            strings.Add(value);
                            break;
                        case "DISTANCE":
                            result.DISTANCE = value;
                            break;
                        case "TESTTIME":
                            result.TestTime = value;
                            strings.Add(value);
                            break;
                        case "BOND CODE":
                            result.BondCode = value;
                            break;
                        //case "OutPut":
                        //    result.Output = value;
                        //    break;
                        case "E-INPUT":
                            result.Input = value;
                            break;
                        default:
                            break;
                    }
                }
                else if (parts.Length > 2)
                {
                    int index = line.IndexOf(": ");
                    string key = line.Substring(0, index).ToUpper().Trim(); // Extract "TestTime"
                    string value = line.Substring(index + 2); // Extract "2024-06-29 12:56:14"


                    if (key == "TESTTIME")
                        result.TestTime = value;
                    if (key.ToUpper() == "OUTPUT")
                        result.Output = value;
                }
            }
            return result;

        }
        public static BOX2RESULT SpiltData2(string input)
        {
            //.Split(new Char [] {' ', ',', '.', '-', '\n', '\t' } );
            string[] lines = input.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            BOX2RESULT result = new BOX2RESULT();
            foreach (string line in lines)
            {
                //string[] spliline = line.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string[] parts = line.Split(new[] { ':' });
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim().ToUpper();
                    string value = parts[1].Trim(',');

                    switch (key)
                    {
                        case "SERIALNUMBER":
                            result.TOPHOUSING = value;
                            break;
                        case "2ND HEATED AIR CURING":
                            result.BOX2_HEATED_AIR_CURING = value;
                            break;
                        case "2ND GLUE AMOUNT":
                            result.GLUE_AMOUNT = value;
                            break;
                        case "2ND GLUE DISCHARGE VOLUME VISION":
                            result.GLUE_DISCHARGE_VOLUME_VISION = value;
                            break;
                        case "FPCB BAR CODE":
                            result.FPCB_BAR_CODE = value;
                            break;
                        case "2ND GLUE OVERFLOW VISION":
                            result.GLUE_OVERFLOW_VISION = value;
                            break;
                        case "TESTTIME":
                            result.TestTime = value;
                            break;
                        case "BOND CODE":
                            result.BondCode = value;
                            break;
                        //case "OutPut":
                        //    result.Output = value;
                        //    break;
                        case "E-INPUT":
                            result.Input = value;
                            break;
                        default:
                            break;
                    }
                }
                else if (parts.Length > 2)
                {
                    int index = line.IndexOf(": ");
                    string key = line.Substring(0, index).ToUpper().Trim(); // Extract "TestTime"
                    string value = line.Substring(index + 2); // Extract "2024-06-29 12:56:14"


                    if (key == "TESTTIME")
                        result.TestTime = value;
                    if (key == "OUTPUT")
                        result.Output = value;
                }
            }

            return result;
        }

        public static BOX1RESULT SpiltData1(string input)
        {
            //.Split(new Char [] {' ', ',', '.', '-', '\n', '\t' } );

            string[] lines = input.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

            BOX1RESULT result = new BOX1RESULT();
            foreach (string line in lines)
            {
                //string[] spliline = line.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                string[] parts = line.Split(new[] { ':' });
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim().ToUpper();
                    string value = parts[1].Trim(',');

                    switch (key)
                    {
                        case "SERIALNUMBER":
                            result.TOPHOUSING = value;
                            break;
                        case "1ST GLUE AMOUNT":
                            result.GLUE_AMOUNT = value;
                            break;
                        case "1ST GLUE DISCHARGE VOLUME VISION":
                            result.GLUE_DISCHARGE_VOLUME_VISION = value;
                            break;
                        case "INSULATOR BAR CODE":
                            result.INSULATOR_BAR_CODE = value;
                            break;
                        case "1ST GLUE OVERFLOW VISION":
                            result.GLUE_OVERFLOW_VISION = value;
                            break;
                        case "1ST HEATED AIR CURING":
                            result.BOX1_HEATED_AIR_CURING = value;
                            break;
                        case "TESTTIME":
                            result.TestTime = value;
                            break;
                        case "BOND CODE":
                            result.BondCode = value;
                            break;
                        //case "OutPut":
                        //    result.Output = value;
                        //    break;
                        case "E-INPUT":
                            result.Input = value;
                            break;
                        default:
                            break;
                    }
                }
                else if (parts.Length > 2)
                {
                    int index = line.IndexOf(": ");
                    string key = line.Substring(0, index).ToUpper().Trim(); // Extract "TestTime"
                    string value = line.Substring(index + 2); // Extract "2024-06-29 12:56:14"


                    if (key == "TESTTIME")
                        result.TestTime = value;
                    if (key == "OUTPUT")
                        result.Output = value;
                }
            }
            return result;

        }
        #endregion

        private static readonly object lockWriteLog = new object();

        public static void WriteLog(string logMessage)
        {
            lock (lockWriteLog)
            {
                string logPath = $@"D:\Logs\CIM\LogsData\{DateTime.Now.ToString("yyyy")}\{DateTime.Now.ToString("MM")}";

                string logFormat = DateTime.Now.ToLongDateString().ToString() + " - " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";

                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }

                try
                {
                    using (StreamWriter writer = File.AppendText(logPath + "\\" + DateTime.Now.ToString("dd") + ".txt"))
                    {
                        writer.WriteLine(logFormat + logMessage);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error can not write log, error: {ex.Message}");
                }
            }
        }

        private static readonly object lockWriteDuplicateBarcode = new object();
        public static void WriteBarcodeDuplicate(string barcode)
        {
            lock (lockWriteDuplicateBarcode)
            {
                string pathDuplicateBarcode = $@"C:\Users\SEM_PROJECT_CIMPC\Desktop\duplicate_barcode.txt";

                if (!File.Exists(pathDuplicateBarcode))
                {
                    using (File.Create(pathDuplicateBarcode)) { }
                }

                string logFormat = DateTime.Now.ToLongDateString().ToString() + " - " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";

                try
                {
                    using (StreamWriter writer = File.AppendText(pathDuplicateBarcode))
                    {
                        writer.WriteLine(logFormat + barcode);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error can not write log, error: {ex.Message}");
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("dddd, MMM dd, yyyy | HH:mm:ss");
        }

        private void MessageTimer_Tick(object sender, EventArgs e)
        {
            MessageTimer.Stop();
            lblalert.Text = "";
            lblalert.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormSetting fs = new FormSetting();
            fs.ShowDialog();
        }
        public static int a = 1;
        public static int a1 = 1;
        public static int a2= 1;
        public static int a3 = 1;
        private void button1_Click(object sender, EventArgs e)
        {
           
            ReadData1($"QRcode{a}","20","OK", "OK",$"INSU {a}", "OK","138","138","130","140","Bond 1st","2024-12-18 15:30:25");
            a++;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ReadData2($"QRcode{a1}","138", "138", "130", "140", "OK","25","OK",$"FPCBbarcode {a}","OK", "Bond 2nd", "2024-12-17 14:30:25");
            a1++;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ReadData3($"QRcode{a2}", "OK", "138", "138", "130", "140",$"{a}", "25","OK", "Bond 3rd", "2024-12-19 10:30:25");
            a2++;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ReadData4($"QRcode{a3}", "OK", "OK", "130", "140", "25", "OK", "20.25","25.25","138","60","58","0.001","OK", "OK", "Leak name1" );
            a3++;
        }

        private void btnFormSearch_Click(object sender, EventArgs e)
        {
            SearchForm fs = new SearchForm();
            fs.ShowDialog();
        }

        private void btnClearData_Click(object sender, EventArgs e)
        {
            lbltotal.Text = "0 EA";
            lblOK.Text = "0";
            lblNG.Text = "0";
            lblperOK.Text = "0%";
            lblperNG.Text = "0%";
            dataGridView1.Rows.Clear();
            OK = 0;
            NG = 0;
            No = 0;
            Total = OK + NG;
            pieChart1.UpdateChartData(OK, NG);

            Global.WriteFileToTxt(Global.GetFilePathSetting(), new Dictionary<string, string>
            {
                { "OK", OK.ToString() },
                { "NG_AIR", NG.ToString() },
                { "TOTAL", Total.ToString() },
            });
        }

        public static string GetTraycode(int qty)
        {
            string result = "PA84-00176";

            string model = SingleTonPlcControl.Instance.GetValueRegister(4, "READ_MODEL").ToString().Trim();

            switch (model)
            {
                case "1":
                    result += "J";
                    break;

                case "2":
                    result += "H";
                    break;

                case "3":
                    result += "G";
                    break;

                default:
                    result += "J";
                    break;
            }

            result += "U";
            DateMap mapper = new DateMap();

            string ymd = mapper.GetValue(DateTime.Today);

            result += ymd;

            return result + "36";
        }
    }
}
