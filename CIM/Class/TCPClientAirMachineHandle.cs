

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CIM
{
    public class TCPClientAirMachineHandle
    {
        private const int PORT = 23;
        private List<TcpClient> clients = new List<TcpClient>();
        private List<NetworkStream> streams = new List<NetworkStream>();
        private List<bool> connected = new List<bool>();
        private List<string> serverIPs = new List<string>();
        private readonly string[] ipAddresses = new string[]
        {
            "192.168.3.170","192.168.3.171", "192.168.3.172", "192.168.3.173", "192.168.3.174",
            "192.168.3.175", "192.168.3.176","192.168.3.177", "192.168.3.178","192.168.3.179"
        };

        public async Task ConnectAllAsync()
        {
            foreach (var ip in ipAddresses)
            {
                var client = new TcpClient();
                clients.Add(client);
                try
                {
                    await ConnectClientAsync(ip, PORT, clients.Count - 1);
                }
                catch (TimeoutException ex)
                {
                    MessageBox.Show($"Timeout connecting to {ip}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error connecting to {ip}: {ex.Message}");
                }
            }
        }

        private async Task ConnectClientAsync(string ipAddress, int port, int clientIndex)
        {
            int attempt = 0;
            while (attempt < 10)
            {
                try
                {
                    await clients[clientIndex].ConnectAsync(ipAddress, port);
                    if (clients[clientIndex].Connected)
                    {
                        streams.Add(clients[clientIndex].GetStream());
                        connected.Add(true);
                        await SendDataConnectAsync(clientIndex);
                       
                            await StartReceivingAsync(clientIndex);
                        
                      
                    }
                }
                catch (SocketException ex)
                {
                    // Handle socket exceptions
                }
                catch (Exception ex)
                {
                    // Handle other exceptions
                }
                attempt++;
                await Task.Delay(100); // Delay before retrying connection
            }
        }

        private async Task SendDataConnectAsync(int clientIndex)
        {
            try
            {
                string message = "1\r\n";
                byte[] data = Encoding.UTF8.GetBytes(message);
                if (connected[clientIndex])
                {
                    await streams[clientIndex].WriteAsync(data, 0, data.Length);
                }
            }
            catch (Exception ex)
            {
                // Handle send data exceptions
            }
        }

        private async Task StartReceivingAsync(int clientIndex)
        {
            try
            {
                byte[] buffer = new byte[1024];
                while (connected[clientIndex])
                {
                    int bytesRead = await streams[clientIndex].ReadAsync(buffer, 0, buffer.Length);
                    string message = Encoding.UTF8.GetString(buffer, 0, bytesRead);
                    await Task.Run(() => ExtractValueSccm(message, clientIndex));
                }
            }
            catch (Exception ex)
            {
                if (connected[clientIndex])
                {
                    connected[clientIndex] = false;
                    // Handle disconnection or receive data exceptions
                }
            }
        }
        private async Task StartReceivingAsync1(int clientIndex)
        {
            try
            {
                byte[] buffer = new byte[1024];
                while (connected[clientIndex])
                {
                    int bytesRead = await streams[clientIndex].ReadAsync(buffer, 0, buffer.Length);
                    string message = Encoding.UTF8.GetString(buffer, 0, bytesRead);
                    await Task.Run(() => ExtractValueSccm1(message, clientIndex));
                }
            }
            catch (Exception ex)
            {
                if (connected[clientIndex])
                {
                    connected[clientIndex] = false;
                    // Handle disconnection or receive data exceptions
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
                    //HandleReadAirTest(sccmValue, clientIndex);
                    Task.Run(() => WriteLog1(clientIndex, sccmValue));
                }
            }
            else if (input.Contains(" SL "))
            {
                Task.Run(() => WriteLog1(clientIndex, "1"));
                //HandleReadAirTest("SL", clientIndex);
            }
        }
        private void ExtractValueSccm1(string input, int clientIndex)
        {
            Task.Run(() => WriteLog(0,input));
            //    if (input.Contains("sccm"))
            //    {
            //        string patternSccm = @"(\d+\.\d+)\s*sccm";
            //        Match match = Regex.Match(input, patternSccm);
            //        if (match.Success)
            //        {
            //            string sccmValue = match.Groups[1].Value;
            //            HandleReadAirTest(sccmValue, clientIndex);
            //        }
            //    }
            //    else if (input.Contains(" SL "))
            //    {
            //        HandleReadAirTest("SL", clientIndex);
            //    }
        }
        public static void WriteLog(int clinetIndex,string logMessage)
        {
            string logFormat = string.Empty;
            string logPath = $"D:\\Logs\\CIM\\";
            logFormat = DateTime.Now.ToLongDateString().ToString() + " - " +
                    DateTime.Now.ToLongTimeString().ToString() + " ==> ";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
                ///NOTE

            }

            try
            {
                //create new log
                using (StreamWriter writer = File.AppendText(logPath + clinetIndex + ".txt"))
                {

                    writer.WriteLine(logFormat + logMessage);
                }
                //}


            }
            catch (Exception ex)
            {

            }
        }
        public static void WriteLog1(int clinetIndex, string logMessage)
        {
            string logFormat = string.Empty;
            string logPath = $"D:\\Logs\\CIM\\";
            logFormat = DateTime.Now.ToLongDateString().ToString() + " - " +
                    DateTime.Now.ToLongTimeString().ToString() + " ==> ";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
                ///NOTE

            }

            try
            {
                //create new log
                using (StreamWriter writer = File.AppendText(logPath + clinetIndex + ".txt"))
                {

                    writer.WriteLine(logFormat + logMessage);
                }
                //}


            }
            catch (Exception ex)
            {

            }
        }

        private void HandleReadAirTest(string sccmValue, int clientIndex)
        {
            //Task.Run(() => Form1.HandleReadAirTest(sccmValue, clientIndex));
        }

        private void Log(string msg)
        {
            MessageBox.Show(msg, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}



//using NPOI.SS.Formula.Functions;
//using System;
//using System.Collections.Generic;
//using System.Net;
//using System.Net.Sockets;
//using System.Text;
//using System.Text.RegularExpressions;
//using System.Threading;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using UnilityCommand.SupportInit;

//namespace CIM
//{
//    public class TCPClientAirMachineHandle
//    {
//        private const int PORT = 23;
//        public ResultAirMachine rsAirMachine = new ResultAirMachine();

//        public static TcpClient[] clients = new TcpClient[10];
//        private NetworkStream[] streams = new NetworkStream[10];
//        private bool[] connected = new bool[10];
//        private List<string> serverIPs = new List<string>();
//        public readonly string[] ipAddresses = new string[]
//      {
//            "192.168.3.170","192.168.3.171", "192.168.3.172", "192.168.3.173", "192.168.3.174",
//            "192.168.3.175", "192.168.3.176","192.168.3.177", "192.168.3.178","192.168.3.179"
//      };

//        public void connectALL()
//        {
//            for (int i = 0; i < ipAddresses.Length; i++)
//            {

//                string ip = ipAddresses[i];

//                clients[i] = new TcpClient();
//                try
//                {
//                    ConnectClients(ip, PORT, i);
//                }
//                catch (TimeoutException ex)  // Xử lý lỗi kết nối timeout ở đây
//                {
//                    MessageBox.Show("bbbbbb");
//                }
//                catch (Exception ex) // Xử lý các lỗi khác nếu có
//                {
//                    MessageBox.Show("aaa");
//                }

//            }
//        }
//        public void ConnectClients(string ipAddress, int port, int clientIndex)
//        {
//            Task.Run(() => ConnectClient(clientIndex, ipAddress, port));
//        }

//        private async void SendataConnect(int connect)
//        {
//            try
//            {
//                string message = "1\r\n";
//                byte[] data = Encoding.UTF8.GetBytes(message);
//                if (connected[connect])
//                {
//                    await streams[connect].WriteAsync(data, 0, data.Length);

//                }
//                else
//                {

//                }

//            }
//            catch (Exception ex)
//            {

//            }
//        }
//        private void ConnectClient(int clientIndex, string ipAddress, int port)
//        {
//            int attempt = 0;
//            while (attempt < 10)
//            {
//                try
//                {
//                    clients[clientIndex].Connect(ipAddress, port); // Synchronous connect
//                    if (clients[clientIndex].Connected)
//                    {
//                        streams[clientIndex] = clients[clientIndex].GetStream();
//                        connected[clientIndex] = true;
//                        SendataConnect(clientIndex); // sendata de ket noi toi LEAK Test
//                        StartReceiving(clientIndex); // bat dau nhan du lieu

//                    }
//                }
//                catch (SocketException ex)
//                {

//                    //throw; // Rethrow the exception to handle retry logic
//                }
//                catch (Exception ex)
//                {

//                }
//                attempt++;
//                Thread.Sleep(100);
//            }
//        }

//        private void StartReceiving(int clientIndex)
//        {

//            try
//            {
//                byte[] buffer = new byte[1024];
//                while (connected[clientIndex])
//                {


//                    int bytesRead = streams[clientIndex].Read(buffer, 0, buffer.Length);
//                    string message = Encoding.UTF8.GetString(buffer, 0, bytesRead);
//                    //Task.Run(() => ProcessReceivedData(clientIndex, message));

//                    Task.Run(() => ExtractValueSccm(message, clientIndex));
//                    //Task.Run(() => ProcessReceivedData(clientIndex, message));
//                    //Task.Run(() => ProcessReceivedData(clientIndex, message));
//                }
//            }
//            catch (Exception ex)
//            {
//                if (connected[clientIndex])
//                {

//                    connected[clientIndex] = false;
//                }
//            }

//        }
//        private void ExtractValueSccm(string input, int clientIndex)
//        {

//            if (input.Contains("sccm"))
//            {
//                string patternSccm = @"(\d+\.\d+)\s*sccm";

//                System.Text.RegularExpressions.Match match = Regex.Match(input, patternSccm);
//                string pattern = @"\b(A|AR|SA)\b";
//                System.Text.RegularExpressions.Match match1 = Regex.Match(input, pattern);
//                if (match.Success)
//                {
//                    rsAirMachine.sccm = match.Groups[1].Value;

//                    Task.Run(() => Form1.HandleReadAirTest(rsAirMachine.sccm, clientIndex));


//                }



//            }
//            else if (input.Contains(" SL "))
//            {
//                rsAirMachine = new ResultAirMachine
//                {
//                    result = "SL",
//                    sccm = ""
//                };

//                Task.Run(() => Form1.HandleReadAirTest(rsAirMachine.sccm, clientIndex));


//            }



//        }


//        //private void ProcessReceivedData(int index, string dataReceived)
//        //{

//        //    if (dataReceived.Contains("sccm"))
//        //    {
//        //        string pattern = @"\b(A|AR|SA)\b";
//        //        Match match = Regex.Match(dataReceived, pattern);

//        //        string patternSccm = @"(\d+\.\d+)\s*sccm";
//        //        Match matchSccm = Regex.Match(dataReceived, patternSccm);


//        //        rsAirMachine.result = match.Success ? match.Value : "";
//        //        rsAirMachine.sccm = matchSccm.Success ? matchSccm.Groups[1].Value : "";



//        //    }
//        //    else if (dataReceived.Contains("SL"))
//        //    {
//        //         rsAirMachine = new ResultAirMachine
//        //        {
//        //            result = "SL",
//        //            sccm = ""
//        //        };


//        //    }





//        private void Log(string msg)
//        {
//            MessageBox.Show(msg, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//        }
//    }


//}
