using System;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConnectAirMachine1
{
    public class TCPClientAirMachineHandle
    {
        private const int PORT = 23;
        private TcpClient[] clients = new TcpClient[10];
        private NetworkStream[] streams = new NetworkStream[10];
        private CancellationTokenSource[] cancellationTokens = new CancellationTokenSource[10];

        public async Task<ResultAirMachine> ConnectTCP(int index)
        {
            try
            {
                var tcs = new TaskCompletionSource<ResultAirMachine>();
                clients[index] = new TcpClient();
                clients[index].Connect(GetIPByIndex(index), PORT);
                streams[index] = clients[index].GetStream();
                cancellationTokens[index] = new CancellationTokenSource();

                Log($"Connect_TCP - Success: Connect to server {index} successfully!");

                Thread clientThread = new Thread(() => HandleReceiveData(index, tcs, cancellationTokens[index].Token));
                clientThread.IsBackground = true;
                clientThread.Start();

                SendSignal(index, "1");
                SendSignal(index, "KEY START");

                return await tcs.Task;
            }
            catch (Exception ex)
            {
                Log($"Connect_TCP - Failed: Connect to server {index} failed! Exception: {ex.Message}");
                return null;
            }
        }

        private void HandleReceiveData(int index, TaskCompletionSource<ResultAirMachine> tcs, CancellationToken token)
        {
            TcpClient client = clients[index];
            NetworkStream stream = streams[index];

            try
            {
                byte[] buffer = new byte[1024];
                int bytesRead;

                while (!token.IsCancellationRequested)
                {
                    bytesRead = stream.Read(buffer, 0, buffer.Length);
                    
                    if (bytesRead > 0)
                    {
                        string dataReceived = Encoding.UTF8.GetString(buffer, 0, bytesRead);

                        if (dataReceived.Contains("sccm"))
                        {
                            //pattern value OK or other
                            string pattern = @"\b(A|AR|SA)\b";
                            Match match = Regex.Match(dataReceived, pattern);

                            //pattern value sccm
                            string patternSccm = @"(\d+\.\d+)\s*sccm";
                            Match matchSccm = Regex.Match(dataReceived, patternSccm);

                            ResultAirMachine rsAirMachine = new ResultAirMachine
                            {
                                result = match.Success ? match.Value : "",
                                sccm = matchSccm.Success ? matchSccm.Groups[1].Value : ""
                            };

                            break;
                        }

                        if (dataReceived.Contains("SL"))
                        {
                            ResultAirMachine rsAirMachine = new ResultAirMachine
                            {
                                result = "SL",
                                sccm = ""
                            };

                            tcs.TrySetResult(rsAirMachine);
                            client.Close();
                            break;
                        }
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                Log($"NetworkStream for client {index} has been disposed.");
            }
            catch (Exception ex)
            {
                Log($"Error in communication with client {index}: {ex.Message}");
            }
            finally
            {
                stream.Dispose();
                client.Close();
                cancellationTokens[index].Cancel();
            }
        }

        private void SendSignal(int index, string message)
        {
            TcpClient client = clients[index];
            NetworkStream stream = streams[index];

            if (client == null || !client.Connected)
            {
                Log("Not connected to server, send message error");
                return;
            }

            try
            {
                byte[] data = Encoding.UTF8.GetBytes(message + "\r\n");
                stream.Write(data, 0, data.Length);

                if (message == "KEY START")
                {
                    Log("Send KEY START successfully");
                }
            }
            catch (Exception ex)
            {
                CloseTCPIP(index);
                Log($"Error send message: {ex.Message}");
            }
        }

        private string GetIPByIndex(int index)
        {
            var rs = "";

            switch (index)
            {
                case 0:
                    rs = "192.168.3.170";
                    break;
                case 1:
                    rs = "192.168.3.171";
                    break;
                case 2:
                    rs = "192.168.3.172";
                    break;
                case 3:
                    rs = "192.168.3.173";
                    break;
                case 4:
                    rs = "192.168.3.174";
                    break;
                case 5:
                    rs = "192.168.3.175";
                    break;
                case 6:
                    rs = "192.168.3.176";
                    break;
                case 7:
                    rs = "192.168.3.177";
                    break;
                case 8:
                    rs = "192.168.3.178";
                    break;
                case 9:
                    rs = "192.168.3.179";
                    break;
            }

            return rs;
        }

        public void Log(string msg)
        {
            MessageBox.Show($"{msg}", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public void CloseTCPIP(int index)
        {
            clients[index]?.Close();
        }
    }

    public class ResultAirMachine
    {
        public string result { get; set; } = string.Empty;
        public string sccm { get; set; } = string.Empty;
    }
}
