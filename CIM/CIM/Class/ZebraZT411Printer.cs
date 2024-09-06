using System;
using Zebra.Sdk.Comm;
using Zebra.Sdk.Printer;


public class ZebraZT411Printer
{
    private Connection _connection;
    private ZebraPrinter _printer;

    public bool Connect(string ipAddress)
    {
        _connection = new TcpConnection(ipAddress, TcpConnection.DEFAULT_ZPL_TCP_PORT);
        try
        {
            _connection.Open();
            _printer = ZebraPrinterFactory.GetInstance(_connection);
            return true;
        }
        catch (ConnectionException ex)
        {
            Console.WriteLine("Could not connect to printer: " + ex.Message);
            return false;
        }
    }

    public void Disconnect()
    {
        try
        {
            _connection.Close();
        }
        catch (ConnectionException ex)
        {
            Console.WriteLine("Error disconnecting from printer: " + ex.Message);
        }
    }


    public void PrintLabel(string labelFormat)
    {
        try
        {
            _printer.SendCommand(labelFormat);
        }
        catch (ConnectionException ex)
        {
            Console.WriteLine("Error printing label: " + ex.Message);
        }
    }
}




//using System;
//using System.Text;
//using System.Net.Sockets;
//using MathNet.Numerics.LinearAlgebra.Factorization;

//namespace CIM
//{
//    public class ZebraZT411Printer
//    {
//        private string _ipAddress;
//        private int _port;

//        public ZebraZT411Printer(string ipAddress, int port)
//        {
//            _ipAddress = ipAddress;
//            _port = port;
//        }

//        public void PrintQRCode(string qrCodeData)
//        {
//            string zplCommand = "^XA";
//            zplCommand += "^LH15,20"; // Đặt label home để căn chỉnh giữa tem (30mm/2 = 15mm, 40mm/2 = 20mm)
//            //zplCommand += "^FO0,0^GB425,305,3^FS"; // Tạo hình chữ nhật có kích thước 30mm x 40mm (240dots x 320dots)
//            zplCommand += "^FO5,5"; // Đặt field origin để căn chỉnh lề trong hình chữ nhật
//            zplCommand += "^BQN,2,40"; // Đặt loại mã vạch QR (2: Auto, 5: Mã hình vuông)

//            zplCommand += "^FDLA," + qrCodeData + "^FS"; // Đặt dữ liệu mã QR
//            zplCommand += "^XZ"; // Kết thúc lệnh ZPL

//            //            ^BQ: Đây là lệnh để tạo mã QR.Cú pháp cơ bản như sau:

//            //^BQa,b,c,d,e
//            //a: Lựa chọn số hiệu của barcode.
//            //N cho mã QR(QR Code).
//            //b: Lựa chọn vị trí của lô barcode(trái, phải, hoặc ở giữa).
//            //N cho căn giữa.
//            //c: Độ rộng của mỗi vạch thanh(1 - 10).

//            byte[] command = Encoding.UTF8.GetBytes(zplCommand);

//            using (Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp))
//            {
//                socket.Connect(_ipAddress, _port);
//                socket.Send(command);
//                socket.Close();
//            }
//        }
//        //public void PrintQRCode(string qrCodeData)
//        //{
//        //    string zplCommand = "^XA"; 
//        //    zplCommand += "^LH15,20"; // Đặt label home để căn chỉnh giữa tem (30mm/2 = 15mm, 40mm/2 = 20mm)
//        //    zplCommand += "^FO0,0^GB425,305,3^FS"; // Tạo hình chữ nhật có kích thước 30mm x 40mm (240dots x 320dots)
//        //    zplCommand += "FO50,50^BQN,2,5";//"^FO40,60^BQN,2,5"; // Đặt field origin và loại mã vạch QR (FO50,50 có thể được điều chỉnh để căn chỉnh tốt hơn)
//        //    zplCommand += "^FDLA," + qrCodeData + "^FS"; // Đặt dữ liệu mã QR
//        //    zplCommand += "^XZ"; // Kết thúc lệnh ZPL

//        //    byte[] command = Encoding.UTF8.GetBytes(zplCommand);

//        //    using (Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp))
//        //    {
//        //        socket.Connect(_ipAddress, _port);
//        //        socket.Send(command);
//        //        socket.Close();
//        //    }
//        //}
//    }
//}