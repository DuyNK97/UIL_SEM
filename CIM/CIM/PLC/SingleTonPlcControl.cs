
using ActUtlType64Lib;
using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UnilityCommand;
using UnilityCommand.Plc;
using UnilityCommand.Plc.Mitsubishi;

using static System.Net.Mime.MediaTypeNames;

namespace UILAlignProject.PLC
{
    public class SingleTonPlcControl : Bindable//, IPlcControl, IRobotJogControl
    {
        public event EventHandler HaveNewSampleEvent;
        int interval = 70;
        bool isTest = false;
        int MxPort1 = 5;
        int MxPort2 = 21;
        int MxPort3 = 36;
        int MxPort4 = 56;
       // int MxPort4 = 1;
        int plcindex = 2;
        private ActUtlType64Class PlcWrite1 { get; set; }
        private ActUtlType64Class PlcRead1 { get; set; }

        private ActUtlType64Class PlcWrite2 { get; set; }
        private ActUtlType64Class PlcRead2 { get; set; }

        private ActUtlType64Class PlcWrite3 { get; set; }
        private ActUtlType64Class PlcRead3 { get; set; }

        private ActUtlType64Class PlcWrite4 { get; set; }
        private ActUtlType64Class PlcRead4 { get; set; }

        public RegisterPc2Plc RegisterWrite { get; set; }
        public RegisterPlc2PC RegisterRead { get; set; }

        public RegisterPc2Plc RegisterWrite2 { get; set; }
        public RegisterPlc2PC RegisterRead2 { get; set; }

        public RegisterPc2Plc RegisterWrite3 { get; set; }
        public RegisterPlc2PC RegisterRead3 { get; set; }

        public RegisterPc2Plc RegisterWrite4 { get; set; }
        public RegisterPlc2PC RegisterRead4 { get; set; }


        private Thread threadReadPlc1;
        private Thread threadReadPlc2;
        private Thread threadReadPlc3;
        private Thread threadReadPlc4;

        private static SingleTonPlcControl instance;

        static object _lock = new object();
        public static SingleTonPlcControl Instance
        {
            get
            {

                lock (_lock)
                {
                    if (instance == null)
                    {
                        instance = new SingleTonPlcControl();
                    }
                }

                return instance;
            }

        }



        private SingleTonPlcControl()
        {

        }

        ~SingleTonPlcControl()
        {
            try
            {
                Disconnect();
            }
            catch (Exception)
            {
            }
        }

        public bool Connect(int numberPort)
        {
            MxPort1 = numberPort;
            return Connect1();
        }

        public bool Connect1()
        {

            PlcRead1 = new ActUtlType64Class();
            PlcWrite1 = new ActUtlType64Class();
            PlcWrite1.ActLogicalStationNumber = MxPort1;
            PlcRead1.ActLogicalStationNumber = MxPort1;



            int connecting = PlcWrite1.Open();
            int connecting1 = PlcRead1.Open();

            if (connecting == 0 && connecting1 == 0)
            {
                IsConnect1 = true;
                RegisterWrite = new RegisterPc2Plc(this);
                RegisterRead = new RegisterPlc2PC(this);

                threadReadPlc1 = new Thread(() =>
                {
                    ReadDataFromPlc1();
                });
                threadReadPlc1.IsBackground = true;
                threadReadPlc1.Start();
                Task.Delay(4).Wait();

                return true;
            }
            else
            {
                IsConnect1 = false;
                return false;
            }
        }

        public bool Connect2()
        {

            PlcRead2 = new ActUtlType64Class();
            PlcWrite2 = new ActUtlType64Class();
            PlcWrite2.ActLogicalStationNumber = MxPort2;
            PlcRead2.ActLogicalStationNumber = MxPort2;


            int connecting = PlcWrite2.Open();
            int connecting1 = PlcRead2.Open();

            if (connecting == 0 && connecting1 == 0)
            {
                IsConnect2 = true;
                //RegisterWrite2 = new RegisterPc2Plc(this);
                //RegisterRead2 = new RegisterPlc2PC(this);

                threadReadPlc2 = new Thread(() =>
                {
                    ReadDataFromPlc2();
                });
                threadReadPlc2.IsBackground = true;
                threadReadPlc2.Start();
                Task.Delay(4).Wait();

                return true;
            }
            else
            {
                IsConnect2 = false;
                return false;
            }
        }

        public bool Connect3()
        {

            PlcRead3 = new ActUtlType64Class();
            PlcWrite3 = new ActUtlType64Class();
            PlcWrite3.ActLogicalStationNumber = MxPort3;
            PlcRead3.ActLogicalStationNumber = MxPort3;


            int connecting = PlcWrite3.Open();
            int connecting1 = PlcRead3.Open();

            if (connecting == 0 && connecting1 == 0)
            {
                IsConnect3 = true;
                //RegisterWrite3 = new RegisterPc2Plc(this);
                //RegisterRead3 = new RegisterPlc2PC(this);

                threadReadPlc3 = new Thread(() =>
                {
                    ReadDataFromPlc3();
                });
                threadReadPlc3.IsBackground = true;
                threadReadPlc3.Start();
                Task.Delay(4).Wait();

                return true;
            }
            else
            {
                IsConnect3 = false;
                return false;
            }
        }
        public bool Connect4()
        {

            PlcRead4 = new ActUtlType64Class();
            PlcWrite4 = new ActUtlType64Class();
            PlcWrite4.ActLogicalStationNumber = MxPort4;
            PlcRead4.ActLogicalStationNumber = MxPort4;


            int connecting = PlcWrite4.Open();
            int connecting1 = PlcRead4.Open();

            if (connecting == 0 && connecting1 == 0)
            {
                IsConnect4 = true;
                //RegisterWrite4 = new RegisterPc2Plc(this);
                //RegisterRead4 = new RegisterPlc2PC(this);

                threadReadPlc4 = new Thread(() =>
                {
                    ReadDataFromPlc4();
                });
                threadReadPlc4.IsBackground = true;
                threadReadPlc4.Start();
                Task.Delay(4).Wait();

                return true;
            }
            else
            {
                IsConnect4 = false;
                return false;
            }
        }

        private void PlcRead_OnDeviceStatus(string szDevice, int lData, int lReturnCode)
        {
            MessageBox.Show($"{szDevice} ,  {lData}, {lReturnCode}");
        }

        private bool _isConnect1;

        public bool IsConnect1
        {
            get { return _isConnect1; }
            set
            {
                _isConnect1 = value;
                Notify();
            }
        }

        private bool _isConnect2;

        public bool IsConnect2
        {
            get { return _isConnect2; }
            set
            {
                _isConnect2 = value;
                Notify();
            }
        }

        private bool _isConnect3;
        public bool IsConnect3
        {
            get { return _isConnect3; }
            set
            {
                _isConnect3 = value;
                Notify();
            }
        }
        private bool _isConnect4;
        public bool IsConnect4
        {
            get { return _isConnect4; }
            set
            {
                _isConnect4 = value;
                Notify();
            }
        }

        private int _readCylcleTime = 100;

        public int ReadCylcleTime
        {
            get { return _readCylcleTime; }
            set { _readCylcleTime = value; }
        }


        private LocationRobot _locationRobot2 = new LocationRobot(0, 0, 0, 0);

        public LocationRobot LocationRobot2
        {
            get { return _locationRobot2; }
            set { _locationRobot2 = value; Notify(); }
        }


        public DateTime GetDateTimePlc
        {
            get
            {
                DateTime dateTime = DateTime.Now;
                Tuple<double, double, double, double, double, double> tuple = null;
                var year = GetValueRegister(plcindex, "TimeYear");
                var month = GetValueRegister(plcindex, "TimeMonth");
                var day = GetValueRegister(plcindex, "TimeDay");
                var hour = GetValueRegister(plcindex, "TimeHour");
                var minute = GetValueRegister(plcindex, "TimeMinute");
                var second = GetValueRegister(plcindex, "TimeSecond");

                int y;
                int m;
                int d;
                int h;
                int mm;
                int s;
                if (Convert.ToInt32(year) >= 2020)
                {
                    y = Convert.ToInt32(year);

                }
                else
                {
                    y = dateTime.Year;
                }
                if (Convert.ToDouble(month) > 0)
                {
                    m = Convert.ToInt32(month);

                }
                else
                {
                    m = dateTime.Month;
                }
                if (Convert.ToInt32(day) > 0)
                {
                    d = Convert.ToInt32(day);

                }
                else
                {
                    d = dateTime.Day;
                }
                h = Convert.ToInt32(hour);
                mm = Convert.ToInt32(minute);
                s = Convert.ToInt32(second);


                return new DateTime(y, m, d, h, mm, s, DateTimeKind.Local);
            }
        }


        public bool Disconnect()
        {

            PlcWrite1?.Close();
            PlcRead1?.Close();
            return true;
        }

        public async void ReadDataFromPlc1()
        {
            if (!_isConnect1) { return; }
            Task.Delay(5000).Wait();
            while (true)
            {
                await Task.Delay(interval);
                try
                {
                    ReadAlway(PlcRead1, RegisterRead.RegisterBit1, RegisterRead.RegisterWord1, RegisterRead.RegisterDWord1, RegisterRead.RegisterString1, RegisterRead.RegisterBitInWord1, RegisterRead.RegisterFloat1);
                }
                catch (Exception)
                {

                }

            }
        }

        public async void ReadDataFromPlc2()
        {

            if (!_isConnect2) { return; }
            Task.Delay(5000).Wait();
            while (true)
            {


                await Task.Delay(interval);
                try
                {
                    ReadAlway(PlcRead2, RegisterRead.RegisterBit2, RegisterRead.RegisterWord2, RegisterRead.RegisterDWord2, RegisterRead.RegisterString2, RegisterRead.RegisterBitInWord2, RegisterRead.RegisterFloat2);
                }
                catch (Exception)
                {

                }

            }
        }

        public async void ReadDataFromPlc3()
        {

            if (!_isConnect3) { return; }
            Task.Delay(5000).Wait();
            while (true)
            {


                await Task.Delay(interval);
                try
                {
                    //*
                    ReadAlway(PlcRead3, RegisterRead.RegisterBit3, RegisterRead.RegisterWord3, RegisterRead.RegisterDWord3, RegisterRead.RegisterString3, RegisterRead.RegisterBitInWord3, RegisterRead.RegisterFloat3);
                }
                catch (Exception)
                {

                }

            }
        }
        public async void ReadDataFromPlc4()
        {

            if (!_isConnect4) { return; }
            Task.Delay(5000).Wait();
            while (true)
            {


                await Task.Delay(interval);
                try
                {
                    //*
                    ReadAlway(PlcRead4, RegisterRead.RegisterBit4, RegisterRead.RegisterWord4, RegisterRead.RegisterDWord4, RegisterRead.RegisterString4, RegisterRead.RegisterBitInWord4, RegisterRead.RegisterFloat4);
                }
                catch (Exception)
                {

                }

            }
        }

        private void ReadAlway(ActUtlType64Class plc, List<PLCIOCollection> rbit, List<PLCIOCollection> rword, List<PLCIOCollection> rdword, List<PLCIOCollection> rString, List<PLCIOCollection> rbitinword, List<PLCIOCollection> rfloat)
        {
            foreach (PLCIOCollection item in rbit)
            {
                int index = item.First().RegisterPLC;
                if (item.Count == 1)
                {
                    short result;
                    plc.GetDevice2($"M{index}", out result);
                    item[0].CurrentValue = result == 1 ? true : false;
                }
                else
                {
                    var count = item.Count;
                    var c = count % 4 == 0 ? count / 4 : count / 4 + 1;
                    if (c <= 4)
                    {
                        short res;
                        var ss = plc.ReadDeviceRandom2($"k{c}M{index}", 1, out res);
                        var arr = BitConverter.GetBytes(res);
                        BitArray b = new BitArray(arr);
                        bool[] bits = b.Cast<bool>().Select(bit => bit ? true : false).ToArray();
                        if (bits.Count() > 0)
                        {
                            Task.Run(() =>
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    item[i].CurrentValue = bits[i];
                                }
                            });
                        }
                    }
                    else
                    {
                        int res;
                        plc.ReadDeviceRandom($"k{c}M{index}", 1, out res);
                        var arr = BitConverter.GetBytes(res);
                        BitArray b = new BitArray(arr);
                        bool[] bits = b.Cast<bool>().Select(bit => bit ? true : false).ToArray();
                        if (bits.Count() > 0)
                        {
                            Task.Run(() =>
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    item[i].CurrentValue = bits[i];
                                }
                            });
                        }
                    }


                }

            }

            foreach (PLCIOCollection item in rword)
            {
                int index = item.First().RegisterPLC;

                var count = item.Count;
                int maxOfOneRead = 10;
                if (count <= maxOfOneRead)
                {
                    short[] res = new short[count];
                    plc.ReadDeviceBlock2($"D{index}", count, out res[0]);
                    Task.Run(() =>
                    {
                        for (int i = 0; i < count; i++)
                        {
                            item[i].CurrentValue = res[i];
                        }
                    });

                }
                else
                {
                    int less = count;
                    int indexLoop = 0;
                    do
                    {
                        ReadWord(indexLoop * maxOfOneRead, (indexLoop + 1) * maxOfOneRead, item);
                        less -= maxOfOneRead;
                        indexLoop++;
                    }
                    while (less - maxOfOneRead > maxOfOneRead);
                    ReadWord(indexLoop * maxOfOneRead, count, item);

                }
            }

            foreach (PLCIOCollection item in rdword)
            {
                int index = item.First().RegisterPLC;

                var count = item.Count;
                int maxOfOneRead = 10;
                if (count <= maxOfOneRead)
                {
                    short[] res = new short[count * 2];
                    var a = plc.ReadDeviceBlock2($"D{index}", count * 2, out res[0]);

                    for (int i = 0; i < count; i++)
                    {
                        var r = PLCExtension.ConvertShortArr2Int(res[i * 2 + 1], res[i * 2]);
                        item[i].CurrentValue = r;
                    }

                }
                else
                {

                    int less = count;
                    int indexLoop = 0;
                    do
                    {
                        ReadDWord(indexLoop * maxOfOneRead, (indexLoop + 1) * maxOfOneRead, item);
                        less -= maxOfOneRead;
                        indexLoop++;
                    }
                    while (less - maxOfOneRead > maxOfOneRead);
                    ReadDWord(indexLoop * maxOfOneRead, count, item);
                }
            }

            foreach (PLCIOCollection item in rfloat)
            {
                int index = item.First().RegisterPLC;

                var count = item.Count;
                int maxOfOneRead = 10;
                if (count <= maxOfOneRead)
                {
                    short[] res = new short[count * 2];
                    var a = plc.ReadDeviceBlock2($"D{index}", count * 2, out res[0]);

                    for (int i = 0; i < count; i++)
                    {

                        int[] v = new int[] { res[i * 2], res[i * 2 + 1] };

                        item[i].CurrentValue = v.ConvertWordToFloat();
                    }

                }
                else
                {

                    int less = count;
                    int indexLoop = 0;
                    do
                    {
                        ReadDWord(indexLoop * maxOfOneRead, (indexLoop + 1) * maxOfOneRead, item);
                        less -= maxOfOneRead;
                        indexLoop++;
                    }
                    while (less - maxOfOneRead > maxOfOneRead);
                    ReadDWord(indexLoop * maxOfOneRead, count, item);
                }
            }

            foreach (PLCIOCollection item in rString)
            {
                int index = item.First().RegisterPLC;

                var count = item.First().Size;


                short[] res = new short[count];
                plc.ReadDeviceBlock2($"D{index}", count, out res[0]);
                Task.Run(() =>
                {
                    item.First().CurrentValue = PLCExtension.ConvertShortArrToString(res);

                });
            }

            foreach (PLCIOCollection item in rbitinword)
            {
                int index = item.First().RegisterPLC;
                var count = 1;
                short[] res = new short[count];
                plc.ReadDeviceBlock2($"D{index}", count, out res[0]);
                var bitArr = PLCExtension.ConvertShortToBitArray(res[0]);
                foreach (var i in item)
                {
                    Task.Run(() =>
                    {
                        i.CurrentValue = (bool)bitArr[i.Size];
                    });
                }

            }
        }



        private void ReadWord(int from, int to, PLCIOCollection item)
        {
            int index = item[from].RegisterPLC;
            int count = to - from;
            if (count > 0)
            {
                short[] res = new short[count];
                PlcRead1?.ReadDeviceBlock2($"D{index}", count, out res[0]);
                Task.Run(() =>
                {
                    for (int i = 0; i < count; i++)
                    {
                        item[i + from].CurrentValue = res[i];
                    }
                });
            }

        }


        private void ReadFloat(int from, int to, PLCIOCollection item)
        {
            int index = item[from].RegisterPLC;
            int count = to - from;
            if (count > 0)
            {
                int[] res = new int[count];
                PlcRead1?.ReadDeviceBlock($"D{index}", count, out res[0]);
                Task.Run(() =>
                {
                    for (int i = 0; i < count; i++)
                    {
                        item[i + from].CurrentValue = res[i].ConvertDWordIntToFloat();
                    }
                });
            }
        }

        private void ReadDWord(int from, int to, PLCIOCollection item)
        {
            int index = item[from].RegisterPLC;
            int count = to - from;
            if (count > 0)
            {
                int[] res = new int[count];
                PlcRead1?.ReadDeviceBlock($"D{index}", count, out res[0]);
                Task.Run(() =>
                {
                    for (int i = 0; i < count; i++)
                    {
                        item[i + from].CurrentValue = res[i];
                    }
                });
            }
        }
        public bool ReadDWord(int index, string register, int leng, out int[] data)
        {
            data = new int[leng];
            var res = new short[leng * 2];
            ActUtlType64Class plc = PlcRead1;
            switch (index)
            {

                case 2:
                    plc = PlcRead2;

                    break;
                case 3:
                    plc = PlcRead3;
                    break;
                case 4:
                    plc = PlcRead4;
                    break;
                default:

                    break;
            }
            if (plc?.ReadDeviceBlock2(register, leng * 2, out res[0]) == 0)
            {
                for (int i = 0; i < leng; i++)
                {
                    var r = PLCExtension.ConvertShortArr2Int(res[i * 2 + 1], res[i * 2]);
                    data[i] = r;
                }
                return true;
            }
            return false;


        }

        public void PrepareBeforeRead()
        {

        }




        public bool ReadOneBit(string bit)
        {
            short val = 0;
            PlcRead1?.GetDevice2(bit, out val);
            return val == 1 ? true : false;
        }
        public bool ReadOneBit(int indexPlc, string bit)
        {
            short val = 0;
            ActUtlType64Class plc = PlcRead1;
            switch (indexPlc)
            {
                case 2:
                    plc = PlcRead2;
                    break;
                case 3:
                    plc = PlcRead3;
                    break;
                case 4:
                    plc = PlcRead4;
                    break;
                default:
                    break;
            }
            plc?.GetDevice2(bit, out val);
            return val == 1 ? true : false;
        }

        public bool WriteOneBit(string bit, bool value, int indexPlc)
        {
            ActUtlType64Class plc = PlcWrite1;
            switch (indexPlc)
            {
                case 2:
                    plc = PlcWrite2;
                    break;
                case 3:
                    plc = PlcWrite3;
                    break;
                case 4:
                    plc = PlcWrite4;
                    break;
                default:
                    plc = PlcWrite1;
                    break;
            }
            short val = 0;
            if (value)
            {
                val = 1;
            }
            else
            {
                val = 0;
            }

            var result = plc?.SetDevice2(bit, val);
            if (result == 1 || result == 0)
            {
                return true;
            }

            return false;
        }

        public bool WriteDWord(string register, int leng, int indexPlc, ref int[] data)
        {
            ActUtlType64Class plc = PlcWrite1;
            if (indexPlc == 2)
            {
                plc = PlcWrite2;
            }
            short[] shorts = new short[leng * 2];
            for (int i = 0; i < leng; i++)
            {
                var arr = PLCExtension.ConvertInt2ShortArrForPLC(data[i]);
                shorts[i * 2] = arr[1];
                shorts[(i * 2) + 1] = arr[0];
            }

            if (plc?.WriteDeviceBlock2(register, leng * 2, ref shorts[0]) == 0)
            {
                return true;
            }
            return false;
        }

        public bool WriteFloat(string register, int leng, int indexPlc, ref float[] data)
        {
            ActUtlType64Class plc = PlcWrite1;
            switch (indexPlc)
            {
                case 2:
                    plc = PlcWrite2;
                    break;
                case 3:
                    plc = PlcWrite3;
                    break;
                case 4:
                    plc = PlcWrite4;
                    break;
                default:
                    plc = PlcWrite1;
                    break;
            }
            
            short[] shorts = new short[leng * 2];
            for (int i = 0; i < leng; i++)
            {
                var arr = PLCExtension.ConvertInt2ShortArrForPLC(data[i].ConvertFloatToDWordInt());
                shorts[i * 2] = arr[1];
                shorts[(i * 2) + 1] = arr[0];
            }

            if (plc?.WriteDeviceBlock2(register, leng * 2, ref shorts[0]) == 0)
            {
                return true;
            }
            return false;
        }

        public bool ReadDWord(string register, int leng, out int[] data)
        {
            data = new int[leng];
            var res = new short[leng * 2];
            if (PlcRead1?.ReadDeviceBlock2(register, leng * 2, out res[0]) == 0)
            {
                for (int i = 0; i < leng; i++)
                {
                    var r = PLCExtension.ConvertShortArr2Int(res[i * 2 + 1], res[i * 2]);
                    data[i] = r;
                }
                return true;
            }
            return false;


        }

        public bool WriteString(string currvalue,int indexPLC,string regrister)
        {
            short[] res4 = PLCExtension.ConvertStringToShortArr(((string)currvalue));
            return  WriteWord($"ZR{regrister}", res4.Length , indexPLC, ref res4);
        }
        public bool WriteWord(string register, int leng, int indexPlc, ref short[] data)
        {
            ActUtlType64Class plc = PlcWrite1;

            switch (indexPlc)
            {
                case 2:
                    plc = PlcWrite2;
                    break;

                case 3:
                    plc = PlcWrite3;
                    break;

                case 4:
                    plc = PlcWrite4;
                    break;
            }

            var result = plc?.WriteDeviceBlock2(register, leng, ref data[0]);

            if (result == 0 || result == 0)
            {
                return true;
            }

            return false;
        }

        public bool ReadFloat(int index,string register, int leng, out float[] data)
        {
            data = new float[leng];
            var res = new short[leng * 2];
            ActUtlType64Class plc = PlcRead1;
            switch (index)
            {

                case 2:
                    plc = PlcRead2;

                    break;
                case 3:
                    plc = PlcRead3;
                    break;
                case 4:
                    plc = PlcRead4;
                    break;
                default:

                    break;
            }

            if (plc?.ReadDeviceBlock2(register, leng * 2, out res[0]) == 0)
            {
                for (int i = 0; i < leng; i++)
                {
                    var r = PLCExtension.ConvertShortArr2Int(res[i * 2 + 1], res[i * 2]).ConvertDWordIntToFloat();
                    data[i] = r;
                }
                return true;
            }
            return false;
        }

        public bool ReadWord(int index , string register, int leng, out short[] data)
        {
            data = new short[leng];
            ActUtlType64Class plc = PlcRead1;
            switch (index) 
            { 
                
                case 2:
                    plc = PlcRead2;
                   
                    break;
                case 3:
                    plc = PlcRead3;
                    break;
                case 4:
                    plc = PlcRead4;
                    break;
                default:
                   
                    break ;
            }
            if (plc?.ReadDeviceBlock2(register, leng, out data[0]) == 0)
            {
                return true;
            }

            return false;
        }

        public LocationRobot ReadPosition()
        {
            return null;
        }

        public bool WritePosition(LocationRobot point3D)
        {
            Task.Run(() =>
            {


            });
            return true;
        }

        public void AddRegisterRead( RegisterPlc2PC plc, PLCIOCollection pLCIOs)
        {
            plc?.PlcIOs?.AddRange(pLCIOs.ToArray());
            plc?.ClassifyRegister();
        }

        public void AddRegisterWrite(RegisterPc2Plc plc, PLCIOCollection pLCIOs)
        {
            RegisterWrite?.PlcIOs?.AddRange(pLCIOs.ToArray());
            RegisterWrite?.ClassifyRegister();
        }

        #region Utility
        public void SetValueRegister(object value,int index, string title, bool isSendPLC = true, EnumReadOrWrite enumReadOrWrite = EnumReadOrWrite.WRITE)
        {
            PLCIOCollection PLCIOs = null;
            switch (enumReadOrWrite)
            {
                case EnumReadOrWrite.READ:
                    PLCIOs = RegisterRead.PlcIOs;
                    break;
                case EnumReadOrWrite.WRITE:
                    PLCIOs = RegisterWrite.PlcIOs;
                    break;
                default:
                    break;
            }
            var r = PLCIOs.Where(p => p.Title == title && p.ReadOrWrite == enumReadOrWrite && p.IndexPLC == index);
            if (r == null || r.Count() == 0)
            {
                return;
            }
            var alive = r?.First();
            if (alive != null)
            {
                alive.CurrentValue = value;
                if (isSendPLC)
                {
                    SendDataToPlc(alive);

                }
            }

        }


        public object GetValueRegister(int index, string title, bool isReadPLC = true, EnumReadOrWrite enumReadOrWrite = EnumReadOrWrite.READ)
        {
            PLCIOCollection PLCIOs = null;
   
            
                    switch (enumReadOrWrite)
                    {
                        case EnumReadOrWrite.READ:
                            PLCIOs = RegisterRead.PlcIOs;
                            break;
                        case EnumReadOrWrite.WRITE:
                            PLCIOs = RegisterWrite.PlcIOs;
                            break;
                        default:
                            break;
                    }
            
            
            var r = PLCIOs.Where(p => p.Title == title && p.ReadOrWrite == EnumReadOrWrite.READ && p.IndexPLC ==index);
            if (r == null || r.Count() == 0)
            {
                return null;
            }           

            var alive = r?.First();
            if (alive != null)
            {
                if (isReadPLC)
                {
                    GetDataFromPlc(alive);
                    return alive.CurrentValue;
                }
                else
                {
                    return alive.CurrentValue;
                }
            }
            return null;
        }



        public void SendDataToPlc(PLCIO pLCIO)
        {
            switch (pLCIO.RegisterType)
            {
                case EnumRegisterType.DWORD:
                    int[] res = new int[1] { (int)pLCIO.CurrentValue };
                    WriteDWord($"D{pLCIO.RegisterPLC}", 2, pLCIO.IndexPLC, ref res);
                    break;
                case EnumRegisterType.WORD:
                    short[] res2 = new short[1] { (short)pLCIO.CurrentValue };
                    WriteWord($"D{pLCIO.RegisterPLC}", 1, pLCIO.IndexPLC, ref res2);
                    break;
                case EnumRegisterType.BIT:

                    WriteOneBit($"M{pLCIO.RegisterPLC}", (bool)pLCIO.CurrentValue, pLCIO.IndexPLC);
                    break;
                case EnumRegisterType.STRING:
                    if (pLCIO.CurrentValue.ToString().Trim().Length %2 == 1)
                    {
                        pLCIO.CurrentValue = pLCIO.CurrentValue.ToString()+ "\0";
                    }
                    short[] res4 = PLCExtension.ConvertStringToShortArr(((string)pLCIO.CurrentValue));
                    WriteWord($"D{pLCIO.RegisterPLC}", pLCIO.Size, pLCIO.IndexPLC, ref res4);
                    break;
                case EnumRegisterType.BITINWORD:
                    WriteOneBit($"D{pLCIO.RegisterPLC}.{pLCIO.Size.ToString("X")}", (bool)pLCIO.CurrentValue, pLCIO.IndexPLC);
                    break;
                case EnumRegisterType.FLOAT:

                    if (pLCIO.CurrentValue is float || pLCIO.CurrentValue is int || pLCIO.CurrentValue is double)
                    {
                        float[] res9 = new float[1] { float.Parse(pLCIO.CurrentValue.ToString()) };
                        WriteFloat($"D{pLCIO.RegisterPLC}", 1, pLCIO.IndexPLC, ref res9);
                    }

                    break;
                default:
                    break;
            }
        }

        public void GetDataFromPlc(PLCIO pLCIO)
        {
            switch (pLCIO.RegisterType)
            {
                case EnumRegisterType.DWORD:
                    int[] res = new int[1];
                    //ReadDWord($"D{pLCIO.RegisterPLC}", 2, out res);
                    ReadDWord(pLCIO.IndexPLC, $"D{pLCIO.RegisterPLC}", 2, out res);
                    pLCIO.CurrentValue = res[0];
                    break;
                case EnumRegisterType.WORD:
                    short[] res2 = new short[1];
                    ReadWord(pLCIO.IndexPLC, $"D{pLCIO.RegisterPLC}", 1, out res2);
                    pLCIO.CurrentValue = res2[0];
                    break;
                case EnumRegisterType.BIT:

                    pLCIO.CurrentValue = ReadOneBit(pLCIO.IndexPLC, $"M{pLCIO.RegisterPLC}");
                    break;
                case EnumRegisterType.STRING:
                    short[] res4;
                    ReadWord(pLCIO.IndexPLC, $"D{pLCIO.RegisterPLC}", pLCIO.Size, out res4);
                    pLCIO.CurrentValue = PLCExtension.ConvertShortArrToString(res4);
                    break;
                case EnumRegisterType.BITINWORD:
                    pLCIO.CurrentValue = ReadOneBit(pLCIO.IndexPLC, $"D{pLCIO.RegisterPLC}.{pLCIO.Size}");
                    break;
                case EnumRegisterType.FLOAT:
                    float[] res9 = new float[1];
                    ReadFloat(pLCIO.IndexPLC, $"D{pLCIO.RegisterPLC}", 2, out res9);
                    pLCIO.CurrentValue = res9[0];
                    break;
                default:
                    break;
            }
        }

        public short[] ConvertStringToShortArr(string val)
        {
            if (val.Length % 2 == 1)
            {
                val += " ";
            }
            short[] arr = new short[val.Length / 2];
            for (int i = 0; i < val.Length; i += 2)
            {
                string s = val.Substring(i, 2);
                byte[] bytes = Encoding.ASCII.GetBytes(s);
                short sh = BitConverter.ToInt16(bytes, 0);
                arr[i / 2] = sh;
            }

            return arr;
        }

        public short[] ConvertOneDiameter(double value, string format = "000.00")
        {
            if (format == "0.00")
            {
                if (value >= 10)
                {
                    throw new Exception("value ko thể lớn hơn  9");
                }
            }

            string v = value.ToString(format) + ", ";
            short[] val = new short[v.Length / 2];
            if (v.Length % 2 == 1)
            {
                v += " ";
            }
            for (int i = 0; i < v.Length; i += 2)
            {
                string s = v.Substring(i, 2);
                byte[] bytes = Encoding.ASCII.GetBytes(s);
                short sh = BitConverter.ToInt16(bytes, 0);
                val[i / 2] = sh;
            }
            return val;
        }

        public void FromShort(short number, out byte byte1, out byte byte2)
        {
            byte2 = (byte)(number >> 8);
            byte1 = (byte)(number & 255);
        }

        public short[] ConvertOneDiameterEnd(double value, string format = "000.00")
        {
            if (format == "0.00")
            {
                if (value >= 10)
                {
                    throw new Exception("value ko thể lớn hơn  9");
                }
            }
            string v = value.ToString(format + "\r\n");
            short[] val = new short[v.Length / 2];
            if (v.Length % 2 == 1)
            {
                v += " ";
            }
            for (int i = 0; i < v.Length; i += 2)
            {
                string s = v.Substring(i, 2);
                byte[] bytes = Encoding.ASCII.GetBytes(s);
                short sh = BitConverter.ToInt16(bytes, 0);
                val[i / 2] = sh;
            }
            return val;
        }

        private object _lockWrite = new object();

        public void SendPlcWhenDataChanged<T>(T newValue, ref T oldValue, string registerPlc)
        {
            //lock (_lockWrite)
            {
                // if (!newValue.Equals(oldValue))
                {
                    oldValue = newValue;
                    if (oldValue is Boolean v)
                    {
                        Task.Run(() =>
                        {
                            var result = PlcWrite1.SetDevice(registerPlc, v == true ? 1 : 0);
                            //Console.WriteLine($"WritePLc {registerPlc}");
                        });
                    }
                }
            }
        }
        public static double ConvertWordToPul(double unit_mm)
        {
            return unit_mm * 10000 / 20;
        }

        public static double ConvertPulToWord(double unit_mm)
        {
            return unit_mm / 10000 * 20;
        }

        public void Triggercamera(int index)
        {

        }

        public void HaveNewSample()
        {
            HaveNewSampleEvent?.Invoke(this, null);
        }


        public void JogRobot(int indexAxis, EnumJog jogType, EnumSpeedType speedType, double speed)
        {

        }

        public void JogStop()
        {

        }

        public void AddRegisterRead(PLCIOCollection pLCIOs)
        {
            throw new NotImplementedException();
        }


        #endregion

    }

}