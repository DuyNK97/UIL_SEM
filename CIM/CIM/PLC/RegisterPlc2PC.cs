
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnilityCommand;
using UnilityCommand.Plc;
using UnilityCommand.Plc.Mitsubishi;
using UnilityCommand.SubControl;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace UILAlignProject.PLC
{
    public class RegisterPlc2PC : Bindable
    {
        private SingleTonPlcControl _plc;
        public event EventHandler OutMes;
        public RegisterPlc2PC(SingleTonPlcControl plc)
        {
            _plc = plc;

        }



        //public string HEART_BEAT = "M2500"; //Pulse 1 sec (flip-flop)


        private bool _heartBeat;

        [NotShowProperty]
        public bool HeartBeat
        {
            get { return _heartBeat; }
            set
            {
                _heartBeat = value;

                Notify();
            }
        }

        private PLCIOCollection _plcIOs = new PLCIOCollection();

        public PLCIOCollection PlcIOs
        {
            get { return _plcIOs; }
            set
            {
                _plcIOs = value;
                Notify();
            }
        }

        public List<PLCIOCollection> RegisterBit1 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterBitInWord1 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterWord1 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterDWord1 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterString1 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterFloat1 = new List<PLCIOCollection>();

        public List<PLCIOCollection> RegisterBit2 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterBitInWord2 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterWord2 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterDWord2 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterString2 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterFloat2 = new List<PLCIOCollection>();

        public List<PLCIOCollection> RegisterBit3 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterBitInWord3 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterWord3 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterDWord3 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterString3 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterFloat3 = new List<PLCIOCollection>();

        public List<PLCIOCollection> RegisterBit4 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterBitInWord4 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterWord4 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterDWord4 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterString4 = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterFloat4 = new List<PLCIOCollection>();

        public void ClassifyRegister()
        {
            var plc1= _plcIOs.Where(p=>p.IndexPLC==1).ToList();
            ClassifyBit(plc1,RegisterBit1, EnumRegisterType.BIT);
            ClassifyBit(plc1, RegisterWord1, EnumRegisterType.WORD);
            ClassifyDWord(plc1, RegisterDWord1, EnumRegisterType.DWORD);
            ClassifyDWord(plc1, RegisterFloat1, EnumRegisterType.FLOAT);
            ClassifyBitInWord(plc1, RegisterBitInWord1);
            ClassifyString(plc1, RegisterString1);

            var plc2= _plcIOs.Where(p => p.IndexPLC == 2).ToList();
            ClassifyBit(plc2, RegisterBit2, EnumRegisterType.BIT);
            ClassifyBit(plc2, RegisterWord2, EnumRegisterType.WORD);
            ClassifyDWord(plc2, RegisterDWord2, EnumRegisterType.DWORD);
            ClassifyDWord(plc2, RegisterFloat2, EnumRegisterType.FLOAT);
            ClassifyBitInWord(plc2, RegisterBitInWord2);
            ClassifyString(plc2, RegisterString2);

            var plc3 = _plcIOs.Where(p => p.IndexPLC == 3).ToList();
            ClassifyBit(plc3, RegisterBit3, EnumRegisterType.BIT);
            ClassifyBit(plc3, RegisterWord3, EnumRegisterType.WORD);
            ClassifyDWord(plc3, RegisterDWord3, EnumRegisterType.DWORD);
            ClassifyDWord(plc3, RegisterFloat3, EnumRegisterType.FLOAT);
            ClassifyBitInWord(plc3, RegisterBitInWord3);
            ClassifyString(plc3, RegisterString3);

            var plc4 = _plcIOs.Where(p => p.IndexPLC == 4).ToList();
            ClassifyBit(plc4, RegisterBit4, EnumRegisterType.BIT);
            ClassifyBit(plc4, RegisterWord4, EnumRegisterType.WORD);
            ClassifyDWord(plc4, RegisterDWord4, EnumRegisterType.DWORD);
            ClassifyDWord(plc4, RegisterFloat4, EnumRegisterType.FLOAT);
            ClassifyBitInWord(plc4, RegisterBitInWord4);
            ClassifyString(plc4, RegisterString4);
        }

        

        private void ClassifyString(List<PLCIO> source, List<PLCIOCollection> pLCIOs)
        {

            //pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(source.Where(p => p.RegisterType == EnumRegisterType.STRING && p.ReadOrWrite == EnumReadOrWrite.READ && p.IsUpdate ==true).ToArray());
            if (io == null || io.Count == 0)
            {
                return;
            }


            var bit2 = io.OrderBy(p => p.RegisterPLC).ToList();
            foreach (var item in bit2)
            {
                PLCIOCollection temp1 = new PLCIOCollection();
                temp1.Add(item);
                pLCIOs.Add(temp1);

            }

        }

        private void ClassifyBitInWord(List<PLCIO> source, List<PLCIOCollection> pLCIOs)
        {

            //pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(source.Where(p => p.RegisterType == EnumRegisterType.BITINWORD && p.ReadOrWrite == EnumReadOrWrite.READ && p.IsUpdate == true).ToArray());
            if (io == null || io.Count == 0)
            {
                return;
            }
            var split = io.GroupBy(p => p.RegisterPLC).ToList();
            foreach(var item in split)
            {
                PLCIOCollection temp = new PLCIOCollection();
                var g = item.ToArray();
                temp.AddRange(g);
               pLCIOs.Add((PLCIOCollection)temp);
            }

          
           

        }

        private void ClassifyBit(List<PLCIO> source, List<PLCIOCollection> pLCIOs, EnumRegisterType enumRegisterType)
        {

            //pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(source.Where(p => p.RegisterType == enumRegisterType && p.ReadOrWrite == EnumReadOrWrite.READ && p.IsUpdate == true).ToArray());
            if (io == null || io.Count == 0)
            {
                return;
            }


            var bit2 = io.OrderBy(p => p.RegisterPLC).ToList();
            int index = bit2[0].RegisterPLC;
            PLCIOCollection temp1 = new PLCIOCollection();
            if (bit2.Count == 1)
            {
                temp1.Add(bit2[0]);
                pLCIOs.Add(temp1);
                return;
            }
            temp1.Add(bit2[0]);
            foreach (PLCIO item in bit2)
            {

                if (item == io[0])
                {
                    continue;
                }
                if (item.RegisterPLC == index + 1)
                {
                    temp1.Add(item);
                    index++;
                }
                else
                {
                    pLCIOs.Add(temp1);
                    temp1 = new PLCIOCollection();
                    temp1.Add(item);
                    index = item.RegisterPLC;
                }

                if (item == bit2.Last())
                {
                    pLCIOs.Add(temp1);
                }
            }
        }

        private void ClassifyDWord(List<PLCIO> source, List<PLCIOCollection> pLCIOs, EnumRegisterType enumRegisterType)
        {

            //pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(source.Where(p => p.RegisterType == enumRegisterType && p.ReadOrWrite == EnumReadOrWrite.READ && p.IsUpdate == true).ToArray());
            if (io == null || io.Count == 0)
            {
                return;
            }


            var bit2 = io.OrderBy(p => p.RegisterPLC).ToList();
            int index = bit2[0].RegisterPLC;
            PLCIOCollection temp1 = new PLCIOCollection();
            if (bit2.Count == 1)
            {
                temp1.Add(bit2[0]);
                pLCIOs.Add(temp1);
                return;
            }
            temp1.Add(bit2[0]);
            foreach (PLCIO item in bit2)
            {

                if (item == io[0])
                {
                    continue;
                }
                if (item.RegisterPLC == index + 2)
                {
                    temp1.Add(item);
                    index += 2;
                }
                else
                {
                    pLCIOs.Add(temp1);
                    temp1 = new PLCIOCollection();
                    temp1.Add(item);
                    index = item.RegisterPLC;
                }

                if (item == bit2.Last())
                {
                    pLCIOs.Add(temp1);
                }
            }
        }

        public void pulseUp<T>(T oldValue, T newValue)
        {
            {
                oldValue = newValue;
                if (oldValue is Boolean v)
                {
                    Notify();
                }
            }
        }
    }
}
