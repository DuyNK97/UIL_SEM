using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnilityCommand;
using UnilityCommand.Plc;
using UnilityCommand.Plc.Mitsubishi;
using UnilityCommand.SubControl;

namespace UILAlignProject.PLC
{
    public class RegisterPc2Plc : Bindable
    {
        public RegisterPc2Plc(SingleTonPlcControl plc)
        {
            _plc = plc;
        }

       


        private SingleTonPlcControl _plc;

       

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

        public List<PLCIOCollection> RegisterBit = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterWord = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterDWord = new List<PLCIOCollection>();
        public List<PLCIOCollection> RegisterString = new List<PLCIOCollection>();

        public void ClassifyRegister()
        {
            ClassifyBit(RegisterBit, EnumRegisterType.BIT);
            ClassifyBit(RegisterWord, EnumRegisterType.WORD);
            ClassifyBit(RegisterDWord, EnumRegisterType.DWORD);
            ClassifyString(RegisterString);
        }
        private void ClassifyString(List<PLCIOCollection> pLCIOs)
        {

            pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(_plcIOs.Where(p => p.RegisterType == EnumRegisterType.STRING && p.ReadOrWrite == EnumReadOrWrite.WRITE).ToArray());
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
        private void ClassifyBit(List<PLCIOCollection> pLCIOs, EnumRegisterType enumRegisterType)
        {

            pLCIOs.Clear();

            PLCIOCollection io = new PLCIOCollection();
            io.AddRange(_plcIOs.Where(p => p.RegisterType == enumRegisterType && p.ReadOrWrite == EnumReadOrWrite.WRITE).ToArray());
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



    }
}
