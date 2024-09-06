using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnilityCommand;

namespace UILAlignProject.PLC
{
    public static class Extensions
    {
        public static string ConvertDateToFormatPlc(this DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd_HH-mm-ss");
        }

        public static int[] ConvertFloatToWord(this float floatNumber)
        {
            byte[] floatBytes = BitConverter.GetBytes(floatNumber);
            byte[] highWordBytes = { floatBytes[2], floatBytes[3] };
            byte[] lowWordBytes = { floatBytes[0], floatBytes[1] };
            int[] returnValue = { BitConverter.ToInt16(lowWordBytes, 0), BitConverter.ToInt16(highWordBytes, 0) };
            return returnValue;
        }

        public static float ConvertWordToFloat(this int[] register)
        {
            byte[] highWordByte = BitConverter.GetBytes(register[1]);
            byte[] lowWordByte = BitConverter.GetBytes(register[0]);
            byte[] valueFloatByte = { lowWordByte[0], lowWordByte[1], highWordByte[0], highWordByte[1] };
            float value = BitConverter.ToSingle(valueFloatByte, 0);
            return value;
        }

    }
}