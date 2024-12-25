using System;
using System.Collections.Generic;

namespace CIM
{
    public class DateMap
    {
        private Dictionary<int, char> yearMapping = new Dictionary<int, char>();
        private Dictionary<string, char> monthMapping = new Dictionary<string, char>();
        private Dictionary<int, char> dateMapping = new Dictionary<int, char>();

        public DateMap()
        {
            yearMapping.Add(2016, 'A');
            yearMapping.Add(2017, 'B');
            yearMapping.Add(2018, 'C');
            yearMapping.Add(2019, 'D');
            yearMapping.Add(2020, 'E');
            yearMapping.Add(2021, 'F');
            yearMapping.Add(2022, 'G');
            yearMapping.Add(2023, 'H');
            yearMapping.Add(2024, 'I');
            yearMapping.Add(2025, 'J');
            yearMapping.Add(2026, 'K');
            yearMapping.Add(2027, 'L');
            yearMapping.Add(2028, 'M');
            yearMapping.Add(2029, 'N');
            yearMapping.Add(2030, 'O');
            yearMapping.Add(2031, 'P');
            yearMapping.Add(2032, 'Q');
            yearMapping.Add(2033, 'R');
            yearMapping.Add(2034, 'S');
            yearMapping.Add(2035, 'T');
            yearMapping.Add(2036, 'U');
            yearMapping.Add(2037, 'V');
            yearMapping.Add(2038, 'W');
            yearMapping.Add(2039, 'X');
            yearMapping.Add(2040, 'Y');
            yearMapping.Add(2041, 'Z');

            monthMapping.Add("JAN", '1');
            monthMapping.Add("FEB", '2');
            monthMapping.Add("MAR", '3');
            monthMapping.Add("APR", '4');
            monthMapping.Add("MAY", '5');
            monthMapping.Add("JUN", '6');
            monthMapping.Add("JUL", '7');
            monthMapping.Add("AUG", '8');
            monthMapping.Add("SEP", '9');
            monthMapping.Add("OCT", 'A');
            monthMapping.Add("NOV", 'B');
            monthMapping.Add("DEC", 'C');

            dateMapping.Add(1, '1');
            dateMapping.Add(2, '2');
            dateMapping.Add(3, '3');
            dateMapping.Add(4, '4');
            dateMapping.Add(5, '5');
            dateMapping.Add(6, '6');
            dateMapping.Add(7, '7');
            dateMapping.Add(8, '8');
            dateMapping.Add(9, '9');
            dateMapping.Add(10, 'A');
            dateMapping.Add(11, 'B');
            dateMapping.Add(12, 'C');
            dateMapping.Add(13, 'D');
            dateMapping.Add(14, 'E');
            dateMapping.Add(15, 'F');
            dateMapping.Add(16, 'G');
            dateMapping.Add(17, 'H');
            dateMapping.Add(18, 'I');
            dateMapping.Add(19, 'J');
            dateMapping.Add(20, 'K');
            dateMapping.Add(21, 'L');
            dateMapping.Add(22, 'M');
            dateMapping.Add(23, 'N');
            dateMapping.Add(24, 'O');
            dateMapping.Add(25, 'P');
            dateMapping.Add(26, 'Q');
            dateMapping.Add(27, 'R');
            dateMapping.Add(28, 'S');
            dateMapping.Add(29, 'T');
            dateMapping.Add(30, 'U');
            dateMapping.Add(31, 'V');
        }

        public string GetValue(DateTime date)
        {
            int year = date.Year;
            string month = date.ToString("MMM").ToUpper();
            int day = date.Day;

            char yearValue = yearMapping.ContainsKey(year) ? yearMapping[year] : '\0';
            char monthValue = monthMapping.ContainsKey(month) ? monthMapping[month] : '\0';
            char dayValue = dateMapping.ContainsKey(day) ? dateMapping[day] : '\0';

            return $"{yearValue}{monthValue}{dayValue}";
        }
    }
}
