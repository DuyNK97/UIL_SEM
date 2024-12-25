using System.IO;
using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CIM.Class
{
    public static class Global
    {
        public static int CurrentModeBox3 = 0; //0: normal, 1 rework
        public static int CurrentStateBox3 = 0; // 0-offline, 1 online

        public static int CurrentModeBox4 = 0; //0: normal, 1 rework
        public static int CurrentStateBox4 = 0; // 0-offline, 1 online

        public static int IsCheckNAS = 1;

        public static int AutoDeleteCSV = 0;

        public static int DayDeleteCSV = 90;

        public static string CSVD = @"D:\mes_automaination_svc";

        public static string CSV = @"Z:\";

        private static readonly object[] _lockWriteBox = new object[4]
        {
            new object(),
            new object(),
            new object(),
            new object()
        };

        private static readonly object _lockData = new object();

        public static void WriteLogBox(string logFilePath, int boxIndex, params string[] logMessages)
        {
            lock (_lockWriteBox[boxIndex]) // Khóa tương ứng với file log được chọn
            {
                try
                {
                    logFilePath = Path.Combine(logFilePath, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"));

                    if (!Directory.Exists(logFilePath))
                    {
                        Directory.CreateDirectory(logFilePath);
                    }

                    logFilePath = Path.Combine(logFilePath, DateTime.Now.ToString("dd") + ".csv");

                    using (StreamWriter writer = new StreamWriter(logFilePath, true, new UTF8Encoding(true)))
                    {
                        string logEntry = string.Join(";", logMessages);
                        writer.WriteLine(logEntry);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error writing log: {ex.Message}");
                }
            }
        }

        public static void WriteFileToTxt(string filePath, Dictionary<string, string> values)
        {
            lock (_lockData)
            {
                try
                {
                    var lines = File.ReadAllLines(filePath).ToList();
                    var keysToUpdate = values.Keys.ToList();

                    var updatedKeys = new HashSet<string>();

                    for (int i = 0; i < lines.Count; i++)
                    {
                        var parts = lines[i].Split(new[] { '=' }, 2);
                        if (parts.Length == 2)
                        {
                            string key = parts[0].Trim();
                            if (values.ContainsKey(key))
                            {
                                lines[i] = $"{key}= {values[key]}";
                                updatedKeys.Add(key);
                            }
                        }
                    }

                    foreach (var key in keysToUpdate)
                    {
                        if (!updatedKeys.Contains(key))
                        {
                            lines.Add($"{key}= {values[key]}");
                        }
                    }

                    File.WriteAllLines(filePath, lines);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error can not write value to file txt: {ex.Message}");
                }
            }
        }

        public static Dictionary<string, string> ReadValueFileTxt(string filePath, List<string> keys)
        {
            Dictionary<string, string> values = new Dictionary<string, string>();

            try
            {
                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    string[] parts = line.Split('=');

                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();

                        if (keys.Contains(key))
                        {
                            values[key] = parts[1].Trim();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error can not read value from file txt: {ex.Message}");
            }

            return values;
        }

        public static string GetFilePathSetting()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "setting.txt");
        }
    }
}
