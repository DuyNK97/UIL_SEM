using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace CIM.Class
{
    public static class Global
    {
        public static int CurrentModeBox3 = 0; //0: normal, 1 rework
        public static int CurrentStateBox3 = 0; // 0-offline, 1 online

        public static int CurrentModeBox4 = 0; //0: normal, 1 rework
        public static int CurrentStateBox4 = 0; // 0-offline, 1 online

        public static int CurrModeB1 = 0;
        public static int CurrModeB2 = 0;
        public static int CurrModeB3 = 0;
        public static int CurrModeB4 = 0;

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
						try
						{
							Directory.CreateDirectory(logFilePath);
						}
						catch (Exception dirEx)
						{
							FormMain.WriteLog($"[{DateTime.Now}] Error creating directory {logFilePath}: {dirEx.Message}\n");
							return;
						}
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
				    FormMain.WriteLog($"Error writing log: {ex.Message}");
					Console.WriteLine($"Error writing log: {ex.Message}");
                }
            }
        }

        public static void WriteLogBox_bak(string logFilePath, int boxIndex, params string[] logMessages)
        {
            lock (_lockWriteBox[boxIndex]) // Khóa tương ứng với file log được chọn
            {
                try
                {
                    logFilePath = Path.Combine(logFilePath, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"));

                    if (!Directory.Exists(logFilePath))
                    {
                        try
                        {
                            Directory.CreateDirectory(logFilePath);
                        }
                        catch (Exception dirEx)
                        {
                            FormMain.WriteLog($"[{DateTime.Now}] Error creating directory {logFilePath}: {dirEx.Message}\n");
                            return;
                        }
                    }

                    logFilePath = Path.Combine(logFilePath, DateTime.Now.ToString("dd") + ".csv");

                    WriteToFile(logFilePath, logMessages);
                }
                catch (Exception ex)
                {
                    FormMain.WriteLog($"Error writing log: {ex.Message}");
                    Console.WriteLine($"Error writing log: {ex.Message}");
                }
            }
        }
        private static void WriteToFile(string filePath, string[] logMessages)
        {
            const int maxRetries = 3;
            const int baseDelayMs = 50;
            const int bufferSize = 8192; // 8KB buffer

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.Read, bufferSize))
                    using (BufferedStream bs = new BufferedStream(fs, bufferSize))
                    using (StreamWriter writer = new StreamWriter(bs, new UTF8Encoding(true)))
                    {
                        string logEntry = string.Join(";", logMessages);
                        writer.WriteLine(logEntry);
                        writer.Flush();
                        bs.Flush();
                        fs.Flush();
                    }
                    return;
                }
                catch (IOException ioEx) when (attempt < maxRetries - 1)
                {
                    int delay = baseDelayMs * (attempt + 1);
                    Thread.Sleep(delay);
                }
                catch (Exception ex)
                {
                    FormMain.WriteLog($"[{DateTime.Now}] Error writing to {filePath}: {ex.Message}");
                    throw;
                }
            }

            throw new IOException($"Failed to write to file {filePath} after {maxRetries} attempts");
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
