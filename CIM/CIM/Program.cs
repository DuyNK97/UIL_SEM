using ComponentFactory.Krypton.Toolkit;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace CIM
{
    internal static class Program
    {
        static Mutex mutex = new Mutex(true, "{b74f4e7a-7326-4b84-8d33-2f2d56bdf1f302052000}");

        [STAThread]
        static void Main()
        {
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                AppUtil.CatchException();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FormMain());
                mutex.ReleaseMutex();
            }
            else
            {
                MessageBox.Show("The program is running!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    public class AppUtil
    {
        protected const int WS_SHOWNORMAL = 1;

        private static Mutex _SingleMutex = null;

        public static bool CheckRunning()
        {
            string processName = Process.GetCurrentProcess().ProcessName;
            bool createdNew = false;
            _SingleMutex = new Mutex(true, processName, out createdNew);
            if (!createdNew)
            {
                Process runningInstance = GetRunningInstance();
                if (runningInstance != null)
                {
                    FlashRunningInstance(runningInstance);
                }
                return true;
            }
            return false;
        }

        [DllImport("User32.dll")]
        protected static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);

        [DllImport("User32.dll")]
        protected static extern bool SetForegroundWindow(IntPtr hWnd);

        protected static Process GetRunningInstance()
        {
            Process currentProcess = Process.GetCurrentProcess();
            string fileName = currentProcess.MainModule.FileName;
            Process[] processesByName = Process.GetProcessesByName(currentProcess.ProcessName);
            Process[] array = processesByName;
            foreach (Process process in array)
            {
                if (process.MainModule.FileName == fileName && process.Id != currentProcess.Id)
                {
                    return process;
                }
            }
            return null;
        }

        protected static bool ShowRunningInstance(Process instance)
        {
            ShowWindowAsync(instance.MainWindowHandle, 1);
            return SetForegroundWindow(instance.MainWindowHandle);
        }

        [DllImport("user32.dll")]
        protected static extern bool FlashWindow(IntPtr hWnd, bool bInvert);

        protected static bool FlashRunningInstance(Process instance)
        {
            return FlashWindow(instance.MainWindowHandle, true);
        }

        public static bool CheckRunInVS()
        {
            return Debugger.IsAttached;
        }

        public static void CatchException()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                Exception ex = e.ExceptionObject as Exception;
                KryptonMessageBox.Show(ex.ToString());
            }
            catch (Exception ex2)
            {
                KryptonMessageBox.Show(ex2.ToString());
            }
        }

        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            try
            {
                KryptonMessageBox.Show(e.Exception.ToString());
            }
            catch (Exception ex)
            {
                KryptonMessageBox.Show(ex.ToString());
            }
        }
    }
}
