using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Winsomnia.Properties;

namespace Winsomnia
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            using var mutex = new Mutex(true, "WinsomniaSingleton", out bool createdNew);
            if (!createdNew)
                return; // Already running

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.Run(new WinsomniaContext());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"A fatal error occurred: {ex.Message}", "Winsomnia Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static class PowerAvailability
        {
            [Flags]
            public enum EXECUTION_STATE : uint
            {
                ES_SYSTEM_REQUIRED = 0x00000001,
                ES_DISPLAY_REQUIRED = 0x00000002,
                ES_AWAYMODE_REQUIRED = 0x00000040,
                ES_CONTINUOUS = 0x80000000
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

            /// <summary>
            /// Prevents the system from going to sleep or turning off the display.
            /// </summary>
            /// <exception cref="Win32Exception">Thrown if the OS call fails.</exception>
            public static void PreventSleep()
            {
                var result = SetThreadExecutionState(
                    EXECUTION_STATE.ES_CONTINUOUS |
                    EXECUTION_STATE.ES_SYSTEM_REQUIRED |
                    EXECUTION_STATE.ES_DISPLAY_REQUIRED);

                if (result == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new Win32Exception(errorCode, $"Failed to set thread execution state (Prevent Sleep). Error Code: {errorCode}");
                }
            }

            /// <summary>
            /// Allows the system to sleep normally again.
            /// </summary>
            public static void RestoreSleep()
            {
                // We only need ES_CONTINUOUS to reset the state
                var result = SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS);

                if (result == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    // We might choose to log this rather than crash, but for strictness:
                    throw new Win32Exception(errorCode, $"Failed to reset execution state. Error Code: {errorCode}");
                }
            }
        }

        // UI Context
        private sealed class WinsomniaContext : ApplicationContext
        {
            private readonly NotifyIcon _trayIcon;

            public WinsomniaContext()
            {
                var menu = new ContextMenuStrip();
                menu.Items.Add("Exit", null, OnExit);

                // Fallback to system icon if resource is missing
                Icon appIcon = Resources.winsomnia ?? SystemIcons.Application;

                _trayIcon = new NotifyIcon
                {
                    Icon = appIcon,
                    Text = "Winsomnia (Initializing...)",
                    ContextMenuStrip = menu,
                    Visible = true
                };

                EnableInsomnia();
            }

            private void EnableInsomnia()
            {
                try
                {
                    PowerAvailability.PreventSleep();
                    _trayIcon.Text = "Winsomnia (Active - PC will not sleep)";
                }
                catch (Win32Exception ex)
                {
                    _trayIcon.Text = "Winsomnia (Error)";
                    _trayIcon.ShowBalloonTip(3000, "Activation Failed",
                        $"Could not keep system awake. Windows Error: {ex.Message}", ToolTipIcon.Error);
                }
            }

            private void OnExit(object? sender, EventArgs e)
            {
                try
                {
                    // Attempt to reset state before exiting
                    PowerAvailability.RestoreSleep();
                }
                catch
                {
                    // Swallow error on exit, as the thread termination 
                    // by the OS cleans up execution states anyway.
                }
                finally
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                    ExitThread();
                }
            }
        }
    }
}