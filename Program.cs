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
            var mutex = new Mutex(true, "WinsomniaSingleton", out bool createdNew);
            if (!createdNew)
            {
                mutex.Dispose();
                return; // Already running
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.Run(new WinsomniaContext(mutex));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"A fatal error occurred: {ex.Message}", "Winsomnia Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                mutex.Dispose();
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
            private readonly ContextMenuStrip _contextMenu;
            private readonly Mutex _mutex;
            private readonly ToolStripMenuItem _toggleMenuItem;
            private bool _isEnabled;

            public WinsomniaContext(Mutex mutex)
            {
                _mutex = mutex;
                _contextMenu = new ContextMenuStrip();

                _toggleMenuItem = new ToolStripMenuItem("Disable", null, OnToggle);
                _contextMenu.Items.Add(_toggleMenuItem);
                _contextMenu.Items.Add(new ToolStripSeparator());
                _contextMenu.Items.Add("Exit", null, OnExit);

                // Fallback to system icon if resource is missing
                Icon appIcon = Resources.winsomnia ?? SystemIcons.Application;

                _trayIcon = new NotifyIcon
                {
                    Icon = appIcon,
                    Text = "Winsomnia (Initializing...)",
                    ContextMenuStrip = _contextMenu,
                    Visible = true
                };

                SetInsomniaState(true);
            }

            private void SetInsomniaState(bool enable)
            {
                try
                {
                    if (enable)
                    {
                        PowerAvailability.PreventSleep();
                        _isEnabled = true;
                        _trayIcon.Text = "Winsomnia (Active - PC will not sleep)";
                        _toggleMenuItem.Text = "Disable";
                    }
                    else
                    {
                        PowerAvailability.RestoreSleep();
                        _isEnabled = false;
                        _trayIcon.Text = "Winsomnia (Inactive - PC will sleep normally)";
                        _toggleMenuItem.Text = "Enable";
                    }
                }
                catch (Win32Exception ex)
                {
                    _trayIcon.Text = "Winsomnia (Error)";
                    _trayIcon.ShowBalloonTip(3000, enable ? "Activation Failed" : "Deactivation Failed",
                        $"Could not change sleep state. Windows Error: {ex.Message}", ToolTipIcon.Error);
                }
            }

            private void OnToggle(object? sender, EventArgs e)
            {
                SetInsomniaState(!_isEnabled);
            }

            private void OnExit(object? sender, EventArgs e)
            {
                try
                {
                    // Attempt to reset state before exiting (only if currently enabled)
                    if (_isEnabled)
                    {
                        PowerAvailability.RestoreSleep();
                    }
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
                    _contextMenu.Dispose();
                    ExitThread();
                }
            }
        }
    }
}