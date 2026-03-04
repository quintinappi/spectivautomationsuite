using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SpectivLauncher
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static void Main()
        {
            // Hide console window
            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(handle, 0);

            try
            {
                // Use Launch_UI.ps1 from the same directory as the executable
                string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string scriptPath = System.IO.Path.Combine(exeDir, "Launch_UI.ps1");

                if (!System.IO.File.Exists(scriptPath))
                {
                    MessageBox.Show(
                        "Launch_UI.ps1 not found in root directory!\n\nPath: " + scriptPath,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }

                // Start PowerShell
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "powershell.exe";
                psi.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + scriptPath + "\"";
                psi.UseShellExecute = false;
                psi.CreateNoWindow = false;

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Failed to launch UI:\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
