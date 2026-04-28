using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

class Program
{
    [STAThread]
    static int Main()
    {
        try
        {
            // Get the directory where the .exe is located
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            string scriptPath = Path.Combine(exeDir, "proxy_installer.ps1");

            // Verify the script exists
            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    $"Error: proxy_installer.ps1 not found in {exeDir}\n\n" +
                    "Make sure the script is in the same folder as this .exe",
                    "Script Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return 1;
            }

            // PowerShell command: run the script with admin elevation
            string psCommand = $@"
                $scriptPath = '{scriptPath}'
                if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {{
                    Start-Process powershell -ArgumentList ""-NoProfile -ExecutionPolicy Bypass -File '$scriptPath'"" -Verb RunAs
                    exit
                }}
                & '$scriptPath'
            ";

            // Run PowerShell
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                UseShellExecute = false,
                RedirectStandardOutput = false,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal
            };

            using (var process = Process.Start(psi))
            {
                process?.WaitForExit();
                return process?.ExitCode ?? 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error launching installer:\n\n{ex.Message}",
                "Launch Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return 1;
        }
    }
}
