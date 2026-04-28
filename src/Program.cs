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
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            string scriptPath = Path.Combine(exeDir, "proxy_installer.ps1");

            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    $"Error: proxy_installer.ps1 not found\n\nExpected: {scriptPath}\n\n" +
                    "Make sure the .ps1 script is in the same folder as this .exe file.",
                    "Script Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return 1;
            }

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoExit -NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
                UseShellExecute = true,
                Verb = "runas"
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
                $"Error launching installer:\n\n{ex.Message}\n\n" +
                "Make sure both PracticePerfect-Installer.exe and proxy_installer.ps1 are in the same folder.",
                "Launch Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return 1;
        }
    }
}
