// SharePoint Storage Tree - Windows launcher (SPStorageTree.exe).
// Finds PowerShell 7, then opens the Start-SPStorageTree window with no console.
// Module lookup: next to this exe, then ..\module\SPStorageTree (repo/ZIP layout), then the installed module.
// Build: build.ps1 compiles this with the .NET Framework C# compiler that ships with Windows.
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("SharePoint Storage Tree")]
[assembly: AssemblyProduct("SharePoint Storage Tree")]
[assembly: AssemblyDescription("Opens the SharePoint Storage Tree window (needs PowerShell 7.4+ and PnP.PowerShell).")]
[assembly: AssemblyCopyright("(c) G Vijai Kumar. MIT License.")]
[assembly: AssemblyVersion("0.2.0.0")]
[assembly: AssemblyFileVersion("0.2.0.0")]

static class Launcher
{
    const string Title = "SharePoint Storage Tree";

    [STAThread]
    static int Main()
    {
        string pwsh = FindPwsh();
        if (pwsh == null)
        {
            MessageBox.Show(
                "SharePoint Storage Tree needs PowerShell 7.4 or later, which isn't installed on this computer.\n\n" +
                "Install it with this command (in Command Prompt or Windows PowerShell):\n\n" +
                "    winget install Microsoft.PowerShell\n\n" +
                "or download it from https://aka.ms/powershell, then run SPStorageTree again.",
                Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 1;
        }

        string psd1 = FindModule();
        string import = psd1 != null
            ? "Import-Module -Name '" + psd1.Replace("'", "''") + "' -Force"
            : "Import-Module -Name SPStorageTree";

        // Everything runs inside pwsh; any failure (old PowerShell, missing PnP.PowerShell, module not found)
        // is shown in a message box with the fix, never in an invisible console.
        string script =
            "$ErrorActionPreference = 'Stop'\n" +
            "try {\n" +
            "  if ($PSVersionTable.PSVersion -lt [version]'7.4') { throw \"PowerShell 7.4 or later is needed (this is $($PSVersionTable.PSVersion)). Update it with: winget upgrade Microsoft.PowerShell\" }\n" +
            "  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version -ge [version]'3.0' })) {\n" +
            "    throw \"PnP.PowerShell 3.x is needed. Open PowerShell 7 and run:`n`n    Install-PSResource PnP.PowerShell -TrustRepository`n`nthen start SPStorageTree again.\" }\n" +
            "  " + import + "\n" +
            "  Start-SPStorageTree\n" +
            "} catch {\n" +
            "  Add-Type -AssemblyName System.Windows.Forms\n" +
            "  $m = $_.Exception.Message\n" +
            "  if ($m -match 'SPStorageTree' -and $m -match 'not loaded|not found|was not loaded') { $m += \"`n`nInstall it with:  Install-PSResource SPStorageTree -TrustRepository\" }\n" +
            "  [void][System.Windows.Forms.MessageBox]::Show($m, '" + Title + "', 'OK', 'Warning')\n" +
            "}\n";

        var psi = new ProcessStartInfo
        {
            FileName = pwsh,
            Arguments = "-NoProfile -NoLogo -STA -ExecutionPolicy Bypass -EncodedCommand " +
                        Convert.ToBase64String(Encoding.Unicode.GetBytes(script)),
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };
        try { Process.Start(psi); }
        catch (Exception ex)
        {
            MessageBox.Show("Couldn't start PowerShell 7 (" + pwsh + "):\n\n" + ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }
        return 0;
    }

    static string FindModule()
    {
        string dir = AppDomain.CurrentDomain.BaseDirectory;
        string[] candidates =
        {
            Path.Combine(dir, "SPStorageTree.psd1"),                                  // exe inside the module folder
            Path.Combine(dir, "module", "SPStorageTree", "SPStorageTree.psd1"),       // exe at the repo/ZIP root
            Path.Combine(dir, "..", "module", "SPStorageTree", "SPStorageTree.psd1")  // exe in a subfolder of the repo
        };
        foreach (string c in candidates)
            if (File.Exists(c)) return Path.GetFullPath(c);
        return null;                                                                  // fall back to the installed module
    }

    static string FindPwsh()
    {
        string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string[] fixedPaths =
        {
            Path.Combine(pf, "PowerShell", "7", "pwsh.exe"),
            Path.Combine(pf, "PowerShell", "7-preview", "pwsh.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "pwsh.exe")
        };
        foreach (string p in fixedPaths)
            if (File.Exists(p)) return p;
        string path = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (string d in path.Split(Path.PathSeparator))
        {
            try
            {
                string p = Path.Combine(d.Trim(), "pwsh.exe");
                if (d.Trim().Length > 0 && File.Exists(p)) return p;
            }
            catch (ArgumentException) { }
        }
        return null;
    }
}
