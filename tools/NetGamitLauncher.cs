using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

[assembly: AssemblyTitle("Net-Gamit")]
[assembly: AssemblyDescription("Windows Connectivity Tool")]
[assembly: AssemblyCompany("Thermo Fisher Scientific")]
[assembly: AssemblyProduct("Net-Gamit")]
[assembly: AssemblyCopyright("Copyright © Ramon Veluz")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class NetGamitLauncher
{
    private const string EmbeddedScriptResourceName = "NetGamitScript.ps1";

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string scriptText = ReadEmbeddedScript();
            string scriptPath = WriteScriptToTemp(scriptText);

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "powershell.exe";
            startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -File " + Quote(scriptPath) + BuildForwardedArguments(args);
            startInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            using (Process process = Process.Start(startInfo))
            {
                return 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Net-Gamit could not start.\r\n\r\n" + ex.Message,
                "Net-Gamit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static string ReadEmbeddedScript()
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        using (Stream stream = assembly.GetManifestResourceStream(EmbeddedScriptResourceName))
        {
            if (stream == null)
            {
                throw new InvalidOperationException("Embedded Net-Gamit script was not found.");
            }

            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
            {
                return reader.ReadToEnd();
            }
        }
    }

    private static string WriteScriptToTemp(string scriptText)
    {
        string hash = ComputeSha256(scriptText).Substring(0, 16);
        string folder = Path.Combine(Path.GetTempPath(), "Net-Gamit");
        Directory.CreateDirectory(folder);

        string scriptPath = Path.Combine(folder, "Net-Gamit-" + hash + ".ps1");
        if (!File.Exists(scriptPath))
        {
            File.WriteAllText(scriptPath, scriptText, new UTF8Encoding(true));
        }

        return scriptPath;
    }

    private static string ComputeSha256(string text)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            byte[] hash = sha256.ComputeHash(bytes);
            StringBuilder builder = new StringBuilder(hash.Length * 2);

            for (int i = 0; i < hash.Length; i++)
            {
                builder.Append(hash[i].ToString("x2"));
            }

            return builder.ToString();
        }
    }

    private static string BuildForwardedArguments(string[] args)
    {
        if (args == null || args.Length == 0)
        {
            return string.Empty;
        }

        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < args.Length; i++)
        {
            builder.Append(' ');
            builder.Append(Quote(args[i]));
        }

        return builder.ToString();
    }

    private static string Quote(string value)
    {
        if (value == null)
        {
            return "\"\"";
        }

        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
