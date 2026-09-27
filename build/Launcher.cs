using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

class Launcher {
    [STAThread] static int Main(string[] args) {
        try {
            string root = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string project = root;
            string current = Path.Combine(root, "current.txt");
            if (File.Exists(current)) {
                string version = File.ReadAllText(current).Trim();
                if (!Regex.IsMatch(version, "^[a-f0-9]{16}$")) throw new Exception("The installation record is invalid. Run the installer again.");
                project = Path.Combine(root, "versions", version);
            }
            string mode = args.Length == 0 ? "" : args[0];
            if (mode != "" && mode != "--ensure-month" && mode != "--self-test" && mode != "--ui-self-test")
                throw new Exception("Unsupported launch option.");
            bool check = mode == "--ensure-month" || mode == "--self-test";
            var start = new ProcessStartInfo();
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.WorkingDirectory = project;
            start.EnvironmentVariables["TNG_AGENT_INSTALL"] = root;
            start.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            if (check) {
                start.FileName = Path.Combine(project, "runtime", "python.exe");
                start.Arguments = mode == "--ensure-month" ? "-B -E \"" + Path.Combine(project, "app", "agent.py") + "\" --ensure-month" :
                    "-B -I -c \"import sys,pdfplumber,openpyxl,sqlite3;print('Bundled runtime OK: Python '+sys.version.split()[0]+'; openpyxl '+openpyxl.__version__)\"";
            } else {
                start.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
                start.Arguments = "-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + Path.Combine(project, "app", "Start.ps1") + "\"";
                if (mode == "--ui-self-test") start.Arguments += " -SelfTest";
            }
            using (var process = Process.Start(start)) {
                if (mode != "") { process.WaitForExit(); return process.ExitCode; }
            }
            return 0;
        } catch (Exception exc) {
            if (args.Length == 0) MessageBox.Show(exc.Message, "TNG Expense Agent", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else Console.Error.WriteLine(exc.Message);
            return 1;
        }
    }
}
