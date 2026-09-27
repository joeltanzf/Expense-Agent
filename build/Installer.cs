using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

class Installer {
    const string Marker = "TNGExpenseAgent-standalone-v1";
    static string Quote(string s) { return "'" + s.Replace("'", "''") + "'"; }
    static void Powershell(string script) {
        var p = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe"));
        p.Arguments = "-NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes("$ErrorActionPreference='Stop';"+script));
        p.UseShellExecute=false; p.CreateNoWindow=true; p.RedirectStandardError=true;
        using (var child=Process.Start(p)) { string error=child.StandardError.ReadToEnd(); child.WaitForExit(); if(child.ExitCode!=0) throw new Exception("The app was installed, but Windows could not create the shortcut. Open TNG Expense Agent.exe in its installation folder. " + error); }
    }
    [STAThread] static int Main(string[] args) {
        bool test=args.Length==2 && args[0]=="--test-install";
        string root=test ? Path.GetFullPath(args[1]) : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Programs","TNGExpenseAgent");
        string stage=null;
        try {
            if(args.Length!=0 && !test) throw new Exception("Unsupported installer option.");
            if(!Environment.Is64BitOperatingSystem) throw new Exception("This version requires 64-bit Windows.");
            if(!test && MessageBox.Show("Install TNG Expense Agent for this Windows account?\n\nNo extra downloads or subscriptions are needed. Existing expense data is kept separately.\n\nInstallation folder:\n"+root,"TNG Expense Agent Setup",MessageBoxButtons.OKCancel,MessageBoxIcon.Information)!=DialogResult.OK) return 0;
            if(Directory.Exists(root) && Directory.GetFileSystemEntries(root).Length>0 && (!File.Exists(Path.Combine(root,".tng-install")) || File.ReadAllText(Path.Combine(root,".tng-install"))!=Marker))
                throw new Exception("The destination contains unrelated files. Installation stopped without replacing them.");
            // Never traverse redirected folders during extraction or installation.
            for(DirectoryInfo d=new DirectoryInfo(root); d!=null; d=d.Parent)
                if(d.Exists && (d.Attributes & FileAttributes.ReparsePoint)!=0) throw new Exception("Choose an installation path without redirected folders.");
            Directory.CreateDirectory(root);
            File.WriteAllText(Path.Combine(root,".tng-install"),Marker);
            string versions=Path.Combine(root,"versions"); Directory.CreateDirectory(versions);
            string hash;
            using(var payload=Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
            using(var sha=SHA256.Create()) hash=BitConverter.ToString(sha.ComputeHash(payload)).Replace("-","").ToLowerInvariant().Substring(0,16);
            // Reinstall into a fresh version so a damaged prior copy is repaired
            // without replacing files used by an already-running app.
            hash=hash.Substring(0,8)+Guid.NewGuid().ToString("N").Substring(0,8);
            string dest=Path.Combine(versions,hash);
            if(!Directory.Exists(dest)) {
                stage=Path.Combine(versions,"s-"+Guid.NewGuid().ToString("N").Substring(0,12)); Directory.CreateDirectory(stage);
                using(var payload=Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
                using(var zip=new ZipArchive(payload,ZipArchiveMode.Read)) {
                    foreach(var entry in zip.Entries) {
                        string file=Path.GetFullPath(Path.Combine(stage,entry.FullName.Replace('/',Path.DirectorySeparatorChar)));
                        if(!file.StartsWith(stage+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase) || entry.FullName.Contains(":")) throw new Exception("Invalid installation archive path.");
                        if(entry.FullName.EndsWith("/")) {Directory.CreateDirectory(file); continue;}
                        Directory.CreateDirectory(Path.GetDirectoryName(file));
                        using(var source=entry.Open()) using(var output=new FileStream(file,FileMode.CreateNew,FileAccess.Write)) source.CopyTo(output);
                    }
                }
                Directory.Move(stage,dest); stage=null;
            }
            string launcher=Path.Combine(root,"TNG Expense Agent.exe");
            string tempLauncher=Path.Combine(root,"launcher-new.exe");
            File.Copy(Path.Combine(dest,"TNG Expense Agent.exe"),tempLauncher,true);
            if(File.Exists(launcher)) File.Replace(tempLauncher,launcher,null); else File.Move(tempLauncher,launcher);
            string next=Path.Combine(root,"current-new.txt"); File.WriteAllText(next,hash);
            string current=Path.Combine(root,"current.txt");
            if(File.Exists(current)) File.Replace(next,current,null); else File.Move(next,current);
            if(!test) {
                string shortcut=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs),"TNG Expense Agent.lnk");
                Powershell("$s=New-Object -ComObject WScript.Shell;$l=$s.CreateShortcut("+Quote(shortcut)+");$l.TargetPath="+Quote(launcher)+";$l.WorkingDirectory="+Quote(root)+";$l.Save()");
                MessageBox.Show("Installation complete. Open TNG Expense Agent from the Start menu.\n\nYour data is stored in:\n"+Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"TNGExpenseAgent","data"),"TNG Expense Agent Setup");
                Process.Start(new ProcessStartInfo(launcher){UseShellExecute=true});
            } else Console.WriteLine("Installed: "+root+"; version "+hash);
            return 0;
        } catch(Exception exc) {
            if(test) Console.Error.WriteLine(exc.Message); else MessageBox.Show(exc.Message,"Installation could not finish",MessageBoxButtons.OK,MessageBoxIcon.Error);
            return 1;
        } finally {
            // stage is always a freshly created child of this checked installation.
            if(stage!=null && Directory.Exists(stage)) Directory.Delete(stage,true);
        }
    }
}
