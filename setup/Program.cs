using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;

[assembly: AssemblyTitle("Office 图片查看器安装程序")]
[assembly: AssemblyVersion("0.2.1.0")]
[assembly: AssemblyFileVersion("0.2.1.0")]

namespace OfficeImageViewer.Setup
{
    public sealed class SetupLock : IDisposable
    {
        readonly Mutex mutex;
        public SetupLock() { mutex = new Mutex(false, @"Global\OfficeImageViewer.Setup.26C7344A"); try { if (!mutex.WaitOne(0)) { mutex.Dispose(); throw new InvalidOperationException("已有安装或维护正在进行，请完成后重试。"); } } catch (AbandonedMutexException) { } }
        public void Dispose() { mutex.ReleaseMutex(); mutex.Dispose(); }
    }
    public static class Program
    {
        [STAThread] public static int Main(string[] args)
        {
            try
            {
                InstallScope? scope = null;
                for (int i = 0; i < args.Length; i++) if (args[i] == "--scope")
                { InstallScope value; if (++i >= args.Length || !Enum.TryParse(args[i], out value) || !Enum.IsDefined(typeof(InstallScope), value)) throw new ArgumentException("无效安装范围。"); scope = value; }
                if (!args.Contains("--worker"))
                {
                    CleanupPreviousWorkers();
                    string temp = Path.Combine(Path.GetTempPath(), "OfficeImageViewer-setup-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(temp);
                    string file = Path.Combine(temp, "OfficeImageViewer.Setup.exe"); File.Copy(Assembly.GetExecutingAssembly().Location, file);
                    File.WriteAllText(Path.Combine(temp, "owner.txt"), ProductRegistry.ProductId);
                    string forwarded = "--worker" + (scope.HasValue ? " --scope " + scope.Value : "") + (args.Contains("--uninstall") ? " --uninstall" : "");
                    Process.Start(new ProcessStartInfo(file, forwarded) { UseShellExecute = false, WorkingDirectory = temp }); return 0;
                }
                SetupEngine engine = new SetupEngine(new ProductRegistry(), SetupEnvironment.Detect, SetupEngine.EmbeddedPayload(), null);
                Application app = new Application(); app.Run(new SetupWindow(engine, scope, args.Contains("--uninstall"))); return 0;
            }
            catch (Exception error) { MessageBox.Show(error.Message, "Office 图片查看器", MessageBoxButton.OK, MessageBoxImage.Warning); return 1; }
        }
        static void CleanupPreviousWorkers()
        {
            foreach (string directory in Directory.GetDirectories(Path.GetTempPath(), "OfficeImageViewer-setup-*"))
            {
                Guid id; string suffix = Path.GetFileName(directory).Substring("OfficeImageViewer-setup-".Length);
                if (!Guid.TryParseExact(suffix, "N", out id) || (File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
                string owner = Path.Combine(directory, "owner.txt"), exe = Path.Combine(directory, "OfficeImageViewer.Setup.exe");
                try
                {
                    if (!File.Exists(owner) || File.ReadAllText(owner) != ProductRegistry.ProductId || Directory.GetFiles(directory).Any(f => f != owner && f != exe) || Directory.GetDirectories(directory).Length > 0) continue;
                    if (File.Exists(exe)) File.Delete(exe); File.Delete(owner); Directory.Delete(directory);
                }
                catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
    }
}
