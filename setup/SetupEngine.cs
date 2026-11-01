using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Xml.Linq;
using Microsoft.Win32;

namespace OfficeImageViewer.Setup
{
    public sealed class ProductMetadataReader : MarshalByRefObject
    {
        public string ReadCodeBase(string file) { return Assembly.ReflectionOnlyLoadFrom(file).CodeBase; }
        public string Read(string file)
        {
            Assembly metadata = Assembly.ReflectionOnlyLoadFrom(file);
            AssemblyName identity = metadata.GetName();
            Type type = metadata.GetType("OfficeImageViewer.AddIn", false);
            bool matches = identity.Name == "OfficeImageViewer" && type != null && CustomAttributeData.GetCustomAttributes(type).Any(a => a.Constructor.DeclaringType.FullName == "System.Runtime.InteropServices.GuidAttribute" && a.ConstructorArguments.Count == 1 && String.Equals(Convert.ToString(a.ConstructorArguments[0].Value), ProductRegistry.ProductId.Trim('{', '}'), StringComparison.OrdinalIgnoreCase));
            if (!matches) throw new InvalidOperationException("插件身份无法确认，未作修改。请从原位置卸载。");
            return identity.FullName;
        }
    }
    public sealed class SetupEnvironment
    {
        public string UserDirectory, MachineDirectory;
        public bool Elevated, FrameworkReady;
        public string[] Hosts, RunningHosts;
        public static SetupEnvironment Detect()
        {
            List<string> hosts = new List<string>();
            foreach (string app in new[] { "Word", "PowerPoint" }) using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(app + ".Application")) if (key != null) hosts.Add(app);
            int release = 0;
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")) if (key != null) release = Convert.ToInt32(key.GetValue("Release", 0));
            return new SetupEnvironment {
                UserDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OfficeImageViewer"),
                MachineDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OfficeImageViewer"),
                Elevated = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator),
                FrameworkReady = release >= 528040, Hosts = hosts.ToArray(),
                RunningHosts = new[] { "WINWORD", "POWERPNT" }.Where(n => { Process[] p = Process.GetProcessesByName(n); bool any = p.Length > 0; foreach (Process item in p) item.Dispose(); return any; }).ToArray()
            };
        }
    }
    public sealed class SetupEngine
    {
        public const string Version = "0.2.2";
        static readonly string[] Files = { "OfficeImageViewer.dll", "OfficeImageViewer.Setup.exe", "LICENSE", "installation.xml" };
        readonly IProductRegistry registry;
        readonly Func<SetupEnvironment> environment;
        readonly Dictionary<string, byte[]> payload;
        readonly Action<string> checkpoint;
        public SetupEngine(IProductRegistry registry, Func<SetupEnvironment> environment, Dictionary<string, byte[]> payload, Action<string> checkpoint)
        { this.registry = registry; this.environment = environment; this.payload = payload; this.checkpoint = checkpoint ?? delegate { }; }
        public static Dictionary<string, byte[]> EmbeddedPayload()
        {
            Dictionary<string, byte[]> result = new Dictionary<string, byte[]>();
            Assembly assembly = Assembly.GetExecutingAssembly();
            foreach (string name in new[] { "OfficeImageViewer.dll", "LICENSE" }) using (Stream stream = assembly.GetManifestResourceStream(name))
            { if (stream == null) throw new InvalidOperationException("安装资源不完整，请重新下载完整安装程序。"); using (MemoryStream buffer = new MemoryStream()) { stream.CopyTo(buffer); result[name] = buffer.ToArray(); } }
            result["OfficeImageViewer.Setup.exe"] = File.ReadAllBytes(assembly.Location);
            return result;
        }
        public static string Hash(byte[] bytes) { using (SHA256 sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", ""); }
        public InstallScope SuggestedScope()
        {
            bool user = registry.RegisteredDll(InstallScope.CurrentUser) != null || registry.ManagedDirectory(InstallScope.CurrentUser) != null;
            bool machine = registry.RegisteredDll(InstallScope.AllUsers) != null || registry.ManagedDirectory(InstallScope.AllUsers) != null;
            if (user && machine) throw new InvalidOperationException("发现当前用户和所有用户两种安装。请从原安装分别卸载后重试。");
            return machine ? InstallScope.AllUsers : InstallScope.CurrentUser;
        }
        string Target(SetupEnvironment env, InstallScope scope)
        {
            string path = Path.GetFullPath(scope == InstallScope.AllUsers ? env.MachineDirectory : env.UserDirectory).TrimEnd(Path.DirectorySeparatorChar);
            if (Path.GetFileName(path) != "OfficeImageViewer" || path.StartsWith(@"\\") || path == Path.GetPathRoot(path)) throw new InvalidOperationException("安装目录不安全，未作修改。");
            string ancestor = path;
            while (!String.IsNullOrEmpty(ancestor))
            { if (Directory.Exists(ancestor) && (File.GetAttributes(ancestor) & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("安装路径含链接目录，无法安全写入。"); ancestor = Path.GetDirectoryName(ancestor); }
            return path;
        }
        void Preflight(SetupEnvironment env, InstallScope scope, bool install)
        {
            if (env.RunningHosts.Length > 0) throw new InvalidOperationException("请先保存文档并关闭 " + String.Join("、", env.RunningHosts.Select(n => n == "WINWORD" ? "Word" : "PowerPoint")) + "，然后点击重试。不会自动关闭文档。");
            if (scope == InstallScope.AllUsers && !env.Elevated) throw new UnauthorizedAccessException("所有用户安装需要管理员权限，请使用窗口中的管理员操作。");
            if (install && !env.FrameworkReady) throw new InvalidOperationException("需要 .NET Framework 4.8。请先安装系统组件后重试。");
            if (install && env.Hosts.Length == 0) throw new InvalidOperationException("未检测到桌面版 Word 或 PowerPoint，请先安装 Office 后重试。");
            InstallScope other = scope == InstallScope.CurrentUser ? InstallScope.AllUsers : InstallScope.CurrentUser;
            if (install && (registry.RegisteredDll(other) != null || registry.ManagedDirectory(other) != null)) throw new InvalidOperationException("另一安装范围已有本插件，请保留原范围升级，或先卸载旧版。");
        }
        static AssemblyName ProductIdentity(string file)
        {
            AppDomain domain = AppDomain.CreateDomain("OfficeImageViewer.Metadata", null, new AppDomainSetup { ApplicationBase = Path.GetDirectoryName(typeof(ProductMetadataReader).Assembly.Location) });
            try
            {
                ProductMetadataReader reader = (ProductMetadataReader)domain.CreateInstanceAndUnwrap(typeof(ProductMetadataReader).Assembly.FullName, typeof(ProductMetadataReader).FullName);
                return new AssemblyName(reader.Read(file));
            }
            finally { AppDomain.Unload(domain); }
        }
        public static string GetClrCodeBase(string file)
        {
            AppDomain domain = AppDomain.CreateDomain("OfficeImageViewer.CodeBase", null, new AppDomainSetup { ApplicationBase = Path.GetDirectoryName(typeof(ProductMetadataReader).Assembly.Location) });
            try
            {
                ProductMetadataReader reader = (ProductMetadataReader)domain.CreateInstanceAndUnwrap(typeof(ProductMetadataReader).Assembly.FullName, typeof(ProductMetadataReader).FullName);
                return reader.ReadCodeBase(file);
            }
            finally { AppDomain.Unload(domain); }
        }
        XElement ReadManifest(string directory, InstallScope scope, bool required)
        {
            string path = Path.Combine(directory, "installation.xml");
            if (!File.Exists(path)) { if (required) throw new InvalidOperationException("缺少安装清单，未删除文件。请先用安装程序修复。"); return null; }
            XElement data = XElement.Parse(File.ReadAllText(path));
            if ((string)data.Attribute("product") != ProductRegistry.ProductId || (string)data.Attribute("scope") != scope.ToString() || !String.Equals((string)data.Attribute("directory"), directory, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("安装清单不匹配，未修改文件。");
            System.Version version;
            if (!System.Version.TryParse((string)data.Attribute("version"), out version) || version > new System.Version(Version)) throw new InvalidOperationException("已有较新版本，请使用对应版本的安装程序。");
            foreach (XElement file in data.Elements("file")) if (!Files.Contains((string)file.Attribute("name"))) throw new InvalidOperationException("安装清单含未知文件，未作修改。");
            string[] listed = data.Elements("file").Select(f => (string)f.Attribute("name")).ToArray();
            if (listed.Length != 3 || listed.Distinct().Count() != 3 || !new[] { "OfficeImageViewer.dll", "OfficeImageViewer.Setup.exe", "LICENSE" }.All(listed.Contains)) throw new InvalidOperationException("安装清单不完整，请使用对应版本修复，未作修改。");
            if (data.Elements("file").Any(f => ((string)f.Attribute("sha256") ?? "").Length != 64)) throw new InvalidOperationException("安装文件校验信息不完整，未作修改。");
            return data;
        }
        public void Install(InstallScope scope)
        {
            SetupEnvironment env = environment(); Preflight(env, scope, true);
            string directory = Target(env, scope);
            string managed = registry.ManagedDirectory(scope);
            string originalRegistration = registry.Snapshot(scope);
            if (managed == null && registry.RegisteredDll(scope) == null && XElement.Parse(originalRegistration).Elements("tree").Any()) throw new InvalidOperationException("发现无法识别的残留产品注册，请先从原安装卸载；未覆盖注册。");
            if (managed != null && !String.Equals(Path.GetFullPath(managed).TrimEnd('\\'), directory, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("现有维护目录与固定目录不一致，请从原安装卸载后重试。");
            bool recoveredManifest = false;
            if (managed != null && !File.Exists(Path.Combine(directory, "installation.xml")))
            {
                string registered = registry.RegisteredDll(scope);
                if (registered == null || !String.Equals(Path.GetFullPath(registered), Path.Combine(directory, "OfficeImageViewer.dll"), StringComparison.OrdinalIgnoreCase) || !File.Exists(registered)) throw new InvalidOperationException("安装清单与主程序均不完整，无法安全确认原安装。请保留目录，并通过项目问题反馈提交安装诊断；未修改原文件。");
                ProductIdentity(registered); recoveredManifest = true;
            }
            XElement oldManifest = ReadManifest(directory, scope, managed != null && !recoveredManifest);
            if (Directory.Exists(directory) && oldManifest == null && !recoveredManifest && Directory.EnumerateFileSystemEntries(directory).Any()) throw new InvalidOperationException("安装目录已有未知文件，请移走该目录后重试；未覆盖任何内容。");
            foreach (string name in new[] { "OfficeImageViewer.dll", "OfficeImageViewer.Setup.exe", "LICENSE" }) if (!payload.ContainsKey(name) || payload[name].Length == 0) throw new InvalidOperationException("安装资源缺失，请重新下载。");
            string staging = Path.Combine(Path.GetTempPath(), "OfficeImageViewer-payload-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(staging);
            string identity;
            try { string file = Path.Combine(staging, "OfficeImageViewer.dll"); File.WriteAllBytes(file, payload["OfficeImageViewer.dll"]); identity = ProductIdentity(file).FullName;
                string old = registry.RegisteredDll(scope);
                if (old != null) {
                    if (!File.Exists(old)) { if (oldManifest == null || !String.Equals(Path.GetFullPath(old), Path.Combine(directory, "OfficeImageViewer.dll"), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("旧插件文件丢失，请先从原安装范围卸载后重试。"); }
                    else if (ProductIdentity(old).Version > new AssemblyName(identity).Version) throw new InvalidOperationException("已有较新插件，不能降级。");
                }
            } finally { File.Delete(Path.Combine(staging, "OfficeImageViewer.dll")); Directory.Delete(staging); }
            string snapshot = originalRegistration;
            Dictionary<string, byte[]> before = SnapshotFiles(directory);
            bool existed = Directory.Exists(directory);
            try
            {
                Directory.CreateDirectory(directory);
                foreach (string name in new[] { "OfficeImageViewer.dll", "OfficeImageViewer.Setup.exe", "LICENSE" }) { File.WriteAllBytes(Path.Combine(directory, name), payload[name]); checkpoint("file:" + name); }
                XElement manifest = new XElement("installation", new XAttribute("product", ProductRegistry.ProductId), new XAttribute("version", Version), new XAttribute("scope", scope), new XAttribute("directory", directory));
                foreach (string name in new[] { "OfficeImageViewer.dll", "OfficeImageViewer.Setup.exe", "LICENSE" }) manifest.Add(new XElement("file", new XAttribute("name", name), new XAttribute("sha256", Hash(payload[name]))));
                File.WriteAllText(Path.Combine(directory, "installation.xml"), manifest.ToString()); checkpoint("manifest");
                registry.Register(scope, directory, identity, Version, env.Hosts); checkpoint("registry");
            }
            catch (Exception failure) { Rollback(scope, snapshot, directory, before, existed, failure); throw; }
        }
        static Dictionary<string, byte[]> SnapshotFiles(string directory)
        {
            Dictionary<string, byte[]> before = new Dictionary<string, byte[]>();
            foreach (string name in Files) { string path = Path.Combine(directory, name); if (File.Exists(path)) { if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("安装文件包含链接，未作修改。"); before[name] = File.ReadAllBytes(path); } }
            return before;
        }
        void Rollback(InstallScope scope, string snapshot, string directory, Dictionary<string, byte[]> before, bool existed, Exception failure)
        {
            List<Exception> errors = new List<Exception>();
            try { registry.Restore(scope, snapshot); } catch (Exception error) { errors.Add(error); }
            foreach (string name in Files) try { string path = Path.Combine(directory, name); if (before.ContainsKey(name)) File.WriteAllBytes(path, before[name]); else if (File.Exists(path)) File.Delete(path); } catch (Exception error) { errors.Add(error); }
            try { if (!existed && Directory.Exists(directory) && !Directory.EnumerateFileSystemEntries(directory).Any()) Directory.Delete(directory); } catch (Exception error) { errors.Add(error); }
            if (errors.Count > 0) throw new AggregateException("操作失败，部分恢复未完成。请保留原目录并使用安装程序修复。", new[] { failure }.Concat(errors));
        }
        public void Uninstall(InstallScope scope)
        {
            SetupEnvironment env = environment(); Preflight(env, scope, false);
            string directory = Target(env, scope);
            string managed = registry.ManagedDirectory(scope);
            if (managed == null || !String.Equals(Path.GetFullPath(managed).TrimEnd('\\'), directory, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("未找到匹配的图形安装。脚本安装请先迁移或使用原卸载脚本。");
            XElement manifest = ReadManifest(directory, scope, true);
            string snapshot = registry.Snapshot(scope); Dictionary<string, byte[]> before = SnapshotFiles(directory);
            try
            {
                registry.Remove(scope); checkpoint("remove-registry");
                foreach (XElement file in manifest.Elements("file"))
                {
                    string path = Path.Combine(directory, (string)file.Attribute("name"));
                    if (File.Exists(path) && Hash(File.ReadAllBytes(path)) == (string)file.Attribute("sha256")) File.Delete(path);
                }
                File.Delete(Path.Combine(directory, "installation.xml")); checkpoint("remove-files");
                if (!Directory.EnumerateFileSystemEntries(directory).Any()) Directory.Delete(directory);
            }
            catch (Exception failure) { Rollback(scope, snapshot, directory, before, true, failure); throw; }
        }
        public string Describe(InstallScope scope)
        {
            SetupEnvironment env = environment(); string dir = Target(env, scope);
            string old = registry.RegisteredDll(scope);
            return "安装器版本：" + Version + "\n.NET Framework 4.8：" + (env.FrameworkReady ? "已具备" : "需要安装") +
                "\n桌面 Office：" + (env.Hosts.Length == 0 ? "未检测到" : String.Join("、", env.Hosts)) +
                "\n运行中的 Office：" + (env.RunningHosts.Length == 0 ? "无" : String.Join("、", env.RunningHosts)) +
                "\n安装范围：" + (scope == InstallScope.AllUsers ? "所有用户" : "当前用户") +
                "\n插件状态：" + (old == null ? "尚未安装" : registry.ManagedDirectory(scope) == null ? "旧版脚本安装，可升级" : "已安装，可修复") +
                "\n程序目录：" + dir + "\n\n重新打开 Word/PPT 后双击图片；也可选图后点击“图片查看 → 查看图片”。\n滚轮缩放，拖动平移，1 原始像素，Home 适应窗口，Esc 关闭。\n卸载：Windows 设置 → 应用 → 已安装的应用 → Office 图片查看器。";
        }
    }
}
