using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Win32;

namespace OfficeImageViewer.Setup
{
    public enum InstallScope { CurrentUser, AllUsers }
    public interface IProductRegistry
    {
        string Snapshot(InstallScope scope);
        void Restore(InstallScope scope, string snapshot);
        string RegisteredDll(InstallScope scope);
        string ManagedDirectory(InstallScope scope);
        void Register(InstallScope scope, string directory, string identity, string version, string[] hosts);
        void Remove(InstallScope scope);
    }

    public sealed class ProductRegistry : IProductRegistry
    {
        public const string ProductId = "{26C7344A-8129-4C5B-A462-034F9B5FC712}";
        public const string ProgId = "OfficeImageViewer.AddIn";
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProductId;
        static readonly string[] Roots = { @"Software\Classes\CLSID\" + ProductId,
            @"Software\Classes\" + ProgId, @"Software\Microsoft\Office\Word\Addins\" + ProgId,
            @"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId, UninstallKey };
        static IEnumerable<RegistryView> Views { get { yield return RegistryView.Registry32; if (Environment.Is64BitOperatingSystem) yield return RegistryView.Registry64; } }
        static RegistryKey Open(InstallScope scope, RegistryView view) { return RegistryKey.OpenBaseKey(scope == InstallScope.AllUsers ? RegistryHive.LocalMachine : RegistryHive.CurrentUser, view); }
        public string Snapshot(InstallScope scope)
        {
            XElement data = new XElement("registry");
            foreach (RegistryView view in Views) using (RegistryKey root = Open(scope, view))
                foreach (string path in Roots) using (RegistryKey key = root.OpenSubKey(path))
                    if (key != null) data.Add(new XElement("tree", new XAttribute("view", view), new XAttribute("path", path), ReadTree(key)));
            return data.ToString(SaveOptions.DisableFormatting);
        }
        static XElement ReadTree(RegistryKey key)
        {
            XElement data = new XElement("key");
            foreach (string name in key.GetValueNames())
            {
                RegistryValueKind kind = key.GetValueKind(name);
                object value = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                XElement item = new XElement("value", new XAttribute("name", name), new XAttribute("kind", kind));
                if (value is byte[]) item.Value = Convert.ToBase64String((byte[])value);
                else if (value is string[]) foreach (string entry in (string[])value) item.Add(new XElement("item", entry));
                else item.Value = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                data.Add(item);
            }
            foreach (string name in key.GetSubKeyNames()) using (RegistryKey child = key.OpenSubKey(name))
                data.Add(new XElement("child", new XAttribute("name", name), ReadTree(child)));
            return data;
        }
        static void WriteTree(RegistryKey key, XElement data)
        {
            foreach (XElement entry in data.Elements("value"))
            {
                RegistryValueKind kind = (RegistryValueKind)Enum.Parse(typeof(RegistryValueKind), (string)entry.Attribute("kind"));
                object value = entry.Value;
                if (kind == RegistryValueKind.Binary || kind == RegistryValueKind.None) value = Convert.FromBase64String(entry.Value);
                else if (kind == RegistryValueKind.MultiString) value = entry.Elements("item").Select(e => e.Value).ToArray();
                else if (kind == RegistryValueKind.DWord) value = Int32.Parse(entry.Value, System.Globalization.CultureInfo.InvariantCulture);
                else if (kind == RegistryValueKind.QWord) value = Int64.Parse(entry.Value, System.Globalization.CultureInfo.InvariantCulture);
                key.SetValue((string)entry.Attribute("name"), value, kind);
            }
            foreach (XElement child in data.Elements("child")) using (RegistryKey target = key.CreateSubKey((string)child.Attribute("name"))) WriteTree(target, child.Element("key"));
        }
        public void Restore(InstallScope scope, string snapshot)
        {
            XElement data = XElement.Parse(snapshot);
            foreach (XElement tree in data.Elements("tree"))
                if (!Roots.Contains((string)tree.Attribute("path"))) throw new InvalidOperationException("注册快照不属于本产品。");
            Remove(scope);
            foreach (XElement tree in data.Elements("tree")) using (RegistryKey root = Open(scope, (RegistryView)Enum.Parse(typeof(RegistryView), (string)tree.Attribute("view"))))
                using (RegistryKey key = root.CreateSubKey((string)tree.Attribute("path"))) WriteTree(key, tree.Element("key"));
        }
        string ReadConsistent(InstallScope scope, string path, string name)
        {
            string found = null;
            foreach (RegistryView view in Views) using (RegistryKey root = Open(scope, view)) using (RegistryKey key = root.OpenSubKey(path))
            {
                string value = key == null ? null : key.GetValue(name) as string;
                if (value == null) continue;
                if (found != null && !String.Equals(found, value, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("本插件的32/64位注册位置不一致，请先修复原安装。");
                found = value;
            }
            return found;
        }
        public string RegisteredDll(InstallScope scope)
        {
            string value = ReadConsistent(scope, Roots[0] + @"\InprocServer32", "CodeBase");
            if (value == null) return null;
            Uri uri;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri) || !uri.IsFile || uri.IsUnc) throw new InvalidOperationException("旧注册不是本地插件路径，未作修改。");
            return uri.LocalPath;
        }
        public string ManagedDirectory(InstallScope scope) { return ReadConsistent(scope, UninstallKey, "InstallLocation"); }
        static void Set(RegistryKey root, string path, string name, object value)
        { using (RegistryKey key = root.CreateSubKey(path)) key.SetValue(name, value, value is int ? RegistryValueKind.DWord : RegistryValueKind.String); }
        public void Register(InstallScope scope, string directory, string identity, string version, string[] hosts)
        {
            // Replace only these product-owned trees, removing stale version registrations.
            Remove(scope);
            string dll = SetupEngine.GetClrCodeBase(System.IO.Path.Combine(directory, "OfficeImageViewer.dll"));
            foreach (RegistryView view in Views) using (RegistryKey root = Open(scope, view))
            {
                Set(root, Roots[1], "", "Office Image Viewer"); Set(root, Roots[1] + @"\CLSID", "", ProductId);
                Set(root, Roots[0], "", ProgId); Set(root, Roots[0] + @"\ProgId", "", ProgId);
                string server = Roots[0] + @"\InprocServer32";
                Set(root, server, "", "mscoree.dll"); Set(root, server, "ThreadingModel", "Both");
                foreach (string path in new[] { server, server + @"\" + new System.Reflection.AssemblyName(identity).Version })
                {
                    Set(root, path, "Class", ProgId); Set(root, path, "Assembly", identity);
                    Set(root, path, "RuntimeVersion", "v4.0.30319"); Set(root, path, "CodeBase", dll);
                }
                using (RegistryKey category = root.CreateSubKey(Roots[0] + @"\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")) { }
                foreach (string host in hosts)
                {
                    string path = @"Software\Microsoft\Office\" + host + @"\Addins\" + ProgId;
                    Set(root, path, "FriendlyName", "Office 图片查看器"); Set(root, path, "Description", "双击 Word / PowerPoint 图片放大查看"); Set(root, path, "LoadBehavior", 3);
                }
            }
            // One entry per scope; COM registration still covers both Office bitnesses.
            RegistryView arpView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey root = Open(scope, arpView))
            {
                string exe = System.IO.Path.Combine(directory, "OfficeImageViewer.Setup.exe");
                string args = " --scope " + scope;
                Set(root, UninstallKey, "DisplayName", "Office 图片查看器"); Set(root, UninstallKey, "DisplayVersion", version);
                Set(root, UninstallKey, "Publisher", "OfficeImageViewer contributors"); Set(root, UninstallKey, "InstallLocation", directory);
                Set(root, UninstallKey, "UninstallString", "\"" + exe + "\" --uninstall" + args);
                Set(root, UninstallKey, "ModifyPath", "\"" + exe + "\"" + args);
                Set(root, UninstallKey, "DisplayIcon", exe); Set(root, UninstallKey, "NoRepair", 0);
                Set(root, UninstallKey, "EstimatedSize", 1024); Set(root, UninstallKey, "InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            }
        }
        public void Remove(InstallScope scope)
        { foreach (RegistryView view in Views) using (RegistryKey root = Open(scope, view)) foreach (string path in Roots) root.DeleteSubKeyTree(path, false); }
    }
}
