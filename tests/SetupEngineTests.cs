using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeImageViewer.Setup;

public static class SetupEngineTests
{
    sealed class MemoryRegistry : IProductRegistry
    {
        public string Dll, Directory, OtherDll, Data = "original";
        public bool Fail;
        public string Snapshot(InstallScope scope) { return new XElement("r", new XAttribute("data", Data), new XAttribute("dll", Dll ?? ""), new XAttribute("dir", Directory ?? "")).ToString(); }
        public void Restore(InstallScope scope, string snapshot) { XElement r=XElement.Parse(snapshot); Data=(string)r.Attribute("data"); Dll=Empty((string)r.Attribute("dll")); Directory=Empty((string)r.Attribute("dir")); }
        static string Empty(string s){return s==""?null:s;}
        public string RegisteredDll(InstallScope scope) { return scope == InstallScope.AllUsers ? OtherDll : Dll; }
        public string ManagedDirectory(InstallScope scope) { return scope == InstallScope.AllUsers ? null : Directory; }
        public void Register(InstallScope scope, string directory, string identity, string version, string[] hosts) { Data="registered:"+version; Directory=directory; Dll=Path.Combine(directory,"OfficeImageViewer.dll"); if(Fail)throw new IOException("Injected registry write failure"); }
        public void Remove(InstallScope scope) { Data="removed"; Directory=null; Dll=null; }
    }
    sealed class Fixture : IDisposable
    {
        public string Root, Target; public MemoryRegistry Registry = new MemoryRegistry();
        public SetupEnvironment Env; public Dictionary<string,byte[]> Payload;
        public Action<string> Failure;
        public Fixture(byte[] dll)
        {
            Root=Path.Combine(Path.GetTempPath(),"OfficeImageViewer-tests-"+Guid.NewGuid().ToString("N")); Target=Path.Combine(Root,"OfficeImageViewer");
            Env=new SetupEnvironment{UserDirectory=Target,MachineDirectory=Path.Combine(Root,"machine","OfficeImageViewer"),Elevated=false,FrameworkReady=true,Hosts=new[]{"Word","PowerPoint"},RunningHosts=new string[0]};
            Payload=new Dictionary<string,byte[]>{{"OfficeImageViewer.dll",dll},{"OfficeImageViewer.Setup.exe",new byte[]{1,2,3}},{"LICENSE",new byte[]{4,5}}};
        }
        public SetupEngine Engine {get{return new SetupEngine(Registry,delegate{return Env;},Payload,delegate(string stage){if(Failure!=null)Failure(stage);});}}
        public void Install(){Engine.Install(InstallScope.CurrentUser);}
        public void Remove(){Engine.Uninstall(InstallScope.CurrentUser);}
        public void Dispose(){ if(Directory.Exists(Root))Directory.Delete(Root,true); }
    }
    static int passed, failed; static byte[] dll;
    public static int Main(string[] args)
    {
        dll=File.ReadAllBytes(args[0]);
        Test("CLR CodeBase preserves space and Unicode path and releases file",delegate(Fixture f){string folder=Path.Combine(f.Root,"space 中文");Directory.CreateDirectory(folder);string file=Path.Combine(folder,"OfficeImageViewer.dll");File.WriteAllBytes(file,dll);string codeBase=SetupEngine.GetClrCodeBase(file);Assert(codeBase.Contains("space 中文")&&!codeBase.Contains("%20"),"CLR unescaped path");File.Delete(file);Assert(!File.Exists(file),"metadata domain releases file");});
        Test("fresh install writes manifest, files and registration",delegate(Fixture f){f.Install();Assert(File.Exists(Path.Combine(f.Target,"installation.xml")),"manifest");Assert(f.Registry.Directory==f.Target,"managed registration");});
        Test("uninstall removes owned files and registration",delegate(Fixture f){f.Install();f.Remove();Assert(!Directory.Exists(f.Target)&&f.Registry.Dll==null,"removed");});
        Test("unknown user files are preserved",delegate(Fixture f){f.Install();File.WriteAllText(Path.Combine(f.Target,"notes.txt"),"keep");f.Remove();Assert(File.ReadAllText(Path.Combine(f.Target,"notes.txt"))=="keep","preserved");});
        Test("modified owned files are preserved",delegate(Fixture f){f.Install();File.WriteAllText(Path.Combine(f.Target,"LICENSE"),"changed");f.Remove();Assert(File.ReadAllText(Path.Combine(f.Target,"LICENSE"))=="changed","preserved");});
        Test("repair restores missing plugin file",delegate(Fixture f){f.Install();File.Delete(Path.Combine(f.Target,"OfficeImageViewer.dll"));f.Install();Assert(File.ReadAllBytes(Path.Combine(f.Target,"OfficeImageViewer.dll")).SequenceEqual(dll),"repaired");});
        Test("missing manifest repair verifies DLL and enables uninstall",delegate(Fixture f){f.Install();File.WriteAllText(Path.Combine(f.Target,"notes.txt"),"keep");File.Delete(Path.Combine(f.Target,"installation.xml"));f.Install();Assert(File.Exists(Path.Combine(f.Target,"installation.xml")),"repaired manifest");f.Remove();Assert(File.ReadAllText(Path.Combine(f.Target,"notes.txt"))=="keep"&&f.Registry.Dll==null,"uninstalled safely");});
        Test("file failure rolls back fresh installation",delegate(Fixture f){f.Failure=delegate(string stage){if(stage=="file:LICENSE")throw new IOException("Injected");};Reject(f.Install);Assert(!Directory.Exists(f.Target)&&f.Registry.Data=="original","rolled back");});
        Test("partial registry write restores old registration and files",delegate(Fixture f){f.Install();string before=f.Registry.Snapshot(InstallScope.CurrentUser);byte[] manifest=File.ReadAllBytes(Path.Combine(f.Target,"installation.xml"));f.Registry.Fail=true;Reject(f.Install);Assert(f.Registry.Snapshot(InstallScope.CurrentUser)==before,"registry restored");Assert(File.ReadAllBytes(Path.Combine(f.Target,"installation.xml")).SequenceEqual(manifest),"files restored");});
        Test("uninstall failure restores registration and files",delegate(Fixture f){f.Install();string before=f.Registry.Snapshot(InstallScope.CurrentUser);f.Failure=delegate(string stage){if(stage=="remove-files")throw new IOException("Injected");};Reject(f.Remove);Assert(f.Registry.Snapshot(InstallScope.CurrentUser)==before&&File.Exists(Path.Combine(f.Target,"OfficeImageViewer.dll")),"restored");});
        Test("running Office blocks before file and registry writes",delegate(Fixture f){f.Env.RunningHosts=new[]{"WINWORD"};Reject(f.Install);Assert(!Directory.Exists(f.Target)&&f.Registry.Data=="original","untouched");});
        Test("missing Office blocks install",delegate(Fixture f){f.Env.Hosts=new string[0];Reject(f.Install);Assert(!Directory.Exists(f.Target),"untouched");});
        Test("missing Framework blocks install",delegate(Fixture f){f.Env.FrameworkReady=false;Reject(f.Install);Assert(!Directory.Exists(f.Target),"untouched");});
        Test("AllUsers requires elevation",delegate(Fixture f){Reject(delegate{f.Engine.Install(InstallScope.AllUsers);});Assert(!Directory.Exists(f.Root),"untouched");});
        Test("other scope registration blocks overwrite",delegate(Fixture f){f.Registry.OtherDll="other.dll";Reject(f.Install);Assert(!Directory.Exists(f.Target),"untouched");});
        Test("unknown target directory is refused",delegate(Fixture f){Directory.CreateDirectory(f.Target);File.WriteAllText(Path.Combine(f.Target,"mine.txt"),"keep");Reject(f.Install);Assert(File.ReadAllText(Path.Combine(f.Target,"mine.txt"))=="keep","untouched");});
        Test("newer manifest refuses downgrade",delegate(Fixture f){f.Install();string file=Path.Combine(f.Target,"installation.xml");XElement data=XElement.Load(file);data.SetAttributeValue("version","99.0.0");data.Save(file);Reject(f.Install);Assert(XElement.Load(file).Attribute("version").Value=="99.0.0","untouched");});
        Test("foreign manifest refuses deletion",delegate(Fixture f){f.Install();string file=Path.Combine(f.Target,"installation.xml");XElement data=XElement.Load(file);data.SetAttributeValue("product","other");data.Save(file);Reject(f.Remove);Assert(File.Exists(Path.Combine(f.Target,"OfficeImageViewer.dll")),"untouched");});
        Test("legacy install migrates registration and keeps original files",delegate(Fixture f){Directory.CreateDirectory(f.Root);string old=Path.Combine(f.Root,"legacy.dll");File.WriteAllBytes(old,dll);f.Registry.Dll=old;f.Install();Assert(File.Exists(old)&&f.Registry.Dll==Path.Combine(f.Target,"OfficeImageViewer.dll"),"migration");});
        Test("unknown legacy DLL is refused",delegate(Fixture f){Directory.CreateDirectory(f.Root);string old=Path.Combine(f.Root,"legacy.dll");File.WriteAllBytes(old,new byte[]{1,2,3});f.Registry.Dll=old;Reject(f.Install);Assert(!Directory.Exists(f.Target)&&f.Registry.Dll==old,"untouched");});
        Test("uninstall running Office keeps installation",delegate(Fixture f){f.Install();f.Env.RunningHosts=new[]{"POWERPNT"};Reject(f.Remove);Assert(File.Exists(Path.Combine(f.Target,"OfficeImageViewer.dll")),"untouched");});
        Test("missing embedded resource is refused",delegate(Fixture f){f.Payload.Remove("LICENSE");Reject(f.Install);Assert(!Directory.Exists(f.Target),"untouched");});
        Console.WriteLine("Setup engine tests: "+passed+" passed, "+failed+" failed (isolated filesystem; simulated registry/environment)");return failed==0?0:1;
    }
    static void Test(string name,Action<Fixture> action){try{using(Fixture f=new Fixture(dll))action(f);passed++;Console.WriteLine("PASS "+name);}catch(Exception e){failed++;Console.WriteLine("FAIL "+name+": "+e.Message);}}
    static void Assert(bool ok,string text){if(!ok)throw new Exception(text);}
    static void Reject(Action action){try{action();}catch{return;}throw new Exception("Expected refusal");}
}
