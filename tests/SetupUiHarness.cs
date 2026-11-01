using System;
using System.IO;
using System.Windows;
using System.Xml.Linq;
using OfficeImageViewer.Setup;

public static class SetupUiHarness
{
    sealed class PreviewRegistry : IProductRegistry
    {
        public string Dll, Directory;
        public string Snapshot(InstallScope scope){return new XElement("r",new XAttribute("dll",Dll??""),new XAttribute("dir",Directory??"")).ToString();}
        public void Restore(InstallScope scope,string data){XElement r=XElement.Parse(data);Dll=(string)r.Attribute("dll");Directory=(string)r.Attribute("dir");if(Dll=="")Dll=null;if(Directory=="")Directory=null;}
        public string RegisteredDll(InstallScope scope){return scope==InstallScope.CurrentUser?Dll:null;}
        public string ManagedDirectory(InstallScope scope){return scope==InstallScope.CurrentUser?Directory:null;}
        public void Register(InstallScope scope,string directory,string identity,string version,string[] hosts){Directory=directory;Dll=Path.Combine(directory,"OfficeImageViewer.dll");}
        public void Remove(InstallScope scope){Directory=null;Dll=null;}
    }
    [STAThread] public static int Main(string[] args)
    {
        string scenario=args.Length>0?args[0]:"install";
        string root=Path.GetFullPath(Path.Combine("work","setup-ui-"+Guid.NewGuid().ToString("N"),"OfficeImageViewer"));
        SetupEnvironment env=new SetupEnvironment{UserDirectory=root,MachineDirectory=root,FrameworkReady=true,Elevated=false,Hosts=new[]{"Word","PowerPoint"},RunningHosts=new string[0]};
        PreviewRegistry registry=new PreviewRegistry();
        SetupEngine engine=new SetupEngine(registry,delegate{return env;},SetupEngine.EmbeddedPayload(),null);
        if(scenario=="maintenance" || scenario=="uninstall")engine.Install(InstallScope.CurrentUser);
        if(scenario=="running")env.RunningHosts=new[]{"WINWORD","POWERPNT"};
        if(scenario=="missing")env.Hosts=new string[0];
        Application app=new Application();SetupWindow window=new SetupWindow(engine,InstallScope.CurrentUser,scenario=="uninstall");
        window.Title+=" · 合成界面验证";
        if(Array.IndexOf(args,"--small")>=0){window.Width=window.MinWidth;window.Height=window.MinHeight;}
        app.Run(window);
        return 0;
    }
}
